# The website, on ECS Express Mode.
#
# This used to be App Runner, which AWS stopped opening to new customers on
# 30 April 2026. Express Mode is what AWS points App Runner users to, and it keeps
# what App Runner was chosen for: it takes a container image and gives back an
# HTTPS address. ECS creates and looks after the load balancer, certificate,
# security group and scaling itself (with the infrastructure role below), so none
# of that has to be written or understood here.
#
# The website never processes a deck. Lambda does that, so the longest request
# this service handles is the upload itself.

terraform {
  required_version = ">= 1.11"

  required_providers {
    # aws_ecs_express_gateway_service arrived in 6.23
    aws = {
      source  = "hashicorp/aws"
      version = ">= 6.23, < 7.0"
    }
  }
}

data "aws_iam_policy_document" "assume_ecs_tasks" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]

    principals {
      type        = "Service"
      identifiers = ["ecs-tasks.amazonaws.com"]
    }
  }
}

# --- the role the running website gets ---

resource "aws_iam_role" "task" {
  name               = "${var.service_name}-task"
  assume_role_policy = data.aws_iam_policy_document.assume_ecs_tasks.json
}

# What the website is allowed to do, and nothing more. Note there is no Bedrock
# permission here: the website never describes a picture, the pipeline does.
data "aws_iam_policy_document" "task" {
  statement {
    sid    = "StoreUploadedDecks"
    effect = "Allow"

    # DeleteObject is here because a visitor can delete their own submission,
    # which has to remove the upload as well as the row
    actions = [
      "s3:PutObject",
      "s3:GetObject",
      "s3:DeleteObject",
    ]

    resources = ["${var.uploads_bucket_arn}/*"]
  }

  statement {
    sid    = "ReadAndClearResults"
    effect = "Allow"

    actions = [
      "s3:GetObject",
      "s3:DeleteObject",
    ]

    resources = ["${var.processed_bucket_arn}/*"]
  }

  statement {
    sid    = "KeepTrackOfSubmissions"
    effect = "Allow"

    actions = [
      "dynamodb:PutItem",
      "dynamodb:GetItem",
      "dynamodb:UpdateItem",
      "dynamodb:DeleteItem",
      "dynamodb:Query",
    ]

    resources = [
      var.submissions_table_arn,
      var.submissions_index_arn,
    ]
  }

  # the key that signs the session cookie. Every task has to sign with the same
  # one or a visitor loses their list of files as requests move between tasks,
  # which is why it is not made up locally.
  statement {
    sid    = "ReadTheCookieKey"
    effect = "Allow"

    actions = ["ssm:GetParameter"]

    resources = [var.secret_key_parameter_arn]
  }
}

resource "aws_iam_role_policy" "task" {
  name   = "${var.service_name}-permissions"
  role   = aws_iam_role.task.id
  policy = data.aws_iam_policy_document.task.json
}

# --- the role ECS uses to start the container: pull the image, write logs ---

resource "aws_iam_role" "execution" {
  name               = "${var.service_name}-execution"
  assume_role_policy = data.aws_iam_policy_document.assume_ecs_tasks.json
}

resource "aws_iam_role_policy_attachment" "execution" {
  role       = aws_iam_role.execution.name
  policy_arn = "arn:aws:iam::aws:policy/service-role/AmazonECSTaskExecutionRolePolicy"
}

# --- the role ECS uses to build the load balancer, certificate and scaling ---

data "aws_iam_policy_document" "assume_ecs" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]

    principals {
      type        = "Service"
      identifiers = ["ecs.amazonaws.com"]
    }
  }
}

resource "aws_iam_role" "infrastructure" {
  name               = "${var.service_name}-infrastructure"
  assume_role_policy = data.aws_iam_policy_document.assume_ecs.json
}

# AWS's own policy for this. It only lets ECS touch load balancers, security
# groups and certificates that ECS itself tagged as its own.
resource "aws_iam_role_policy_attachment" "infrastructure" {
  role       = aws_iam_role.infrastructure.name
  policy_arn = "arn:aws:iam::aws:policy/service-role/AmazonECSInfrastructureRoleforExpressGatewayServices"
}

# --- logs ---

# made here so it has a retention period, rather than keeping logs forever
resource "aws_cloudwatch_log_group" "this" {
  name              = "/aws/ecs/${var.service_name}"
  retention_in_days = var.log_retention_days
}

# --- the service ---

# its own cluster rather than "default", which a fresh account does not have
resource "aws_ecs_cluster" "this" {
  name = var.service_name
}

resource "aws_ecs_express_gateway_service" "this" {
  service_name            = var.service_name
  cluster                 = aws_ecs_cluster.this.name
  execution_role_arn      = aws_iam_role.execution.arn
  infrastructure_role_arn = aws_iam_role.infrastructure.arn
  task_role_arn           = aws_iam_role.task.arn

  cpu               = tostring(var.cpu)
  memory            = tostring(var.memory)
  health_check_path = "/"

  # so an apply (and the deploy job that runs it) fails when the new image does
  # not come up, instead of reporting success and leaving the old one running
  wait_for_steady_state = true

  primary_container {
    image          = var.image_uri
    container_port = 8080

    # the API requires the prefix, although the provider docs call it optional
    aws_logs_configuration {
      log_group         = aws_cloudwatch_log_group.this.name
      log_stream_prefix = "web"
    }

    dynamic "environment" {
      for_each = {
        UPLOADS_BUCKET           = var.uploads_bucket_name
        PROCESSED_BUCKET         = var.processed_bucket_name
        SUBMISSIONS_TABLE        = var.submissions_table_name
        SECRET_KEY_PARAMETER     = var.secret_key_parameter_name
        PIPELINE_TIMEOUT_SECONDS = tostring(var.pipeline_timeout_seconds)
        DEBUG                    = "false"
      }

      content {
        name  = environment.key
        value = environment.value
      }
    }
  }

  scaling_target {
    auto_scaling_metric       = "AVERAGE_CPU"
    auto_scaling_target_value = 60
    min_task_count            = var.min_tasks
    max_task_count            = var.max_tasks
  }

  # Left out, ECS uses the account's default VPC and its default security group.
  # Accounts without a default VPC (some university ones) pass two or more
  # subnets here.
  dynamic "network_configuration" {
    for_each = var.subnet_ids == null ? [] : [1]

    content {
      subnets         = var.subnet_ids
      security_groups = var.security_group_ids
    }
  }

  # Without these the roles can be deleted first on destroy, and the service is
  # left stuck draining because ECS can no longer tear its load balancer down.
  depends_on = [
    aws_iam_role_policy.task,
    aws_iam_role_policy_attachment.execution,
    aws_iam_role_policy_attachment.infrastructure,
  ]
}
