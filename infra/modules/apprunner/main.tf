# The website, on App Runner.
#
# App Runner rather than ECS on Fargate because Fargate needs a VPC, subnets,
# security groups, a load balancer, a target group, a listener, a cluster, a task
# definition and a service before a page can be served. App Runner takes a
# container image and gives back an HTTPS address. This project is handed on to
# someone else at the end of the internship, so there is a real cost to every
# resource left behind for them to understand.
#
# The tradeoff is App Runner's 120 second request limit. That is fine here
# because the deck is processed by Lambda, not by this service: the longest
# request the website serves is the upload itself.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

# App Runner pulls the image as itself, so it needs its own role to read ECR,
# separate from the role the running container gets
data "aws_iam_policy_document" "assume_build" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]

    principals {
      type        = "Service"
      identifiers = ["build.apprunner.amazonaws.com"]
    }
  }
}

resource "aws_iam_role" "build" {
  name               = "${var.service_name}-ecr-access"
  assume_role_policy = data.aws_iam_policy_document.assume_build.json
}

resource "aws_iam_role_policy_attachment" "build" {
  role       = aws_iam_role.build.name
  policy_arn = "arn:aws:iam::aws:policy/service-role/AWSAppRunnerServicePolicyForECRAccess"
}

data "aws_iam_policy_document" "assume_instance" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]

    principals {
      type        = "Service"
      identifiers = ["tasks.apprunner.amazonaws.com"]
    }
  }
}

resource "aws_iam_role" "instance" {
  name               = "${var.service_name}-instance"
  assume_role_policy = data.aws_iam_policy_document.assume_instance.json
}

# What the website is allowed to do, and nothing more. Note there is no Bedrock
# permission here: the website never describes a picture, the pipeline does.
data "aws_iam_policy_document" "instance" {
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

  # the key that signs the session cookie. Every instance has to sign with the
  # same one or a visitor loses their list of files as they are moved between
  # instances, which is why it is not made up locally.
  statement {
    sid    = "ReadTheCookieKey"
    effect = "Allow"

    actions = ["ssm:GetParameter"]

    resources = [var.secret_key_parameter_arn]
  }
}

resource "aws_iam_role_policy" "instance" {
  name   = "${var.service_name}-permissions"
  role   = aws_iam_role.instance.id
  policy = data.aws_iam_policy_document.instance.json
}

resource "aws_apprunner_auto_scaling_configuration_version" "this" {
  auto_scaling_configuration_name = var.service_name

  min_size        = var.min_instances
  max_size        = var.max_instances
  max_concurrency = var.max_requests_per_instance
}

resource "aws_apprunner_service" "this" {
  service_name = var.service_name

  source_configuration {
    # a new image with the same tag is picked up on its own, so a deploy is a
    # docker push rather than a terraform apply
    auto_deployments_enabled = var.deploy_on_push

    authentication_configuration {
      access_role_arn = aws_iam_role.build.arn
    }

    image_repository {
      image_identifier      = var.image_uri
      image_repository_type = "ECR"

      image_configuration {
        port = "8080"

        runtime_environment_variables = {
          UPLOADS_BUCKET           = var.uploads_bucket_name
          PROCESSED_BUCKET         = var.processed_bucket_name
          SUBMISSIONS_TABLE        = var.submissions_table_name
          SECRET_KEY_PARAMETER     = var.secret_key_parameter_name
          PIPELINE_TIMEOUT_SECONDS = tostring(var.pipeline_timeout_seconds)
          DEBUG                    = "false"
        }
      }
    }
  }

  instance_configuration {
    cpu               = var.cpu
    memory            = var.memory
    instance_role_arn = aws_iam_role.instance.arn
  }

  health_check_configuration {
    protocol = "HTTP"
    path     = "/"
    interval = 10
    timeout  = 5
  }

  auto_scaling_configuration_arn = aws_apprunner_auto_scaling_configuration_version.this.arn
}
