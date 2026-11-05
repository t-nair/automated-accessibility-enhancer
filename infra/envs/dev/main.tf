# The dev environment: one of everything, wired together.
#
# Copy this directory to make a second environment. The only things that have to
# change are the backend key in backend.tf and the environment name.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }

    random = {
      source  = "hashicorp/random"
      version = "~> 3.6"
    }
  }
}

provider "aws" {
  region = var.aws_region

  default_tags {
    tags = {
      Project     = var.project
      Environment = var.environment
      ManagedBy   = "terraform"
    }
  }
}

locals {
  name_prefix = "${var.project}-${var.environment}"

  # named here rather than taken from the lambda module's output, so that the
  # observability module can create the log group and the lambda module can
  # scope its logging permission to it without the two depending on each other
  function_name = "${local.name_prefix}-pipeline"
}

module "buckets" {
  source = "../../modules/s3_buckets"

  name_prefix               = local.name_prefix
  uploads_expiration_days   = var.uploads_expiration_days
  processed_expiration_days = var.processed_expiration_days
}

module "ecr" {
  source = "../../modules/ecr"

  repository_name = "${local.name_prefix}-pipeline"
}

# a second repository, because the website and the pipeline are different images:
# the pipeline carries LibreOffice and never serves a page, the website serves
# pages and never opens a deck
module "ecr_web" {
  source = "../../modules/ecr"

  repository_name = "${local.name_prefix}-web"
}

module "submissions" {
  source = "../../modules/dynamodb"

  table_name             = "${local.name_prefix}-submissions"
  point_in_time_recovery = var.submissions_point_in_time_recovery
}

# The key that signs the session cookie. It is generated here rather than typed
# in, and kept out of the Terraform outputs, so it exists in the state file and
# in Parameter Store and nowhere else.
#
# Changing it signs everyone out, in the sense that every visitor loses the list
# of files they have sent in. There are no accounts, so that is the whole of what
# a session holds.
resource "random_password" "session_key" {
  length  = 64
  special = false
}

resource "aws_ssm_parameter" "session_key" {
  name        = "/${var.project}/${var.environment}/session-key"
  description = "Signs the website's session cookie. Every instance has to use the same one."
  type        = "SecureString"
  value       = random_password.session_key.result

  lifecycle {
    # so a hand-rotated key is not put back on the next apply
    ignore_changes = [value]
  }
}

module "observability" {
  source = "../../modules/observability"

  function_name         = local.function_name
  environment           = var.environment
  log_retention_days    = var.log_retention_days
  timeout_seconds       = var.lambda_timeout_seconds
  alarm_email_addresses = var.alarm_email_addresses
}

module "pipeline" {
  source = "../../modules/lambda"

  function_name = local.function_name
  image_uri     = "${module.ecr.repository_url}:${var.image_tag}"
  architecture  = var.lambda_architecture

  uploads_bucket_name   = module.buckets.uploads_bucket_name
  uploads_bucket_arn    = module.buckets.uploads_bucket_arn
  processed_bucket_name = module.buckets.processed_bucket_name
  processed_bucket_arn  = module.buckets.processed_bucket_arn

  submissions_table_name = module.submissions.table_name
  submissions_table_arn  = module.submissions.table_arn

  bedrock_model_id = var.bedrock_model_id
  bedrock_region   = var.bedrock_region

  timeout_seconds      = var.lambda_timeout_seconds
  memory_mb            = var.lambda_memory_mb
  reserved_concurrency = var.lambda_reserved_concurrency

  # so the log group exists with its retention already set, rather than Lambda
  # creating one that keeps logs forever on the first invocation
  depends_on = [module.observability]
}

module "website" {
  source = "../../modules/apprunner"

  service_name = "${local.name_prefix}-web"
  image_uri    = "${module.ecr_web.repository_url}:${var.image_tag}"

  uploads_bucket_name   = module.buckets.uploads_bucket_name
  uploads_bucket_arn    = module.buckets.uploads_bucket_arn
  processed_bucket_name = module.buckets.processed_bucket_name
  processed_bucket_arn  = module.buckets.processed_bucket_arn

  submissions_table_name = module.submissions.table_name
  submissions_table_arn  = module.submissions.table_arn
  submissions_index_arn  = module.submissions.index_arn

  secret_key_parameter_name = aws_ssm_parameter.session_key.name
  secret_key_parameter_arn  = aws_ssm_parameter.session_key.arn

  # the website decides a run has died once this has passed, so it has to agree
  # with what the pipeline is actually allowed
  pipeline_timeout_seconds = var.lambda_timeout_seconds

  cpu            = var.web_cpu
  memory         = var.web_memory
  min_instances  = var.web_min_instances
  max_instances  = var.web_max_instances
  deploy_on_push = var.web_deploy_on_push
}

module "ci_role" {
  source = "../../modules/github_oidc"

  github_repository    = var.github_repository
  create_oidc_provider = var.create_oidc_provider
  role_name            = "${local.name_prefix}-github-plan"
  state_bucket_name    = var.state_bucket_name

  # CI plans infrastructure; it has no business reading the decks themselves
  deny_object_read_bucket_arns = [
    module.buckets.uploads_bucket_arn,
    module.buckets.processed_bucket_arn,
  ]
}
