# The function that runs the pipeline, its execution role, and the S3
# notification that starts it.
#
# The role is deliberately narrow: read the deck that arrived, write the fixed
# copy and report back, write its own logs, and ask Claude on Bedrock to
# describe a picture. Nothing else.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

data "aws_caller_identity" "current" {}
data "aws_region" "current" {}

locals {
  log_group_arn = "arn:aws:logs:${data.aws_region.current.region}:${data.aws_caller_identity.current.account_id}:log-group:/aws/lambda/${var.function_name}"
}

data "aws_iam_policy_document" "assume_role" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRole"]

    principals {
      type        = "Service"
      identifiers = ["lambda.amazonaws.com"]
    }
  }
}

resource "aws_iam_role" "this" {
  name               = "${var.function_name}-role"
  assume_role_policy = data.aws_iam_policy_document.assume_role.json
}

data "aws_iam_policy_document" "this" {
  # the managed AWSLambdaBasicExecutionRole policy grants this on every log
  # group in the account, so it is written out here scoped to this one
  statement {
    sid    = "WriteOwnLogs"
    effect = "Allow"

    actions = [
      "logs:CreateLogStream",
      "logs:PutLogEvents",
    ]

    resources = ["${local.log_group_arn}:*"]
  }

  statement {
    sid    = "ReadUploadedDecks"
    effect = "Allow"

    actions = ["s3:GetObject"]

    resources = ["${var.uploads_bucket_arn}/*"]
  }

  statement {
    sid    = "WriteProcessedDecks"
    effect = "Allow"

    actions = ["s3:PutObject"]

    resources = ["${var.processed_bucket_arn}/*"]
  }

  # the pipeline says how far along it is on the submission's own row, which is
  # where the website's progress bar reads from. It only ever updates a row the
  # website has already created, so it cannot put one there or take one away.
  statement {
    sid    = "ReportProgress"
    effect = "Allow"

    actions = ["dynamodb:UpdateItem"]

    resources = [var.submissions_table_arn]
  }

  # Captioning goes through the Messages API endpoint on Bedrock
  # (bedrock-mantle), whose inference call authorizes
  # bedrock-mantle:CreateInference. That is a different action from
  # bedrock:InvokeModel, which belongs to the older InvokeModel and Converse
  # APIs and will not authorize this.
  statement {
    sid    = "DescribeImagesWithClaude"
    effect = "Allow"

    actions = ["bedrock-mantle:CreateInference"]

    resources = var.bedrock_model_arns
  }
}

resource "aws_iam_role_policy" "this" {
  name   = "${var.function_name}-policy"
  role   = aws_iam_role.this.id
  policy = data.aws_iam_policy_document.this.json
}

resource "aws_lambda_function" "this" {
  function_name = var.function_name
  role          = aws_iam_role.this.arn

  package_type  = "Image"
  image_uri     = var.image_uri
  architectures = [var.architecture]

  # a deck is one Bedrock call per undescribed picture, so a large lecture can
  # legitimately take minutes
  timeout = var.timeout_seconds

  # Lambda scales CPU with memory, and re-saving a large deck is CPU-bound, so
  # this buys speed as much as it buys headroom
  memory_size = var.memory_mb

  ephemeral_storage {
    size = var.ephemeral_storage_mb
  }

  # a ceiling on how many decks can be in flight at once. It keeps a bulk
  # upload inside the Bedrock rate limit, and caps what a runaway loop could
  # spend before anyone notices.
  reserved_concurrent_executions = var.reserved_concurrency

  environment {
    variables = {
      PROCESSED_BUCKET  = var.processed_bucket_name
      SUBMISSIONS_TABLE = var.submissions_table_name
      BEDROCK_MODEL_ID  = var.bedrock_model_id

      # AWS_REGION is reserved and set by Lambda itself, so the Bedrock region
      # is its own variable. It matters because bedrock-mantle is served in a
      # subset of regions, which need not include the one the function runs in.
      BEDROCK_REGION = var.bedrock_region

      LOG_LEVEL = var.log_level
    }
  }

  tags = var.tags
}

resource "aws_lambda_permission" "allow_uploads_bucket" {
  statement_id  = "AllowExecutionFromUploadsBucket"
  action        = "lambda:InvokeFunction"
  function_name = aws_lambda_function.this.function_name
  principal     = "s3.amazonaws.com"
  source_arn    = var.uploads_bucket_arn

  # without this, a bucket in someone else's account could invoke this function
  source_account = data.aws_caller_identity.current.account_id
}

# S3 filters by suffix rather than the function waking up and checking, so a
# stray file dropped in the bucket never starts a run at all
resource "aws_s3_bucket_notification" "uploads" {
  bucket = var.uploads_bucket_name

  lambda_function {
    lambda_function_arn = aws_lambda_function.this.arn
    events              = ["s3:ObjectCreated:*"]
    filter_suffix       = ".pptx"
  }

  lambda_function {
    lambda_function_arn = aws_lambda_function.this.arn
    events              = ["s3:ObjectCreated:*"]
    filter_suffix       = ".ppt"
  }

  depends_on = [aws_lambda_permission.allow_uploads_bucket]
}
