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
  # The model is called through a cross-region inference profile ("us." + model
  # id). Bedrock checks the profile itself and the model in each region the
  # profile can route the request to, so both need granting.
  inference_profile_arn = "arn:aws:bedrock:${var.bedrock_region}:${data.aws_caller_identity.current.account_id}:inference-profile/${var.bedrock_model_id}"
  foundation_model_arn  = "arn:aws:bedrock:*::foundation-model/${replace(var.bedrock_model_id, "/^[a-z]+[.]/", "")}"

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

  # Image descriptions go through the ordinary Bedrock endpoint (InvokeModel),
  # which this account is entitled to, not the newer Messages API endpoint
  # (bedrock-mantle), which it is not. Granting bedrock-mantle:CreateInference
  # alone left every call denied and every picture with placeholder alt text.
  statement {
    sid    = "DescribeImagesWithClaude"
    effect = "Allow"

    actions = ["bedrock:InvokeModel"]

    resources = [local.inference_profile_arn]
  }

  # the model itself, in whichever US region the profile sends the request to,
  # but only when the call comes through this profile
  statement {
    sid    = "DescribeImagesWithClaudeModel"
    effect = "Allow"

    actions = ["bedrock:InvokeModel"]

    resources = [local.foundation_model_arn]

    condition {
      test     = "StringEquals"
      variable = "bedrock:InferenceProfileArn"
      values   = [local.inference_profile_arn]
    }
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
      # is its own variable. It is where the inference profile lives, which
      # need not be the region the function runs in.
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

  lambda_function {
    lambda_function_arn = aws_lambda_function.this.arn
    events              = ["s3:ObjectCreated:*"]
    filter_suffix       = ".pdf"
  }

  depends_on = [aws_lambda_permission.allow_uploads_bucket]
}
