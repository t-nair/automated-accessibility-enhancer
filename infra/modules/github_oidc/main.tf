# The role GitHub Actions assumes to run terraform plan.
#
# The point of this module is that CI never holds an AWS access key. GitHub
# mints a short-lived token for the workflow run, AWS trades it for temporary
# credentials, and nothing long-lived is ever stored in the repository.
#
# The role can plan, not apply: it gets read access plus the writes the S3
# backend needs to take and release its lock. An apply from CI would need its
# own role with the permissions to match, added deliberately.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

locals {
  oidc_hostname = "token.actions.githubusercontent.com"

  # an account can only hold one provider per URL, so an account that already
  # uses GitHub Actions elsewhere has to reuse the existing one
  provider_arn = var.create_oidc_provider ? aws_iam_openid_connect_provider.github[0].arn : data.aws_iam_openid_connect_provider.github[0].arn

  # "repo:owner/name:*" trusts every workflow in the repository. Narrow it to
  # something like "repo:owner/name:ref:refs/heads/main" to trust only one branch.
  allowed_subjects = coalesce(var.allowed_subjects, ["repo:${var.github_repository}:*"])
}

resource "aws_iam_openid_connect_provider" "github" {
  count = var.create_oidc_provider ? 1 : 0

  url = "https://${local.oidc_hostname}"

  client_id_list = ["sts.amazonaws.com"]

  # no thumbprint_list: AWS verifies GitHub's certificate against its own trust
  # store now, and a pinned thumbprint is one more thing to rotate
}

data "aws_iam_openid_connect_provider" "github" {
  count = var.create_oidc_provider ? 0 : 1

  url = "https://${local.oidc_hostname}"
}

data "aws_iam_policy_document" "assume_role" {
  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRoleWithWebIdentity"]

    principals {
      type        = "Federated"
      identifiers = [local.provider_arn]
    }

    # without the audience check, a token minted for a different AWS-facing
    # purpose would be accepted here
    condition {
      test     = "StringEquals"
      variable = "${local.oidc_hostname}:aud"
      values   = ["sts.amazonaws.com"]
    }

    # and without the subject check, any repository on GitHub could assume it
    condition {
      test     = "StringLike"
      variable = "${local.oidc_hostname}:sub"
      values   = local.allowed_subjects
    }
  }
}

resource "aws_iam_role" "plan" {
  name               = var.role_name
  description        = "Assumed by GitHub Actions to run terraform plan. Read access plus the state lock."
  assume_role_policy = data.aws_iam_policy_document.assume_role.json

  max_session_duration = var.max_session_duration
}

# a plan reads the current state of everything the configuration touches, which
# spans S3, ECR, Lambda, IAM, CloudWatch and SNS
resource "aws_iam_role_policy_attachment" "read_only" {
  role       = aws_iam_role.plan.name
  policy_arn = "arn:aws:iam::aws:policy/ReadOnlyAccess"
}

data "aws_iam_policy_document" "state_access" {
  statement {
    sid    = "FindStateFile"
    effect = "Allow"

    actions = ["s3:ListBucket"]

    resources = ["arn:aws:s3:::${var.state_bucket_name}"]
  }

  # a plan is not read-only against the backend: it writes a lock file while it
  # runs and deletes it afterwards, so it needs more than GetObject
  statement {
    sid    = "ReadStateAndHoldTheLock"
    effect = "Allow"

    actions = [
      "s3:GetObject",
      "s3:PutObject",
      "s3:DeleteObject",
    ]

    resources = ["arn:aws:s3:::${var.state_bucket_name}/*"]
  }

  # ReadOnlyAccess would otherwise let a workflow download real course material
  # out of the deck buckets. A plan has no reason to read an object out of
  # them, and an explicit deny beats the allow in the managed policy.
  dynamic "statement" {
    for_each = length(var.deny_object_read_bucket_arns) > 0 ? [1] : []

    content {
      sid    = "NeverReadUploadedDecks"
      effect = "Deny"

      actions = ["s3:GetObject"]

      resources = [for arn in var.deny_object_read_bucket_arns : "${arn}/*"]
    }
  }
}

resource "aws_iam_role_policy" "state_access" {
  name   = "${var.role_name}-state"
  role   = aws_iam_role.plan.id
  policy = data.aws_iam_policy_document.state_access.json
}
