# The roles GitHub Actions assumes to run terraform plan and terraform apply.
#
# The point of this module is that CI never holds an AWS access key. GitHub
# mints a short-lived token for the workflow run, AWS trades it for temporary
# credentials, and nothing long-lived is ever stored in the repository.
#
# The plan role can plan, not apply: it gets read access plus the writes the S3
# backend needs to take and release its lock. Applying from CI uses the separate
# apply role further down, which only one GitHub Environment can assume.

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

      # GetObjectVersion too: the buckets are versioned, so old copies of a
      # deck are another way to read it
      actions = ["s3:GetObject", "s3:GetObjectVersion"]

      resources = [for arn in var.deny_object_read_bucket_arns : "${arn}/*"]
    }
  }
}

resource "aws_iam_role_policy" "state_access" {
  name   = "${var.role_name}-state"
  role   = aws_iam_role.plan.id
  policy = data.aws_iam_policy_document.state_access.json
}

# --- the role the deploy job assumes to push images and run terraform apply ---
#
# Only made when apply_role_name is set. It trusts one GitHub Environment, not
# the whole repository, so a pull request or a branch can never assume it: only
# a job that names that environment, which GitHub can hold for a reviewer's
# approval first.
#
# It can do a lot, because applying this configuration means creating buckets,
# a function, a service and IAM roles. Creating roles is the sharp part: a role
# that can make roles and pass them to Lambda can, in effect, give itself more.
# Its IAM writes are limited to roles under managed_role_prefix, but the real
# guard is who may approve the environment.

locals {
  apply_subjects = coalesce(var.apply_subjects, ["repo:${var.github_repository}:environment:production"])
}

data "aws_iam_policy_document" "assume_apply" {
  count = var.apply_role_name == null ? 0 : 1

  statement {
    effect  = "Allow"
    actions = ["sts:AssumeRoleWithWebIdentity"]

    principals {
      type        = "Federated"
      identifiers = [local.provider_arn]
    }

    condition {
      test     = "StringEquals"
      variable = "${local.oidc_hostname}:aud"
      values   = ["sts.amazonaws.com"]
    }

    condition {
      test     = "StringLike"
      variable = "${local.oidc_hostname}:sub"
      values   = local.apply_subjects
    }
  }
}

resource "aws_iam_role" "apply" {
  count = var.apply_role_name == null ? 0 : 1

  name               = var.apply_role_name
  description        = "Assumed by the GitHub Actions deploy job to push images and run terraform apply."
  assume_role_policy = data.aws_iam_policy_document.assume_apply[0].json

  max_session_duration = var.max_session_duration
}

# everything except IAM, which comes from the two below instead
resource "aws_iam_role_policy_attachment" "apply_power_user" {
  count = var.apply_role_name == null ? 0 : 1

  role       = aws_iam_role.apply[0].name
  policy_arn = "arn:aws:iam::aws:policy/PowerUserAccess"
}

# reading IAM is how a plan finds out what the roles look like now
resource "aws_iam_role_policy_attachment" "apply_iam_read" {
  count = var.apply_role_name == null ? 0 : 1

  role       = aws_iam_role.apply[0].name
  policy_arn = "arn:aws:iam::aws:policy/IAMReadOnlyAccess"
}

data "aws_iam_policy_document" "apply" {
  count = var.apply_role_name == null ? 0 : 1

  statement {
    sid    = "ManageThisProjectsRoles"
    effect = "Allow"

    actions = [
      "iam:CreateRole",
      "iam:DeleteRole",
      "iam:UpdateRole",
      "iam:UpdateRoleDescription",
      "iam:UpdateAssumeRolePolicy",
      "iam:TagRole",
      "iam:UntagRole",
      "iam:PutRolePolicy",
      "iam:DeleteRolePolicy",
      "iam:AttachRolePolicy",
      "iam:DetachRolePolicy",
      "iam:PassRole",
    ]

    resources = ["arn:aws:iam::*:role/${var.managed_role_prefix}*"]
  }

  statement {
    sid    = "ManageTheGitHubProvider"
    effect = "Allow"

    actions = [
      "iam:CreateOpenIDConnectProvider",
      "iam:DeleteOpenIDConnectProvider",
      "iam:TagOpenIDConnectProvider",
      "iam:UntagOpenIDConnectProvider",
      "iam:UpdateOpenIDConnectProviderThumbprint",
      "iam:AddClientIDToOpenIDConnectProvider",
      "iam:RemoveClientIDFromOpenIDConnectProvider",
    ]

    resources = ["arn:aws:iam::*:oidc-provider/${local.oidc_hostname}"]
  }

  # the same reason as the plan role: PowerUserAccess would let a workflow
  # download real course material, and deploying never needs to read a deck
  dynamic "statement" {
    for_each = length(var.deny_object_read_bucket_arns) > 0 ? [1] : []

    content {
      sid    = "NeverReadUploadedDecks"
      effect = "Deny"

      # GetObjectVersion too: the buckets are versioned, so old copies of a
      # deck are another way to read it
      actions = ["s3:GetObject", "s3:GetObjectVersion"]

      resources = [for arn in var.deny_object_read_bucket_arns : "${arn}/*"]
    }
  }
}

resource "aws_iam_role_policy" "apply" {
  count = var.apply_role_name == null ? 0 : 1

  name   = "${var.apply_role_name}-iam"
  role   = aws_iam_role.apply[0].id
  policy = data.aws_iam_policy_document.apply[0].json
}
