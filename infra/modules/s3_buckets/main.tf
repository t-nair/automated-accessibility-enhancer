# The two buckets a deck passes through: the one it is uploaded to, and the one
# the fixed copy and its report are written to.
#
# Both expire their contents on a timer. A deck is a working file, not a record:
# once the faculty member has downloaded the accessible version there is no
# reason to keep paying to store a copy of their lecture, and holding real
# course material longer than needed is a liability rather than a feature.

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

locals {
  # bucket names are global across all of AWS, so the account id is appended to
  # keep a second person deploying this from colliding with the first
  buckets = {
    uploads = {
      name            = "${var.name_prefix}-uploads-${data.aws_caller_identity.current.account_id}"
      expiration_days = var.uploads_expiration_days
    }
    processed = {
      name            = "${var.name_prefix}-processed-${data.aws_caller_identity.current.account_id}"
      expiration_days = var.processed_expiration_days
    }
  }
}

resource "aws_s3_bucket" "this" {
  for_each = local.buckets

  bucket = each.value.name

  # these hold transient working files, so tearing the environment down should
  # not need a manual pass to empty them first
  force_destroy = var.force_destroy
}

resource "aws_s3_bucket_public_access_block" "this" {
  for_each = aws_s3_bucket.this

  bucket = each.value.id

  block_public_acls       = true
  block_public_policy     = true
  ignore_public_acls      = true
  restrict_public_buckets = true
}

resource "aws_s3_bucket_ownership_controls" "this" {
  for_each = aws_s3_bucket.this

  bucket = each.value.id

  rule {
    object_ownership = "BucketOwnerEnforced"
  }
}

# Versioning is the backup: a deck deleted or overwritten by mistake can be
# brought back from its previous version until the lifecycle rule below removes
# it, a few days later. It is not an archive, so old versions do not outlive the
# expiry above by more than that window.
resource "aws_s3_bucket_versioning" "this" {
  for_each = aws_s3_bucket.this

  bucket = each.value.id

  versioning_configuration {
    status = "Enabled"
  }
}

resource "aws_s3_bucket_server_side_encryption_configuration" "this" {
  for_each = aws_s3_bucket.this

  bucket = each.value.id

  rule {
    apply_server_side_encryption_by_default {
      sse_algorithm = "AES256"
    }

    bucket_key_enabled = true
  }
}

# the expiry that keeps this from becoming a permanent archive of other
# people's course material
resource "aws_s3_bucket_lifecycle_configuration" "this" {
  for_each = local.buckets

  bucket = aws_s3_bucket.this[each.key].id

  rule {
    id     = "expire-decks"
    status = "Enabled"

    filter {}

    expiration {
      days = each.value.expiration_days
    }

    # an upload that was abandoned partway still costs storage until it is
    # cleaned up, and it is invisible in the console until then
    abort_incomplete_multipart_upload {
      days_after_initiation = 1
    }
  }

  # With versioning on, the expiry above and a delete from the app only hide the
  # file behind a delete marker. This rule is what actually removes it, and then
  # clears the marker left with nothing behind it.
  rule {
    id     = "remove-old-versions"
    status = "Enabled"

    filter {}

    noncurrent_version_expiration {
      noncurrent_days = var.noncurrent_version_expiration_days
    }

    expiration {
      expired_object_delete_marker = true
    }
  }

  depends_on = [
    aws_s3_bucket_ownership_controls.this,
    aws_s3_bucket_versioning.this,
  ]
}

resource "aws_s3_bucket_policy" "this" {
  for_each = aws_s3_bucket.this

  bucket = each.value.id
  policy = data.aws_iam_policy_document.deny_insecure_transport[each.key].json

  depends_on = [aws_s3_bucket_public_access_block.this]
}

data "aws_iam_policy_document" "deny_insecure_transport" {
  for_each = aws_s3_bucket.this

  statement {
    sid    = "DenyInsecureTransport"
    effect = "Deny"

    principals {
      type        = "*"
      identifiers = ["*"]
    }

    actions = ["s3:*"]

    resources = [
      each.value.arn,
      "${each.value.arn}/*",
    ]

    condition {
      test     = "Bool"
      variable = "aws:SecureTransport"
      values   = ["false"]
    }
  }
}
