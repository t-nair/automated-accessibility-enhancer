# The table of submissions.
#
# The website writes a row when a deck is uploaded, the pipeline updates it as it
# works, and the website reads it back to draw the status page. It replaces the
# sqlite file the local version uses, which cannot be shared between instances.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

resource "aws_dynamodb_table" "submissions" {
  name = var.table_name

  # on demand, because the traffic is a handful of uploads a day with long gaps.
  # Provisioned capacity would mean paying for a steady rate nobody is using.
  billing_mode = "PAY_PER_REQUEST"

  hash_key = "id"

  attribute {
    name = "id"
    type = "S"
  }

  attribute {
    name = "owner_id"
    type = "S"
  }

  attribute {
    name = "submitted_at"
    type = "S"
  }

  # one visitor's submissions, newest first, without reading the whole table.
  # The name has to match OWNER_INDEX in storage_aws.py.
  global_secondary_index {
    name            = "owner_id-submitted_at-index"
    projection_type = "ALL"

    key_schema {
      attribute_name = "owner_id"
      key_type       = "HASH"
    }

    key_schema {
      attribute_name = "submitted_at"
      key_type       = "RANGE"
    }
  }

  # rows delete themselves once expires_at passes, which is how the retention
  # period is kept without anything having to sweep the table
  ttl {
    attribute_name = "expires_at"
    enabled        = true
  }

  point_in_time_recovery {
    enabled = var.point_in_time_recovery
  }

  server_side_encryption {
    enabled = true
  }
}
