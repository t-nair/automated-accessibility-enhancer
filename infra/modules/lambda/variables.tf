variable "function_name" {
  description = "Name of the Lambda function. Its log group is /aws/lambda/<this>."
  type        = string
}

variable "image_uri" {
  description = "Full ECR image URI including the tag or digest, for example 123456789012.dkr.ecr.us-west-2.amazonaws.com/aae-pipeline:latest."
  type        = string
}

variable "architecture" {
  description = "Has to match what the image was built for. Building on an Apple Silicon machine without --platform produces arm64."
  type        = string
  default     = "x86_64"

  validation {
    condition     = contains(["x86_64", "arm64"], var.architecture)
    error_message = "architecture has to be x86_64 or arm64."
  }
}

variable "uploads_bucket_name" {
  description = "Bucket decks are uploaded to. The S3 notification is attached to it."
  type        = string
}

variable "uploads_bucket_arn" {
  description = "ARN of the uploads bucket, used to scope both the read permission and the invoke permission."
  type        = string
}

variable "processed_bucket_name" {
  description = "Bucket the fixed deck and its report are written to."
  type        = string
}

variable "processed_bucket_arn" {
  description = "ARN of the processed bucket, used to scope the write permission."
  type        = string
}

variable "bedrock_model_id" {
  description = "Claude model used for captioning. Model IDs on Bedrock carry an anthropic. prefix."
  type        = string
  default     = "anthropic.claude-opus-5"
}

variable "bedrock_region" {
  description = "Region Bedrock requests are sent to. Has to be one that serves the bedrock-mantle endpoint, which is not every region."
  type        = string
  default     = "us-west-2"
}

variable "bedrock_model_arns" {
  description = <<-DESC
    Model ARNs the function is allowed to run inference against.

    Defaults to * because AWS does not currently publish the resource ARN
    format for bedrock-mantle:CreateInference. Narrow it to the specific model
    once that is confirmed for the account: run the pipeline once and read the
    resource off the CloudTrail event for the call.
  DESC
  type        = list(string)
  default     = ["*"]
}

variable "timeout_seconds" {
  description = "How long one deck may take. 900 is the Lambda maximum."
  type        = number
  default     = 900

  validation {
    condition     = var.timeout_seconds > 0 && var.timeout_seconds <= 900
    error_message = "Lambda timeouts have to be between 1 and 900 seconds."
  }
}

variable "memory_mb" {
  description = "Memory, and with it the CPU share the function gets."
  type        = number
  default     = 2048
}

variable "ephemeral_storage_mb" {
  description = "Size of /tmp, which holds the uploaded deck and the rewritten copy at the same time."
  type        = number
  default     = 1024

  validation {
    condition     = var.ephemeral_storage_mb >= 512 && var.ephemeral_storage_mb <= 10240
    error_message = "Lambda ephemeral storage has to be between 512 MB and 10240 MB."
  }
}

variable "reserved_concurrency" {
  description = "Ceiling on decks in flight at once. Keeps a bulk upload inside the Bedrock rate limit and caps the worst case spend."
  type        = number
  default     = 5
}

variable "log_level" {
  description = "Level the pipeline's own log lines are emitted at."
  type        = string
  default     = "INFO"
}

variable "tags" {
  description = "Extra tags for the function, on top of the provider's default tags."
  type        = map(string)
  default     = {}
}

variable "submissions_table_name" {
  description = "Name of the submissions table. The pipeline updates the row for the deck it is working on so the website can show progress."
  type        = string
}

variable "submissions_table_arn" {
  description = "ARN of the submissions table, for the policy."
  type        = string
}
