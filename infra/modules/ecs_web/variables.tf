variable "service_name" {
  description = "Name of the service and its cluster, and the start of its three IAM role names."
  type        = string
}

variable "image_uri" {
  description = "The website image in ECR, including its tag."
  type        = string
}

variable "uploads_bucket_name" {
  description = "Bucket the website puts uploaded decks in. Putting one there is also what starts the pipeline."
  type        = string
}

variable "uploads_bucket_arn" {
  description = "ARN of the uploads bucket, for the policy."
  type        = string
}

variable "processed_bucket_name" {
  description = "Bucket the pipeline writes fixed decks and reports to."
  type        = string
}

variable "processed_bucket_arn" {
  description = "ARN of the processed bucket, for the policy."
  type        = string
}

variable "submissions_table_name" {
  description = "Name of the submissions table."
  type        = string
}

variable "submissions_table_arn" {
  description = "ARN of the submissions table."
  type        = string
}

variable "submissions_index_arn" {
  description = "ARN covering the table's indexes. Querying one visitor's submissions needs this as well as the table itself."
  type        = string
}

variable "secret_key_parameter_name" {
  description = "Parameter Store name holding the key that signs the session cookie."
  type        = string
}

variable "secret_key_parameter_arn" {
  description = "ARN of that parameter, for the policy."
  type        = string
}

variable "pipeline_timeout_seconds" {
  description = "How long the pipeline is allowed to take. The website uses it to decide a run has died, so it should match the Lambda's own timeout."
  type        = number
  default     = 900
}

variable "cpu" {
  description = "CPU units for each task (1024 is one vCPU). The website only serves pages and hands out signed links, so the smallest size, 256, is plenty."
  type        = number
  default     = 256
}

variable "memory" {
  description = "Memory for each task in MiB. 512 is the smallest Express Mode allows."
  type        = number
  default     = 512
}

variable "min_tasks" {
  description = "Tasks kept running. One stops the first visitor of the day waiting for a cold start."
  type        = number
  default     = 1
}

variable "max_tasks" {
  description = "How far the service may scale out."
  type        = number
  default     = 2
}

variable "log_retention_days" {
  description = "How long CloudWatch keeps the website's logs."
  type        = number
  default     = 30
}

variable "subnet_ids" {
  description = "Subnets to run in, at least two. Leave null to use the account's default VPC."
  type        = list(string)
  default     = null
}

variable "security_group_ids" {
  description = "Security groups for the tasks when subnet_ids is set. Leave null for the VPC's default group."
  type        = list(string)
  default     = null
}

variable "public_url" {
  description = "The address people reach the website at, without a trailing slash."
  type        = string
}

variable "cognito" {
  description = "The user pool the website signs people in with. Leave null to keep the signed-cookie identity."
  type = object({
    issuer                       = string
    domain_url                   = string
    client_id                    = string
    client_secret_parameter_name = string
    client_secret_parameter_arn  = string
    google_enabled               = bool
  })
  default = null
}
