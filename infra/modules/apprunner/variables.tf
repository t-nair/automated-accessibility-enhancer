variable "service_name" {
  description = "Name of the App Runner service, also used to name its two IAM roles."
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
  description = "vCPU for each instance. The website only serves pages and hands out signed links, so the smallest size is plenty."
  type        = string
  default     = "0.25 vCPU"
}

variable "memory" {
  description = "Memory for each instance."
  type        = string
  default     = "0.5 GB"
}

variable "min_instances" {
  description = "Instances kept warm. One is the lowest App Runner allows, and is what stops the first visitor of the day waiting for a cold start."
  type        = number
  default     = 1
}

variable "max_instances" {
  description = "How far the service may scale out."
  type        = number
  default     = 2
}

variable "max_requests_per_instance" {
  description = "Requests one instance handles at once before another is started."
  type        = number
  default     = 50
}

variable "deploy_on_push" {
  description = "Whether pushing a new image with the same tag deploys it. Convenient while developing, worth turning off once real people use the site."
  type        = bool
  default     = true
}
