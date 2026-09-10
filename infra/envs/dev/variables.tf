variable "project" {
  description = "Short project name. Used as the first part of every resource name."
  type        = string
  default     = "aae"
}

variable "environment" {
  description = "Environment name. Used in resource names and to keep custom metrics separate."
  type        = string
  default     = "dev"
}

variable "aws_region" {
  description = "Region everything is created in."
  type        = string
  default     = "us-west-2"
}

variable "state_bucket_name" {
  description = "The bucket created by infra/bootstrap. The CI role needs it by name to be granted access to the state and its lock."
  type        = string
}

# --- container image ---

variable "image_tag" {
  description = "Tag of the image in ECR to run. Move this to a build number or digest once deploys are automated; 'latest' is convenient but makes it impossible to tell what is deployed."
  type        = string
  default     = "latest"
}

variable "lambda_architecture" {
  description = "Has to match the architecture the image was built for."
  type        = string
  default     = "x86_64"
}

# --- how long decks are kept ---

variable "uploads_expiration_days" {
  description = "Days an uploaded deck is kept before S3 deletes it."
  type        = number
  default     = 7
}

variable "processed_expiration_days" {
  description = "Days a fixed deck and its report are kept before S3 deletes them."
  type        = number
  default     = 30
}

# --- captioning ---

variable "bedrock_model_id" {
  description = "Claude model used for captioning. Model IDs on Bedrock carry an anthropic. prefix."
  type        = string
  default     = "anthropic.claude-opus-5"
}

variable "bedrock_region" {
  description = "Region Bedrock requests go to. Has to serve the bedrock-mantle endpoint."
  type        = string
  default     = "us-west-2"
}

# --- function sizing ---

variable "lambda_timeout_seconds" {
  description = "How long one deck may take. 900 is the Lambda maximum."
  type        = number
  default     = 900
}

variable "lambda_memory_mb" {
  description = "Memory, and with it the CPU share the function gets."
  type        = number
  default     = 2048
}

variable "lambda_reserved_concurrency" {
  description = "Ceiling on decks in flight at once."
  type        = number
  default     = 5
}

# --- logs and alarms ---

variable "log_retention_days" {
  description = "How long CloudWatch keeps the function's logs."
  type        = number
  default     = 30
}

variable "alarm_email_addresses" {
  description = "Addresses notified when an alarm fires. Each has to confirm the subscription by email before it receives anything."
  type        = list(string)
  default     = []
}

# --- CI ---

variable "github_repository" {
  description = "Repository whose workflows may assume the plan role, as owner/name."
  type        = string
  default     = "t-nair/automated-accessibility-enhancer"
}

variable "create_oidc_provider" {
  description = "Set false if the account already has a GitHub OIDC provider, since only one per URL is allowed."
  type        = bool
  default     = true
}

variable "submissions_point_in_time_recovery" {
  description = "Whether the submissions table keeps continuous backups. Worth turning on for anything people rely on."
  type        = bool
  default     = false
}

variable "web_cpu" {
  description = "vCPU for each website instance."
  type        = string
  default     = "0.25 vCPU"
}

variable "web_memory" {
  description = "Memory for each website instance."
  type        = string
  default     = "0.5 GB"
}

variable "web_min_instances" {
  description = "Website instances kept warm. One is App Runner's lowest."
  type        = number
  default     = 1
}

variable "web_max_instances" {
  description = "How far the website may scale out."
  type        = number
  default     = 2
}

variable "web_deploy_on_push" {
  description = "Whether pushing a new website image deploys it straight away."
  type        = bool
  default     = true
}
