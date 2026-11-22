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
  description = "Region everything is created in. us-east-1 by default so that it matches the Bedrock region and image bytes never cross between regions."
  type        = string
  default     = "us-east-1"
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

# --- backups ---

variable "noncurrent_version_expiration_days" {
  description = "Days a deleted or overwritten file can still be restored from its previous version."
  type        = number
  default     = 3
}

variable "buckets_force_destroy" {
  description = "Whether terraform destroy empties the buckets, old versions included. Set true only for a scratch environment."
  type        = bool
  default     = false
}

# --- captioning ---

variable "bedrock_model_id" {
  description = "Claude model that writes image descriptions: a Bedrock inference profile id, with \"us.\" in front and a version on the end. Haiku 4.5 is what production runs; Opus 5 costs about eight times as much per picture (see docs/cost-report.docx)."
  type        = string
  default     = "us.anthropic.claude-haiku-4-5-20251001-v1:0"
}

variable "bedrock_region" {
  description = "Region Bedrock requests go to, where the inference profile lives. us-east-1 is where image descriptions were proven to work; us-west-2 and us-east-2 answer that the model does not exist."
  type        = string
  default     = "us-east-1"
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
  description = "Ceiling on decks in flight at once. -1 reserves nothing, which a new account needs: its total of 10 has to stay unreserved until AWS raises the limit."
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
  description = "Whether the submissions table keeps continuous backups, so it can be restored to any second in the last 35 days. Turn off only for a scratch environment."
  type        = bool
  default     = true
}

variable "web_cpu" {
  description = "CPU units for each website task. 1024 is one vCPU; 256 is the smallest."
  type        = number
  default     = 256
}

variable "web_memory" {
  description = "Memory for each website task in MiB. 512 is the smallest."
  type        = number
  default     = 512
}

variable "web_min_tasks" {
  description = "Website tasks kept running. One stops the first visitor waiting for a cold start."
  type        = number
  default     = 1
}

variable "web_max_tasks" {
  description = "How far the website may scale out."
  type        = number
  default     = 2
}

variable "web_subnet_ids" {
  description = "Subnets for the website, at least two. Leave null to use the account's default VPC, which an account without one cannot do."
  type        = list(string)
  default     = null
}

# --- sign-in (Cognito, #77) ---

variable "public_url" {
  description = "The address people reach the website at, without a trailing slash. Sign-in redirects are built from it."
  type        = string
  default     = "https://accessibilityenhancer.com/app"
}

variable "cognito_domain_prefix" {
  description = "Sign-in pages are served at https://<prefix>.auth.<region>.amazoncognito.com. Unique across the region."
  type        = string
  default     = "accessibility-enhancer"
}

variable "cognito_extra_callback_urls" {
  description = "More sign-in callback URLs, such as http://localhost:5000/auth/callback for trying sign-in locally."
  type        = list(string)
  default     = []
}

variable "cognito_extra_logout_urls" {
  description = "More sign-out URLs, to match cognito_extra_callback_urls."
  type        = list(string)
  default     = []
}

variable "google_client_id" {
  description = "OAuth client id from Google Cloud. Leave null to leave Google sign-in off. The secret goes in Parameter Store, not here (see infra/README.md)."
  type        = string
  default     = null
}

# Defaults rather than TFVARS: the domain is public, and the secret cannot be read back to add a line.
variable "email_domain" {
  description = "Domain to send sign-in and bug-report emails from through SES. Set null to use Cognito's own sender (about 50 a day)."
  type        = string
  default     = "accessibilityenhancer.com"
}

variable "email_domain_verified" {
  description = "Set true once SES shows email_domain as verified. Cognito refuses an unverified domain, so this cannot be true on the apply that creates it."
  type        = bool
  default     = true
}

variable "bug_report_to" {
  description = "Inbox the \"Report a bug\" form emails. Sent once email_domain is verified."
  type        = string
  default     = "hello@accessibilityenhancer.com"
}
