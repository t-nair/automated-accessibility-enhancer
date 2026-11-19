variable "github_repository" {
  description = "Repository allowed to assume the role, as owner/name."
  type        = string

  validation {
    condition     = can(regex("^[^/]+/[^/]+$", var.github_repository))
    error_message = "github_repository has to be in owner/name form, for example t-nair/automated-accessibility-enhancer."
  }
}

variable "allowed_subjects" {
  description = "Override the trusted OIDC subjects. Leave null to trust every workflow in the repository; set something like [\"repo:owner/name:ref:refs/heads/main\"] to trust only one branch."
  type        = list(string)
  default     = null
}

variable "create_oidc_provider" {
  description = "Whether to create the GitHub OIDC provider. An account can only hold one per URL, so set this false if the account already uses GitHub Actions."
  type        = bool
  default     = true
}

variable "role_name" {
  description = "Name of the role GitHub Actions assumes."
  type        = string
  default     = "github-actions-terraform-plan"
}

variable "state_bucket_name" {
  description = "Terraform state bucket. The role needs to read state and to take and release the lock."
  type        = string
}

variable "deny_object_read_bucket_arns" {
  description = "Buckets the role is explicitly denied GetObject on, to stop ReadOnlyAccess from reaching real course material. Usually the uploads and processed buckets."
  type        = list(string)
  default     = []
}

variable "max_session_duration" {
  description = "How long the assumed credentials last. A plan finishes in minutes, so there is no reason for this to be long."
  type        = number
  default     = 3600
}

variable "apply_role_name" {
  description = "Name of the role the deploy job assumes to run terraform apply. Leave null for no such role, so CI can only plan."
  type        = string
  default     = null
}

variable "apply_subjects" {
  description = "OIDC subjects the apply role trusts. Leave null to trust only jobs that use the repository's production environment."
  type        = list(string)
  default     = null
}

variable "managed_role_prefix" {
  description = "The apply role may only create, change and pass IAM roles whose names start with this, for example \"aae-dev-\"."
  type        = string
  default     = ""

  validation {
    condition     = var.apply_role_name == null || length(var.managed_role_prefix) > 0
    error_message = "Set managed_role_prefix when apply_role_name is set, or the apply role could change every role in the account."
  }
}
