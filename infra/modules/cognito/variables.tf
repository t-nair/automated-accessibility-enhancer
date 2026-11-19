variable "name_prefix" {
  description = "Prefix for the pool and client names, usually project and environment."
  type        = string
}

variable "project" {
  description = "Project name, for the Parameter Store path."
  type        = string
}

variable "environment" {
  description = "Environment name, for the Parameter Store path."
  type        = string
}

variable "domain_prefix" {
  description = "The sign-in pages are served at https://<domain_prefix>.auth.<region>.amazoncognito.com. Unique across the region; may not contain aws, amazon or cognito."
  type        = string
}

variable "callback_urls" {
  description = "Where Cognito may send people back to after sign-in. http is allowed only for localhost."
  type        = list(string)
}

variable "logout_urls" {
  description = "Where Cognito may send people after sign-out."
  type        = list(string)
}

# --- Google ---

variable "google_client_id" {
  description = "OAuth client id from Google Cloud. Leave null to leave Google sign-in off."
  type        = string
  default     = null
}

variable "google_client_secret_parameter_name" {
  description = "The SecureString parameter holding the Google OAuth client secret. Read only when google_client_id is set."
  type        = string
  default     = null
}

# --- email ---

variable "email_domain" {
  description = "Domain to send verification and reset emails from through SES. Leave null to use Cognito's own sender (about 50 emails a day)."
  type        = string
  default     = null
}

variable "email_domain_verified" {
  description = "Set true once SES shows email_domain as verified. Cognito refuses an unverified identity."
  type        = bool
  default     = false
}

variable "email_from_local_part" {
  description = "The part before the @ in the From address."
  type        = string
  default     = "no-reply"
}
