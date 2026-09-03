variable "aws_region" {
  description = "Region the state bucket lives in. Every environment's backend has to name this same region."
  type        = string
  default     = "us-west-2"
}

variable "state_bucket_name" {
  description = "Name of the Terraform state bucket. S3 bucket names are global, so this has to be unique across all of AWS, not just this account."
  type        = string
}

variable "state_version_retention_days" {
  description = "How long a superseded state file is kept before it is deleted. Long enough to undo a bad apply."
  type        = number
  default     = 90
}

variable "tags" {
  description = "Tags put on everything this configuration creates."
  type        = map(string)

  default = {
    Project   = "automated-accessibility-enhancer"
    ManagedBy = "terraform"
  }
}
