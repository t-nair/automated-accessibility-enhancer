variable "repository_name" {
  description = "Name of the ECR repository holding the pipeline image."
  type        = string
}

variable "image_tag_mutability" {
  description = "MUTABLE lets a tag be moved, which is convenient while developing. IMMUTABLE is the safer setting once deploys are automated."
  type        = string
  default     = "MUTABLE"

  validation {
    condition     = contains(["MUTABLE", "IMMUTABLE"], var.image_tag_mutability)
    error_message = "image_tag_mutability has to be MUTABLE or IMMUTABLE."
  }
}

variable "untagged_image_retention_days" {
  description = "How long an untagged image is kept. These are leftover layers from a newer build, so they do not need long."
  type        = number
  default     = 7
}

variable "tagged_image_retention_count" {
  description = "How many tagged images to keep, oldest expired first. Enough to roll back a bad deploy."
  type        = number
  default     = 10
}

variable "force_delete" {
  description = "Whether destroying the environment also deletes images still in the repository."
  type        = bool
  default     = true
}
