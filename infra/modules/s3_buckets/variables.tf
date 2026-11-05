variable "name_prefix" {
  description = "Prefix for both bucket names, usually project and environment. The account id is appended so the names are globally unique."
  type        = string
}

variable "uploads_expiration_days" {
  description = "How many days an uploaded deck is kept before S3 deletes it. The pipeline is finished with it in minutes; this is the window for a retry."
  type        = number
  default     = 7

  validation {
    condition     = var.uploads_expiration_days >= 1
    error_message = "S3 lifecycle expiry has to be at least 1 day."
  }
}

variable "processed_expiration_days" {
  description = "How many days a fixed deck and its report are kept. Long enough that someone can come back for the download after a weekend."
  type        = number
  default     = 30

  validation {
    condition     = var.processed_expiration_days >= 1
    error_message = "S3 lifecycle expiry has to be at least 1 day."
  }
}

variable "force_destroy" {
  description = "Whether destroying the environment also empties the buckets. True because these only ever hold transient working files."
  type        = bool
  default     = true
}
