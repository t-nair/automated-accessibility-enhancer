variable "table_name" {
  description = "Name of the DynamoDB table holding submissions."
  type        = string
}

variable "point_in_time_recovery" {
  description = "Whether to keep continuous backups. Worth having in production; the extra cost is not worth it for a scratch environment."
  type        = bool
  default     = false
}
