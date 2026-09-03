variable "function_name" {
  description = "Name of the Lambda function being watched. The log group is /aws/lambda/<this>."
  type        = string
}

variable "environment" {
  description = "Environment name, used to keep one environment's custom metrics apart from another's."
  type        = string
}

variable "log_retention_days" {
  description = "How long CloudWatch keeps the function's logs. Left to Lambda this would be forever, which is the main reason this module owns the log group."
  type        = number
  default     = 30
}

variable "timeout_seconds" {
  description = "The function's timeout. The duration alarm is a fraction of this, so the two have to be kept in step."
  type        = number
}

variable "duration_alarm_fraction" {
  description = "Fraction of the timeout a run has to reach before the duration alarm fires."
  type        = number
  default     = 0.8

  validation {
    condition     = var.duration_alarm_fraction > 0 && var.duration_alarm_fraction < 1
    error_message = "duration_alarm_fraction has to be between 0 and 1, exclusive."
  }
}

variable "error_threshold" {
  description = "Failed runs in a five minute window before the error alarm fires."
  type        = number
  default     = 1
}

variable "caption_failure_threshold" {
  description = "Pictures falling back to placeholder alt text in a five minute window before the alarm fires. A single failure can be one odd image; a run of them means Bedrock is unreachable."
  type        = number
  default     = 3
}

variable "alarm_email_addresses" {
  description = "Addresses to notify. Each one gets a confirmation email from AWS and receives nothing until someone clicks the link in it."
  type        = list(string)
  default     = []
}
