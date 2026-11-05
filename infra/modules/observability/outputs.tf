output "log_group_name" {
  value = aws_cloudwatch_log_group.lambda.name
}

output "log_group_arn" {
  value = aws_cloudwatch_log_group.lambda.arn
}

output "alarm_topic_arn" {
  description = "Subscribe anything else that should hear about failures to this."
  value       = aws_sns_topic.alarms.arn
}
