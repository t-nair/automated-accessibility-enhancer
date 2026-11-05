output "uploads_bucket" {
  description = "Upload a deck here to start the pipeline."
  value       = module.buckets.uploads_bucket_name
}

output "processed_bucket" {
  description = "The fixed deck, its report, and a status file appear here."
  value       = module.buckets.processed_bucket_name
}

output "ecr_repository_url" {
  description = "Push the pipeline image here."
  value       = module.ecr.repository_url
}

output "function_name" {
  value = module.pipeline.function_name
}

output "log_group_name" {
  value = module.observability.log_group_name
}

output "alarm_topic_arn" {
  value = module.observability.alarm_topic_arn
}

output "github_actions_role_arn" {
  description = "Give this to aws-actions/configure-aws-credentials as role-to-assume."
  value       = module.ci_role.role_arn
}

output "website_url" {
  description = "The address to open in a browser. App Runner provides the certificate."
  value       = module.website.service_url
}

output "web_repository_url" {
  description = "Where to push the website image."
  value       = module.ecr_web.repository_url
}

output "submissions_table_name" {
  description = "The DynamoDB table holding submissions."
  value       = module.submissions.table_name
}
