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
  description = "Role pull requests assume to run terraform plan. Goes in the AWS_PLAN_ROLE_ARN repository variable."
  value       = module.ci_role.role_arn
}

output "github_actions_apply_role_arn" {
  description = "Role the deploy job assumes to push images and run terraform apply. Goes in the AWS_APPLY_ROLE_ARN repository variable."
  value       = module.ci_role.apply_role_arn
}

output "website_url" {
  description = "The address to open in a browser. ECS provides the load balancer and its certificate."
  value       = module.website.endpoint
}

output "web_repository_url" {
  description = "Where to push the website image."
  value       = module.ecr_web.repository_url
}

output "submissions_table_name" {
  description = "The DynamoDB table holding submissions."
  value       = module.submissions.table_name
}

output "cognito_user_pool_id" {
  value = module.auth.user_pool_id
}

output "cognito_sign_in_domain" {
  description = "Where the sign-in pages are served."
  value       = module.auth.domain_url
}

output "cognito_identity_provider_redirect_uri" {
  description = "The redirect URI to register in the Google and Microsoft apps."
  value       = module.auth.identity_provider_redirect_uri
}

output "email_dkim_records" {
  description = "CNAME records to add in Cloudflare (DNS only) so SES can send from email_domain."
  value       = module.auth.email_dkim_records
}
