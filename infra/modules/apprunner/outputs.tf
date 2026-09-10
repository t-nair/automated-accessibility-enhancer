output "service_url" {
  description = "The address the website is served on. App Runner provides the HTTPS certificate for it."
  value       = "https://${aws_apprunner_service.this.service_url}"
}

output "service_arn" {
  description = "ARN of the App Runner service."
  value       = aws_apprunner_service.this.arn
}

output "instance_role_arn" {
  description = "Role the running website uses, for anything that needs to grant it access later."
  value       = aws_iam_role.instance.arn
}
