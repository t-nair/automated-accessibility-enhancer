output "role_arn" {
  description = "Pass this to aws-actions/configure-aws-credentials as role-to-assume."
  value       = aws_iam_role.plan.arn
}

output "role_name" {
  value = aws_iam_role.plan.name
}

output "oidc_provider_arn" {
  value = local.provider_arn
}

output "apply_role_arn" {
  description = "Role the deploy job assumes, or null when there is none."
  value       = one(aws_iam_role.apply[*].arn)
}
