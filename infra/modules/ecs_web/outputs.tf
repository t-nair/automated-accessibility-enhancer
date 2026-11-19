output "endpoint" {
  description = "The address the website is served on. ECS provides the load balancer and its HTTPS certificate."
  value       = aws_ecs_express_gateway_service.this.ingress_paths[0].endpoint
}

output "service_arn" {
  value = aws_ecs_express_gateway_service.this.service_arn
}

output "task_role_arn" {
  description = "Role the running website uses, for anything that needs to grant it access later."
  value       = aws_iam_role.task.arn
}
