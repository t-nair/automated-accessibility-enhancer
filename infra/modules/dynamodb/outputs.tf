output "table_name" {
  description = "Name of the submissions table, for the website and the pipeline."
  value       = aws_dynamodb_table.submissions.name
}

output "table_arn" {
  description = "ARN of the submissions table, for the policies that grant access to it."
  value       = aws_dynamodb_table.submissions.arn
}

output "index_arn" {
  description = "ARN covering the table's indexes. Querying one needs its own permission, separate from the table."
  value       = "${aws_dynamodb_table.submissions.arn}/index/*"
}
