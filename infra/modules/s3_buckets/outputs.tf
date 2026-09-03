output "uploads_bucket_name" {
  value = aws_s3_bucket.this["uploads"].id
}

output "uploads_bucket_arn" {
  value = aws_s3_bucket.this["uploads"].arn
}

output "processed_bucket_name" {
  value = aws_s3_bucket.this["processed"].id
}

output "processed_bucket_arn" {
  value = aws_s3_bucket.this["processed"].arn
}
