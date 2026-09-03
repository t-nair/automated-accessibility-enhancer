output "state_bucket_name" {
  description = "Put this in the bucket field of each environment's backend block."
  value       = aws_s3_bucket.state.id
}

output "state_bucket_region" {
  description = "Put this in the region field of each environment's backend block."
  value       = var.aws_region
}

output "backend_block" {
  description = "The backend configuration to paste into an environment, already filled in."
  value       = <<-CONFIG
    terraform {
      backend "s3" {
        bucket       = "${aws_s3_bucket.state.id}"
        key          = "<environment>/terraform.tfstate"
        region       = "${var.aws_region}"
        encrypt      = true
        use_lockfile = true
      }
    }
  CONFIG
}
