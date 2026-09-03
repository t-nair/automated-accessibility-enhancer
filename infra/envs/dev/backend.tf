# State lives in the bucket created by infra/bootstrap.
#
# use_lockfile is S3's own locking, which replaced the separate DynamoDB table
# older setups needed. There is nothing to create for it: Terraform writes a
# .tflock object next to the state file and deletes it when it is done.
#
# The values here cannot be variables, so bucket and region are filled in from
# the bootstrap outputs by hand, once.
terraform {
  backend "s3" {
    bucket       = "REPLACE_WITH_STATE_BUCKET_NAME"
    key          = "dev/terraform.tfstate"
    region       = "us-west-2"
    encrypt      = true
    use_lockfile = true
  }
}
