# The registry the Lambda container image is pushed to.
#
# Lambda reads the image from ECR every time it creates an execution
# environment, so the repository has to stay in the same region and account as
# the function.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

resource "aws_ecr_repository" "this" {
  name = var.repository_name

  # MUTABLE so that a tag like "latest" can be moved during development.
  # Set this to IMMUTABLE once deploys are pinned to a digest or a build number.
  image_tag_mutability = var.image_tag_mutability

  image_scanning_configuration {
    scan_on_push = true
  }

  encryption_configuration {
    encryption_type = "AES256"
  }

  force_delete = var.force_delete
}

# without this the repository grows by one image per deploy forever, and old
# images are pure cost
resource "aws_ecr_lifecycle_policy" "this" {
  repository = aws_ecr_repository.this.name

  policy = jsonencode({
    rules = [
      {
        rulePriority = 1
        description  = "Expire untagged images, which are only ever superseded build layers"
        selection = {
          tagStatus   = "untagged"
          countType   = "sinceImagePushed"
          countUnit   = "days"
          countNumber = var.untagged_image_retention_days
        }
        action = { type = "expire" }
      },
      {
        rulePriority = 2
        description  = "Keep the most recent tagged images so a rollback is still possible"
        selection = {
          tagStatus     = "tagged"
          tagPatternList = ["*"]
          countType     = "imageCountMoreThan"
          countNumber   = var.tagged_image_retention_count
        }
        action = { type = "expire" }
      },
    ]
  })
}
