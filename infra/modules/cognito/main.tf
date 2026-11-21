# Sign-in, with an Amazon Cognito user pool.
#
# Cognito holds the accounts, hashes the passwords, verifies email addresses,
# resets forgotten passwords and locks out repeated guesses, so the website does
# none of that itself. Email and password are typed into the website's own pages,
# which pass them to Cognito's API. Google goes through Cognito's OAuth endpoints
# and comes back with an authorization code, the usual OIDC flow. Accounts are
# keyed by Cognito's `sub`, never by email (see #77).
#
# Google is added only once its client id is set, so the pool works with email
# and password alone until that app is registered. Microsoft is left off for now
# (#77): Cognito can only sign in one Entra tenant per provider.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.23"
    }
  }
}

data "aws_caller_identity" "current" {}
data "aws_region" "current" {}

locals {
  google_enabled = var.google_client_id != null

  identity_providers = concat(["COGNITO"], local.google_enabled ? ["Google"] : [])

  # Cognito's sender allows about 50 emails a day, which a real launch would run
  # through in an afternoon. SES takes over once its domain is verified.
  use_ses = var.email_domain != null && var.email_domain_verified
}

resource "aws_cognito_user_pool" "this" {
  name = "${var.name_prefix}-users"

  # Essentials: the first 10,000 monthly active users are free, and it includes
  # managed login and SAML for NetID later
  user_pool_tier = "ESSENTIALS"

  username_attributes      = ["email"]
  auto_verified_attributes = ["email"]

  username_configuration {
    case_sensitive = false
  }

  # open sign-up, decided in #102
  admin_create_user_config {
    allow_admin_create_user_only = false
  }

  account_recovery_setting {
    recovery_mechanism {
      name     = "verified_email"
      priority = 1
    }
  }

  # deleting the pool deletes every account, and Terraform would do it to
  # replace the pool for a change as small as a renamed attribute
  deletion_protection = "ACTIVE"

  dynamic "email_configuration" {
    for_each = local.use_ses ? [1] : []

    content {
      email_sending_account = "DEVELOPER"
      source_arn            = aws_sesv2_email_identity.this[0].arn
      from_email_address    = "Accessibility Enhancer <${var.email_from_local_part}@${var.email_domain}>"
    }
  }
}

resource "aws_cognito_user_pool_domain" "this" {
  domain       = var.domain_prefix
  user_pool_id = aws_cognito_user_pool.this.id

  # 2 is managed login, the current sign-in pages. 1 is the older hosted UI.
  managed_login_version = 2
}

# The secret is put in Parameter Store by hand (see infra/README.md), so it never
# passes through GitHub. Terraform still keeps a copy in its state, because
# Cognito's provider_details is stored there; the state bucket guards it.
data "aws_ssm_parameter" "google_client_secret" {
  count = local.google_enabled ? 1 : 0

  name = var.google_client_secret_parameter_name
}

resource "aws_cognito_identity_provider" "google" {
  count = local.google_enabled ? 1 : 0

  user_pool_id  = aws_cognito_user_pool.this.id
  provider_name = "Google"
  provider_type = "Google"

  provider_details = {
    client_id        = var.google_client_id
    client_secret    = data.aws_ssm_parameter.google_client_secret[0].value
    authorize_scopes = "openid email profile"
  }

  attribute_mapping = {
    username       = "sub"
    email          = "email"
    email_verified = "email_verified"
    name           = "name"
  }
}

resource "aws_cognito_user_pool_client" "web" {
  name         = "${var.name_prefix}-web"
  user_pool_id = aws_cognito_user_pool.this.id

  # the website is a server, so it can keep a secret
  generate_secret = true

  allowed_oauth_flows_user_pool_client = true
  allowed_oauth_flows                  = ["code"]
  allowed_oauth_scopes                 = ["openid", "email", "profile"]
  callback_urls                        = var.callback_urls
  logout_urls                          = var.logout_urls
  supported_identity_providers         = local.identity_providers

  # the website's own sign-in page sends the email and password to InitiateAuth,
  # over HTTPS, with the client secret's hash
  explicit_auth_flows = ["ALLOW_USER_PASSWORD_AUTH", "ALLOW_REFRESH_TOKEN_AUTH"]

  prevent_user_existence_errors = "ENABLED"
  enable_token_revocation       = true

  depends_on = [aws_cognito_identity_provider.google]
}

# Cognito's default look for the sign-in pages. Without a style, managed login
# shows nothing for this client.
resource "aws_cognito_managed_login_branding" "web" {
  user_pool_id = aws_cognito_user_pool.this.id
  client_id    = aws_cognito_user_pool_client.web.id

  use_cognito_provided_values = true
}

# The client secret, where the website reads it at startup, the same way it
# reads the session key. It is never an environment variable or an output.
resource "aws_ssm_parameter" "client_secret" {
  name        = "/${var.project}/${var.environment}/cognito-client-secret"
  description = "The website's Cognito app client secret."
  type        = "SecureString"
  value       = aws_cognito_user_pool_client.web.client_secret
}

# --- email through SES ---
#
# Two applies. The first creates the domain identity, and its DKIM records are in
# the outputs to add in Cloudflare. Once SES shows the domain as verified, set
# email_domain_verified = true and apply again to switch Cognito over. Cognito
# refuses an identity that is not verified yet, which is why it is not one step.

resource "aws_sesv2_email_identity" "this" {
  count = var.email_domain == null ? 0 : 1

  email_identity = var.email_domain
}
