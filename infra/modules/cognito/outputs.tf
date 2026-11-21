output "user_pool_id" {
  value = aws_cognito_user_pool.this.id
}

output "user_pool_arn" {
  value = aws_cognito_user_pool.this.arn
}

# what the website points its OIDC client at; the discovery document is under it
output "issuer" {
  value = "https://cognito-idp.${data.aws_region.current.region}.amazonaws.com/${aws_cognito_user_pool.this.id}"
}

output "client_id" {
  value = aws_cognito_user_pool_client.web.id
}

output "client_secret_parameter_name" {
  value = aws_ssm_parameter.client_secret.name
}

output "client_secret_parameter_arn" {
  value = aws_ssm_parameter.client_secret.arn
}

output "domain_url" {
  value = "https://${aws_cognito_user_pool_domain.this.domain}.auth.${data.aws_region.current.region}.amazoncognito.com"
}

# the redirect URI to register in the Google OAuth client
output "identity_provider_redirect_uri" {
  value = "https://${aws_cognito_user_pool_domain.this.domain}.auth.${data.aws_region.current.region}.amazoncognito.com/oauth2/idpresponse"
}

# the three CNAME records to add in Cloudflare (DNS only, not proxied)
output "email_dkim_records" {
  value = var.email_domain == null ? [] : [
    for token in aws_sesv2_email_identity.this[0].dkim_signing_attributes[0].tokens : {
      name  = "${token}._domainkey.${var.email_domain}"
      type  = "CNAME"
      value = "${token}.dkim.amazonses.com"
    }
  ]
}

# so the website shows "Continue with Google" only when it would work
output "google_enabled" {
  value = local.google_enabled
}

# the SES domain identity, for the website's bug-report emails. Null until the
# domain is set and verified, the same rule Cognito itself follows.
output "email_identity_arn" {
  value = local.use_ses ? aws_sesv2_email_identity.this[0].arn : null
}
