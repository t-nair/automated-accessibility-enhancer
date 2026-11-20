# Sign-in with Amazon Cognito (#77). Optional: anyone can use the tool without an
# account, and signing in only means their files follow them to another browser.
#
# The standard OAuth authorization code flow with PKCE, done with the standard
# library. After the code is exchanged, the person's id and email come from
# Cognito's userInfo endpoint, over HTTPS with the access token, so no JWT has to
# be decoded or verified here.
#
# Turned on by the COGNITO_* variables Terraform gives the website. Without them
# (a laptop, or a server of your own) everything works as before, with the
# signed-cookie identity alone.

import base64
import hashlib
import json
import os
import secrets
import urllib.parse
import urllib.request

COGNITO_DOMAIN = os.environ.get("COGNITO_DOMAIN", "").rstrip("/")
COGNITO_CLIENT_ID = os.environ.get("COGNITO_CLIENT_ID", "")

# the address people reach the site at. The Cloudflare proxy drops Host, so the
# redirect URI Cognito checks cannot be built from the request.
PUBLIC_URL = os.environ.get("PUBLIC_URL", "").rstrip("/")

ENABLED = bool(COGNITO_DOMAIN and COGNITO_CLIENT_ID)

SCOPES = "openid email profile"

_client_secret = None


# read once, on the first sign-in rather than at startup, so a problem with
# Parameter Store cannot stop the site from serving people who never sign in
def client_secret():
    global _client_secret

    if _client_secret is None:
        if os.environ.get("COGNITO_CLIENT_SECRET"):
            _client_secret = os.environ["COGNITO_CLIENT_SECRET"]
        else:
            import boto3

            answer = boto3.client("ssm").get_parameter(
                Name=os.environ["COGNITO_CLIENT_SECRET_PARAMETER"], WithDecryption=True
            )
            _client_secret = answer["Parameter"]["Value"]

    return _client_secret


def new_state_and_verifier():
    return secrets.token_urlsafe(32), secrets.token_urlsafe(64)


def code_challenge(verifier):
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    return base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")


# /signup and /oauth2/authorize take the same parameters; one opens on the
# create-account form, the other on sign-in
def sign_in_url(redirect_uri, state, verifier, create_account=False):
    path = "/signup" if create_account else "/oauth2/authorize"
    query = urllib.parse.urlencode({
        "response_type": "code",
        "client_id": COGNITO_CLIENT_ID,
        "redirect_uri": redirect_uri,
        "scope": SCOPES,
        "state": state,
        "code_challenge": code_challenge(verifier),
        "code_challenge_method": "S256",
    })
    return f"{COGNITO_DOMAIN}{path}?{query}"


def sign_out_url(logout_uri):
    query = urllib.parse.urlencode({"client_id": COGNITO_CLIENT_ID, "logout_uri": logout_uri})
    return f"{COGNITO_DOMAIN}/logout?{query}"


def _call(request):
    with urllib.request.urlopen(request, timeout=10) as response:
        return json.loads(response.read().decode("utf-8"))


# gives back {"sub": ..., "email": ...} for the person who signed in.
# Raises on anything unexpected; the caller turns that into a message.
def user_for_code(code, verifier, redirect_uri):
    basic = base64.b64encode(f"{COGNITO_CLIENT_ID}:{client_secret()}".encode()).decode()
    tokens = _call(urllib.request.Request(
        f"{COGNITO_DOMAIN}/oauth2/token",
        data=urllib.parse.urlencode({
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
            "code_verifier": verifier,
        }).encode(),
        headers={"Authorization": f"Basic {basic}", "Content-Type": "application/x-www-form-urlencoded"},
    ))

    info = _call(urllib.request.Request(
        f"{COGNITO_DOMAIN}/oauth2/userInfo",
        headers={"Authorization": f"Bearer {tokens['access_token']}"},
    ))

    # accounts are keyed by sub, never by email (see #77)
    if not info.get("sub"):
        raise ValueError("Cognito's answer had no sub")

    return {"sub": info["sub"], "email": info.get("email", "")}
