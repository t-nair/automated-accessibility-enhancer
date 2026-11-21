# Sign-in with Amazon Cognito (#77). When it is turned on, every page needs an
# account (see require_sign_in in app.py).
#
# Email and password happen on the app's own pages, which pass what people type
# to Cognito's API (InitiateAuth, SignUp, ForgotPassword and the rest). Cognito
# still keeps the accounts, hashes the passwords, sends the emails and locks out
# repeated guesses. Those calls need no AWS credentials, only the app client's
# id and secret.
#
# Google cannot be done on our pages, so it is the standard OAuth authorization
# code flow with PKCE, done with the standard library, straight to Google
# (identity_provider skips Cognito's own sign-in page).
#
# Either way the person's id and email come from Cognito over HTTPS with the
# access token (GetUser or userInfo), so no JWT has to be decoded or verified here.
#
# Turned on by the COGNITO_* variables Terraform gives the website. Without them
# (a laptop, or a server of your own) the site is open, with the signed-cookie
# identity alone.

import base64
import hashlib
import hmac
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

# Terraform sets this once a Google client is registered with the pool
GOOGLE_ENABLED = os.environ.get("COGNITO_GOOGLE") == "true"

SCOPES = "openid email profile"

_client_secret = None


# read once, on the first sign-in rather than at startup, so a problem with
# Parameter Store does not stop the container from starting
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


def google_sign_in_url(redirect_uri, state, verifier):
    query = urllib.parse.urlencode({
        "identity_provider": "Google",
        "response_type": "code",
        "client_id": COGNITO_CLIENT_ID,
        "redirect_uri": redirect_uri,
        "scope": SCOPES,
        "state": state,
        "code_challenge": code_challenge(verifier),
        "code_challenge_method": "S256",
    })
    return f"{COGNITO_DOMAIN}/oauth2/authorize?{query}"


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


# --- email and password, through Cognito's API ---

class SignInError(Exception):
    """A problem whose message is fine to show the person. code is Cognito's name for it."""

    def __init__(self, message, code=""):
        super().__init__(message)
        self.message = message
        self.code = code


# what Cognito's errors mean, in words for the person at the form
TOO_MANY_TRIES = "Too many tries. Please wait a few minutes and try again."

MESSAGES = {
    "NotAuthorizedException": "That email and password did not match an account.",
    "UserNotFoundException": "That email and password did not match an account.",
    "UserNotConfirmedException": "Your email is not confirmed yet. We have sent you a new code.",
    "PasswordResetRequiredException": "Your password has to be reset before you can sign in.",
    "UsernameExistsException": "There is already an account with that email. Sign in, or reset your password if you have forgotten it.",
    "CodeMismatchException": "That code is not right. Check the email we sent and try again.",
    "ExpiredCodeException": "That code has expired. Send a new one and try again.",
    "InvalidParameterException": "Check the email address and try again.",
    "LimitExceededException": TOO_MANY_TRIES,
    "TooManyRequestsException": TOO_MANY_TRIES,
    "TooManyFailedAttemptsException": TOO_MANY_TRIES,
}

_idp = None


# These calls are unauthenticated in Cognito's API, so the client sends no AWS
# signature and works without credentials, on a laptop as on ECS. The region is
# the one in the pool's domain (<prefix>.auth.<region>.amazoncognito.com).
def _cognito():
    global _idp

    if _idp is None:
        import boto3
        from botocore import UNSIGNED
        from botocore.config import Config

        region = urllib.parse.urlsplit(COGNITO_DOMAIN).hostname.split(".")[2]
        _idp = boto3.client("cognito-idp", region_name=region, config=Config(signature_version=UNSIGNED))

    return _idp


def _ask_cognito(operation, **params):
    from botocore.exceptions import ClientError

    try:
        return getattr(_cognito(), operation)(**params)
    except ClientError as e:
        code = e.response["Error"]["Code"]

        # Cognito's own message names the rule the password broke, e.g.
        # "Password did not conform with policy: Password must have symbol characters"
        if code == "InvalidPasswordException":
            raise SignInError(e.response["Error"]["Message"].split(":", 1)[-1].strip(), code) from e

        if code in MESSAGES:
            raise SignInError(MESSAGES[code], code) from e

        raise


# Cognito wants this alongside the username whenever the app client has a secret
def _secret_hash(email):
    digest = hmac.new(client_secret().encode(), (email + COGNITO_CLIENT_ID).encode(), hashlib.sha256).digest()
    return base64.b64encode(digest).decode()


def _with_client(email, **params):
    return {"ClientId": COGNITO_CLIENT_ID, "Username": email, "SecretHash": _secret_hash(email), **params}


# gives back {"sub": ..., "email": ...}, like user_for_code
def user_for_password(email, password):
    answer = _ask_cognito(
        "initiate_auth",
        ClientId=COGNITO_CLIENT_ID,
        AuthFlow="USER_PASSWORD_AUTH",
        AuthParameters={"USERNAME": email, "PASSWORD": password, "SECRET_HASH": _secret_hash(email)},
    )

    # The pool has no MFA, so the only challenge is NEW_PASSWORD_REQUIRED, for an
    # account someone made by hand in the console with a temporary password
    if "AuthenticationResult" not in answer:
        raise SignInError("This account was set up with a temporary password and cannot sign in here. "
                          "Please contact us to have it reset.", answer.get("ChallengeName", ""))

    return _user_for_tokens(answer, email)


# Signs in someone who has just confirmed their email, with the session
# ConfirmSignUp gave back instead of their password (Cognito's "sign in after
# sign-up", the USER_AUTH flow). Only works once, within a few minutes.
def user_for_sign_up_session(email, session):
    answer = _ask_cognito(
        "initiate_auth",
        ClientId=COGNITO_CLIENT_ID,
        AuthFlow="USER_AUTH",
        AuthParameters={"USERNAME": email, "SECRET_HASH": _secret_hash(email)},
        Session=session,
    )

    if "AuthenticationResult" not in answer:
        raise SignInError("Please sign in.", answer.get("ChallengeName", ""))

    return _user_for_tokens(answer, email)


def _user_for_tokens(answer, email):
    user = _ask_cognito("get_user", AccessToken=answer["AuthenticationResult"]["AccessToken"])
    attributes = {a["Name"]: a["Value"] for a in user["UserAttributes"]}

    return {"sub": attributes["sub"], "email": attributes.get("email", email)}


# gives back the session to hand to confirm_sign_up, or None
def sign_up(email, password):
    answer = _ask_cognito("sign_up", **_with_client(email, Password=password,
                                                    UserAttributes=[{"Name": "email", "Value": email}]))
    return answer.get("Session")


# A wrong code or too many tries is the person's to fix, so it is shown. Anything
# else with a session (most likely the session expired while they found the
# email) is tried again without one: confirming matters more than skipping the
# sign-in after. Gives back the session for user_for_sign_up_session, or None.
NOT_THE_SESSIONS_FAULT = {"CodeMismatchException", "LimitExceededException",
                          "TooManyRequestsException", "TooManyFailedAttemptsException"}


def confirm_sign_up(email, code, session=None):
    if session:
        try:
            return _ask_cognito("confirm_sign_up", **_with_client(email, ConfirmationCode=code, Session=session)).get("Session")
        except SignInError as e:
            if e.code in NOT_THE_SESSIONS_FAULT:
                raise
        except Exception:
            pass

    _ask_cognito("confirm_sign_up", **_with_client(email, ConfirmationCode=code))
    return None


def resend_code(email):
    _ask_cognito("resend_confirmation_code", **_with_client(email))


# With prevent_user_existence_errors on, Cognito answers the same for an email
# with no account, so this page cannot be used to find out who has one
def forgot_password(email):
    _ask_cognito("forgot_password", **_with_client(email))


def reset_password(email, code, password):
    _ask_cognito("confirm_forgot_password", **_with_client(email, ConfirmationCode=code, Password=password))
