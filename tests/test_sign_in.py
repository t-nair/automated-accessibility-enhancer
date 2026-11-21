import base64
import hashlib
import hmac
import os
import sys
import urllib.parse

import pytest
from botocore.stub import Stubber

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import auth
import storage_local as storage_module
from test_app import MY_OWNER_ID, SOMEONE_ELSE, add_submission, client, make_upload  # noqa: F401 (client is a fixture)

COGNITO = "https://example-pool.auth.us-east-1.amazoncognito.com"
PUBLIC_URL = "https://accessibilityenhancer.com/app"
ACCOUNT = {"sub": "cognito-sub-1234", "email": "prof@uw.edu"}


class Calls(list):
    """What the app asked Cognito for. Put an exception in errors[name] to make that call fail."""

    def __init__(self):
        super().__init__()
        self.errors = {}


# Cognito turned on, with every call auth makes to it replaced by a pretend one
@pytest.fixture
def cognito(monkeypatch):
    monkeypatch.setattr(auth, "ENABLED", True)
    monkeypatch.setattr(auth, "GOOGLE_ENABLED", True)
    monkeypatch.setattr(auth, "COGNITO_DOMAIN", COGNITO)
    monkeypatch.setattr(auth, "COGNITO_CLIENT_ID", "client-id")
    monkeypatch.setattr(auth, "PUBLIC_URL", PUBLIC_URL)

    calls = Calls()

    def pretend(name, answer=None):
        def call(*args):
            calls.append((name, *args))
            if name in calls.errors:
                raise calls.errors[name]
            return answer
        return call

    for name in ("resend_code", "forgot_password", "reset_password"):
        monkeypatch.setattr(auth, name, pretend(name))

    monkeypatch.setattr(auth, "sign_up", pretend("sign_up", "sign-up-session"))
    monkeypatch.setattr(auth, "confirm_sign_up", pretend("confirm_sign_up", "sign-in-session"))
    monkeypatch.setattr(auth, "user_for_sign_up_session", pretend("user_for_sign_up_session", dict(ACCOUNT)))
    monkeypatch.setattr(auth, "user_for_password", pretend("user_for_password", dict(ACCOUNT)))
    monkeypatch.setattr(auth, "user_for_code", pretend("user_for_code", dict(ACCOUNT)))

    return calls


def flashes(client):
    with client.session_transaction() as browser_session:
        return " ".join(message for _, message in browser_session.get("_flashes", []))


def signed_in_email(client):
    with client.session_transaction() as browser_session:
        return browser_session.get("email")


def post(client, path, **fields):
    return client.post(path, data=fields)


def sign_in(client, email="prof@uw.edu", password="Correct-horse-1"):
    return post(client, "/sign-in", email=email, password=password)


def query_of(location):
    return dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(location).query))


def start_google(client):
    return query_of(client.get("/auth/google").headers["Location"])["state"]


# --- without Cognito, nothing changes ---

def test_without_cognito_the_site_is_open(client):
    assert client.get("/").status_code == 200
    assert client.get("/status").status_code == 200


def test_without_cognito_the_account_pages_go_home(client):
    for path in ("/sign-in", "/sign-up", "/forgot-password", "/auth/google", "/auth/callback?code=x&state=y"):
        assert client.get(path).headers["Location"].endswith("/"), path


def test_without_cognito_there_is_no_sign_out(client):
    assert b"Sign out" not in client.get("/").data


# --- every page needs an account ---

@pytest.mark.parametrize("path", ["/", "/status", "/details/mine0001", "/progress/mine0001",
                                  "/report/mine0001", "/download/mine0001"])
def test_signed_out_visitors_are_sent_to_the_sign_in_page(client, cognito, path):
    add_submission("mine0001", owner_id=MY_OWNER_ID)

    response = client.get(path)

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/sign-in")


def test_a_signed_out_upload_is_sent_to_sign_in_and_saves_nothing(client, cognito):
    response = client.post("/upload", data={"presentation": make_upload("control_no_issues.pptx")},
                           content_type="multipart/form-data")

    assert response.headers["Location"].endswith("/sign-in")
    assert client.queued == []


def test_the_health_check_needs_no_account(client, cognito):
    response = client.get("/health")

    assert response.status_code == 200
    assert response.data == b"ok"


def test_signed_in_visitors_can_use_the_site(client, cognito):
    sign_in(client)

    with client.session_transaction() as browser_session:
        browser_session["csrf_token"] = "token-after-sign-in"

    assert client.get("/").status_code == 200

    client.post_without_token(
        "/upload", data={"csrf_token": "token-after-sign-in", "presentation": make_upload("control_no_issues.pptx")},
        content_type="multipart/form-data",
    )

    assert len(client.queued) == 1


def test_signed_in_visitors_skip_the_account_pages(client, cognito):
    sign_in(client)

    assert client.get("/sign-in").headers["Location"].endswith("/")


# --- the sign-in page ---

def test_the_sign_in_page_is_part_of_the_site(client, cognito):
    page = client.get("/sign-in").data

    assert b'<form method="POST" action="/sign-in">' in page
    assert b'autocomplete="current-password"' in page
    assert b"Continue with Google" in page
    assert b"Check Status" not in page  # no links that would only come back here


def test_the_google_button_is_hidden_until_google_is_set_up(client, cognito, monkeypatch):
    monkeypatch.setattr(auth, "GOOGLE_ENABLED", False)

    assert b"Continue with Google" not in client.get("/sign-in").data
    assert client.get("/auth/google").headers["Location"].endswith("/sign-in")


def test_signing_in_with_a_password(client, cognito):
    response = sign_in(client, email="  prof@uw.edu ")

    assert response.headers["Location"].endswith("/")
    assert ("user_for_password", "prof@uw.edu", "Correct-horse-1") in cognito
    assert signed_in_email(client) == "prof@uw.edu"

    with client.session_transaction() as browser_session:
        assert browser_session["owner_id"] == ACCOUNT["sub"]

    assert b"Signed in as prof@uw.edu" in client.get("/").data


def test_a_wrong_password_says_so_and_signs_nobody_in(client, cognito):
    cognito.errors["user_for_password"] = auth.SignInError(auth.MESSAGES["NotAuthorizedException"], "NotAuthorizedException")

    response = sign_in(client)

    assert response.status_code == 400
    assert b"did not match an account" in response.data
    assert b'value="prof@uw.edu"' in response.data  # the email is kept, the password is not
    assert signed_in_email(client) is None


def test_an_unconfirmed_account_is_sent_a_new_code(client, cognito):
    cognito.errors["user_for_password"] = auth.SignInError(auth.MESSAGES["UserNotConfirmedException"], "UserNotConfirmedException")

    response = sign_in(client)

    assert response.headers["Location"].endswith("/confirm-email")
    assert ("resend_code", "prof@uw.edu") in cognito


def test_cognito_being_down_is_explained(client, cognito):
    cognito.errors["user_for_password"] = OSError("connection reset")

    response = sign_in(client)

    assert response.status_code == 400
    assert b"did not answer" in response.data


def test_the_password_is_never_kept_in_the_session(client, cognito):
    sign_in(client)

    with client.session_transaction() as browser_session:
        assert "Correct-horse-1" not in repr(dict(browser_session))


def test_signing_in_moves_this_browsers_files_into_the_account(client, cognito):
    add_submission("mine0001", owner_id=MY_OWNER_ID)
    add_submission("their001", owner_id=SOMEONE_ELSE)

    sign_in(client)

    assert storage_module.get_submission("mine0001")["owner_id"] == ACCOUNT["sub"]
    assert storage_module.get_submission("their001")["owner_id"] == SOMEONE_ELSE
    assert b"Lecture.pptx" in client.get("/status").data


def test_nobody_can_sign_in_over_someone_still_signed_in(client, cognito):
    # person A signed in on a shared computer and never signed out
    with client.session_transaction() as browser_session:
        browser_session["owner_id"] = "account-a"
        browser_session["email"] = "a@uw.edu"
        browser_session["sign_in_state"] = "s"
        browser_session["sign_in_verifier"] = "v"

    add_submission("aaaa0001", owner_id="account-a")

    # person B tries to sign in, by password or coming back from Google
    assert sign_in(client).headers["Location"].endswith("/")
    assert client.get("/auth/callback?code=the-code&state=s").headers["Location"].endswith("/")

    assert cognito == []
    assert storage_module.get_submission("aaaa0001")["owner_id"] == "account-a"
    assert signed_in_email(client) == "a@uw.edu"


def test_a_stale_form_on_the_sign_in_page_comes_back_to_it(client, cognito):
    response = client.post_without_token("/sign-in", data={"csrf_token": "wrong", "email": "a@uw.edu", "password": "x"})

    assert response.headers["Location"].endswith("/sign-in")
    assert "expired" in flashes(client)
    assert cognito == []


# --- creating an account ---

def test_creating_an_account_then_confirming_the_email_signs_in(client, cognito):
    response = post(client, "/sign-up", email="new@uw.edu", password="Correct-horse-1")

    assert response.headers["Location"].endswith("/confirm-email")
    assert ("sign_up", "new@uw.edu", "Correct-horse-1") in cognito

    page = client.get("/confirm-email").data
    assert b"We sent a code to new@uw.edu" in page
    assert b'value="new@uw.edu"' in page

    response = post(client, "/confirm-email", email="new@uw.edu", code=" 123456 ")

    # no second sign-in: the session from sign-up carries through, not the password
    assert response.headers["Location"].endswith("/")
    assert ("confirm_sign_up", "new@uw.edu", "123456", "sign-up-session") in cognito
    assert ("user_for_sign_up_session", "new@uw.edu", "sign-in-session") in cognito
    assert signed_in_email(client) == ACCOUNT["email"]

    with client.session_transaction() as browser_session:
        assert "sign_up_session" not in browser_session


def test_confirming_with_another_email_does_not_use_the_sign_up_session(client, cognito):
    post(client, "/sign-up", email="new@uw.edu", password="Correct-horse-1")

    post(client, "/confirm-email", email="other@uw.edu", code="123456")

    assert ("confirm_sign_up", "other@uw.edu", "123456", None) in cognito


def test_confirming_later_still_works_with_a_password_sign_in(client, cognito):
    # another browser, or the next day: no sign-up session to carry through
    cognito.errors["user_for_sign_up_session"] = auth.SignInError("Please sign in.", "NotAuthorizedException")

    response = post(client, "/confirm-email", email="new@uw.edu", code="123456")

    assert response.headers["Location"].endswith("/sign-in")
    assert "Your email is confirmed" in flashes(client)
    assert signed_in_email(client) is None


def test_a_weak_password_shows_cognitos_rule(client, cognito):
    cognito.errors["sign_up"] = auth.SignInError("Password must have symbol characters", "InvalidPasswordException")

    response = post(client, "/sign-up", email="new@uw.edu", password="weak")

    assert response.status_code == 400
    assert b"Password must have symbol characters" in response.data


def test_a_wrong_code_says_so(client, cognito):
    cognito.errors["confirm_sign_up"] = auth.SignInError(auth.MESSAGES["CodeMismatchException"], "CodeMismatchException")

    response = post(client, "/confirm-email", email="new@uw.edu", code="000000")

    assert response.status_code == 400
    assert b"That code is not right" in response.data


def test_sending_a_new_code(client, cognito):
    response = post(client, "/confirm-email/resend", email="new@uw.edu")

    assert response.headers["Location"].endswith("/confirm-email")
    assert ("resend_code", "new@uw.edu") in cognito


# --- forgotten passwords ---

def test_resetting_a_forgotten_password(client, cognito):
    response = post(client, "/forgot-password", email="prof@uw.edu")

    assert response.headers["Location"].endswith("/reset-password")
    assert ("forgot_password", "prof@uw.edu") in cognito

    response = post(client, "/reset-password", email="prof@uw.edu", code="123456", password="New-horse-2")

    assert response.headers["Location"].endswith("/sign-in")
    assert ("reset_password", "prof@uw.edu", "123456", "New-horse-2") in cognito


def test_the_forgot_password_page_does_not_say_who_has_an_account(client, cognito):
    post(client, "/forgot-password", email="nobody@uw.edu")

    assert "If there is an account for nobody@uw.edu" in flashes(client)


# --- Google ---

def test_google_goes_straight_to_google_with_pkce(client, cognito):
    location = client.get("/auth/google").headers["Location"]
    query = query_of(location)

    assert location.startswith(COGNITO + "/oauth2/authorize?")
    assert query["identity_provider"] == "Google"
    assert query["redirect_uri"] == PUBLIC_URL + "/auth/callback"
    assert query["code_challenge_method"] == "S256"

    with client.session_transaction() as browser_session:
        verifier = browser_session["sign_in_verifier"]
        assert browser_session["sign_in_state"] == query["state"]

    expected = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    assert query["code_challenge"] == expected


def test_the_challenge_is_url_safe_with_no_padding():
    challenge = auth.code_challenge("a" * 64)

    assert len(challenge) == 43
    assert "=" not in challenge and "+" not in challenge and "/" not in challenge


def test_coming_back_from_google_signs_in(client, cognito):
    state = start_google(client)

    with client.session_transaction() as browser_session:
        verifier = browser_session["sign_in_verifier"]

    response = client.get(f"/auth/callback?code=the-code&state={state}")

    assert response.headers["Location"].endswith("/")
    assert ("user_for_code", "the-code", verifier, PUBLIC_URL + "/auth/callback") in cognito
    assert signed_in_email(client) == ACCOUNT["email"]


def test_a_wrong_state_signs_nobody_in(client, cognito):
    start_google(client)

    response = client.get("/auth/callback?code=the-code&state=not-the-one")

    assert response.headers["Location"].endswith("/sign-in")
    assert "expired" in flashes(client)
    assert cognito == []


def test_a_callback_nobody_asked_for_signs_nobody_in(client, cognito):
    client.get("/auth/callback?code=the-code&state=anything")

    assert cognito == []


def test_cancelling_at_google_comes_back_to_sign_in(client, cognito):
    state = start_google(client)

    response = client.get(f"/auth/callback?error=access_denied&state={state}")

    assert response.headers["Location"].endswith("/sign-in")
    assert "You were not signed in" in flashes(client)


def test_a_problem_after_google_is_explained(client, cognito):
    cognito.errors["user_for_code"] = OSError("connection reset")
    state = start_google(client)

    response = client.get(f"/auth/callback?code=the-code&state={state}")

    assert response.headers["Location"].endswith("/sign-in")
    assert "Sign-in did not work" in flashes(client)


# --- signing out ---

def sign_out(client):
    with client.session_transaction() as browser_session:
        browser_session["csrf_token"] = "token-after-sign-in"

    return client.post_without_token("/auth/sign-out", data={"csrf_token": "token-after-sign-in"})


def test_signing_out_after_a_password_stays_on_the_site(client, cognito):
    sign_in(client)

    response = sign_out(client)

    assert response.headers["Location"].endswith("/")
    assert signed_in_email(client) is None


def test_signing_out_after_google_also_signs_out_of_cognito(client, cognito):
    state = start_google(client)
    client.get(f"/auth/callback?code=the-code&state={state}")

    response = sign_out(client)
    query = query_of(response.headers["Location"])

    assert response.headers["Location"].startswith(COGNITO + "/logout?")
    assert query == {"client_id": "client-id", "logout_uri": PUBLIC_URL + "/"}
    assert signed_in_email(client) is None


def test_signing_out_needs_the_form_token(client, cognito):
    sign_in(client)

    response = client.post_without_token("/auth/sign-out", data={"csrf_token": "wrong"})

    assert response.status_code == 400
    assert signed_in_email(client) == "prof@uw.edu"


# --- what auth actually sends to Cognito ---

@pytest.fixture
def cognito_api(monkeypatch):
    monkeypatch.setattr(auth, "COGNITO_DOMAIN", COGNITO)
    monkeypatch.setattr(auth, "COGNITO_CLIENT_ID", "client-id")
    monkeypatch.setattr(auth, "_client_secret", "client-secret")
    monkeypatch.setattr(auth, "_idp", None)

    with Stubber(auth._cognito()) as stubber:
        yield stubber
        stubber.assert_no_pending_responses()


def secret_hash(email):
    digest = hmac.new(b"client-secret", (email + "client-id").encode(), hashlib.sha256).digest()
    return base64.b64encode(digest).decode()


def test_the_client_is_in_the_pools_region_and_unsigned(cognito_api):
    from botocore import UNSIGNED

    client = auth._cognito()

    assert client.meta.region_name == "us-east-1"
    assert client.meta.config.signature_version is UNSIGNED


def test_password_sign_in_sends_the_secret_hash_and_reads_the_account(cognito_api):
    cognito_api.add_response(
        "initiate_auth",
        {"AuthenticationResult": {"AccessToken": "access-token"}},
        {"ClientId": "client-id", "AuthFlow": "USER_PASSWORD_AUTH",
         "AuthParameters": {"USERNAME": "prof@uw.edu", "PASSWORD": "pw", "SECRET_HASH": secret_hash("prof@uw.edu")}},
    )
    cognito_api.add_response(
        "get_user",
        {"Username": "x", "UserAttributes": [{"Name": "sub", "Value": "sub-1"}, {"Name": "email", "Value": "prof@uw.edu"}]},
        {"AccessToken": "access-token"},
    )

    assert auth.user_for_password("prof@uw.edu", "pw") == {"sub": "sub-1", "email": "prof@uw.edu"}


def test_a_challenge_is_not_a_sign_in(cognito_api):
    cognito_api.add_response("initiate_auth", {"ChallengeName": "NEW_PASSWORD_REQUIRED", "Session": "s" * 20})

    with pytest.raises(auth.SignInError) as problem:
        auth.user_for_password("prof@uw.edu", "pw")

    assert problem.value.code == "NEW_PASSWORD_REQUIRED"


def test_cognitos_errors_become_plain_messages(cognito_api):
    cognito_api.add_client_error("initiate_auth", "NotAuthorizedException", "Incorrect username or password.")

    with pytest.raises(auth.SignInError) as problem:
        auth.user_for_password("prof@uw.edu", "pw")

    assert problem.value.code == "NotAuthorizedException"
    assert problem.value.message == auth.MESSAGES["NotAuthorizedException"]


def test_a_password_rule_comes_from_cognitos_message(cognito_api):
    cognito_api.add_client_error("sign_up", "InvalidPasswordException",
                                 "Password did not conform with policy: Password must have symbol characters")

    with pytest.raises(auth.SignInError) as problem:
        auth.sign_up("new@uw.edu", "weak")

    assert problem.value.message == "Password must have symbol characters"


def test_sign_up_sends_the_email_as_username_and_attribute(cognito_api):
    cognito_api.add_response(
        "sign_up", {"UserConfirmed": False, "UserSub": "sub-2"},
        {"ClientId": "client-id", "Username": "new@uw.edu", "SecretHash": secret_hash("new@uw.edu"),
         "Password": "Correct-horse-1", "UserAttributes": [{"Name": "email", "Value": "new@uw.edu"}]},
    )

    auth.sign_up("new@uw.edu", "Correct-horse-1")


def test_signing_in_after_sign_up_uses_the_session_not_a_password(cognito_api):
    cognito_api.add_response(
        "confirm_sign_up", {"Session": "s" * 20 + "-in"},
        {"ClientId": "client-id", "Username": "new@uw.edu", "SecretHash": secret_hash("new@uw.edu"),
         "ConfirmationCode": "123456", "Session": "s" * 20 + "-up"},
    )
    cognito_api.add_response(
        "initiate_auth", {"AuthenticationResult": {"AccessToken": "access-token"}},
        {"ClientId": "client-id", "AuthFlow": "USER_AUTH", "Session": "s" * 20 + "-in",
         "AuthParameters": {"USERNAME": "new@uw.edu", "SECRET_HASH": secret_hash("new@uw.edu")}},
    )
    cognito_api.add_response(
        "get_user", {"Username": "x", "UserAttributes": [{"Name": "sub", "Value": "sub-2"}]},
        {"AccessToken": "access-token"},
    )

    session = auth.confirm_sign_up("new@uw.edu", "123456", "s" * 20 + "-up")

    assert auth.user_for_sign_up_session("new@uw.edu", session) == {"sub": "sub-2", "email": "new@uw.edu"}


def test_an_expired_sign_up_session_still_confirms_the_email(cognito_api):
    cognito_api.add_client_error("confirm_sign_up", "NotAuthorizedException", "Invalid session for the user.")
    cognito_api.add_response(
        "confirm_sign_up", {},
        {"ClientId": "client-id", "Username": "new@uw.edu", "SecretHash": secret_hash("new@uw.edu"),
         "ConfirmationCode": "123456"},
    )

    assert auth.confirm_sign_up("new@uw.edu", "123456", "s" * 20) is None


def test_a_wrong_code_is_not_tried_again_without_the_session(cognito_api):
    cognito_api.add_client_error("confirm_sign_up", "CodeMismatchException")

    with pytest.raises(auth.SignInError) as problem:
        auth.confirm_sign_up("new@uw.edu", "000000", "s" * 20)

    assert problem.value.code == "CodeMismatchException"


def test_an_unexpected_error_is_not_hidden(cognito_api):
    cognito_api.add_client_error("forgot_password", "InternalErrorException")

    with pytest.raises(Exception) as problem:
        auth.forgot_password("prof@uw.edu")

    assert not isinstance(problem.value, auth.SignInError)


# --- storage ---

def test_moving_submissions_leaves_other_people_alone(client):
    add_submission("mine0001", owner_id="browser-a")
    add_submission("mine0002", owner_id="browser-a")
    add_submission("their001", owner_id="browser-b")

    storage_module.move_submissions("browser-a", "account-1")

    assert sorted(s["id"] for s in storage_module.get_submissions_for_owner("account-1")) == ["mine0001", "mine0002"]
    assert storage_module.get_submission("their001")["owner_id"] == "browser-b"
