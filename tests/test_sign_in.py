import base64
import hashlib
import os
import sys
import urllib.parse

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module
import auth
import storage_local as storage_module
from test_app import MY_OWNER_ID, SOMEONE_ELSE, add_submission, client, make_upload  # noqa: F401 (client is a fixture)

COGNITO = "https://example-pool.auth.us-east-1.amazoncognito.com"
PUBLIC_URL = "https://accessibilityenhancer.com/app"
ACCOUNT = {"sub": "cognito-sub-1234", "email": "prof@uw.edu"}


@pytest.fixture
def cognito(monkeypatch):
    monkeypatch.setattr(auth, "ENABLED", True)
    monkeypatch.setattr(auth, "COGNITO_DOMAIN", COGNITO)
    monkeypatch.setattr(auth, "COGNITO_CLIENT_ID", "client-id")
    monkeypatch.setattr(auth, "PUBLIC_URL", PUBLIC_URL)

    calls = []

    def pretend_cognito(code, verifier, redirect_uri):
        calls.append({"code": code, "verifier": verifier, "redirect_uri": redirect_uri})
        return dict(ACCOUNT)

    monkeypatch.setattr(auth, "user_for_code", pretend_cognito)
    return calls


def flashes(client):
    with client.session_transaction() as browser_session:
        return " ".join(message for _, message in browser_session.get("_flashes", []))


def query_of(location):
    return dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(location).query))


def start_sign_in(client, path="/auth/sign-in"):
    response = client.get(path)
    return query_of(response.headers["Location"])


# --- without Cognito, nothing changes ---

def test_without_cognito_there_is_no_sign_in_link(client):
    assert b"Sign in" not in client.get("/").data


def test_without_cognito_the_sign_in_routes_go_home(client):
    assert client.get("/auth/sign-in").headers["Location"].endswith("/")
    assert client.get("/auth/callback?code=x&state=y").headers["Location"].endswith("/")


def test_without_cognito_the_site_is_open(client):
    assert client.get("/").status_code == 200
    assert client.get("/status").status_code == 200


# --- going to Cognito ---

def test_sign_in_sends_the_person_to_cognito_with_pkce(client, cognito):
    response = client.get("/auth/sign-in")
    location = response.headers["Location"]
    query = query_of(location)

    assert location.startswith(COGNITO + "/oauth2/authorize?")
    assert query["response_type"] == "code"
    assert query["client_id"] == "client-id"
    assert query["redirect_uri"] == PUBLIC_URL + "/auth/callback"
    assert query["code_challenge_method"] == "S256"

    with client.session_transaction() as browser_session:
        verifier = browser_session["sign_in_verifier"]
        assert browser_session["sign_in_state"] == query["state"]

    expected = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
    assert query["code_challenge"] == expected


def test_create_account_opens_the_sign_up_form(client, cognito):
    location = client.get("/auth/sign-in?create=1").headers["Location"]

    assert location.startswith(COGNITO + "/signup?")


def test_the_challenge_is_url_safe_with_no_padding():
    challenge = auth.code_challenge("a" * 64)

    assert len(challenge) == 43
    assert "=" not in challenge and "+" not in challenge and "/" not in challenge


# --- coming back from Cognito ---

def test_signing_in_moves_this_browsers_files_into_the_account(client, cognito):
    add_submission("mine0001", owner_id=MY_OWNER_ID)
    add_submission("their001", owner_id=SOMEONE_ELSE)
    state = start_sign_in(client)["state"]

    response = client.get(f"/auth/callback?code=the-code&state={state}")

    assert response.headers["Location"].endswith("/status")
    assert storage_module.get_submission("mine0001")["owner_id"] == ACCOUNT["sub"]
    assert storage_module.get_submission("their001")["owner_id"] == SOMEONE_ELSE

    with client.session_transaction() as browser_session:
        assert browser_session["owner_id"] == ACCOUNT["sub"]
        assert browser_session["email"] == ACCOUNT["email"]
        assert "sign_in_state" not in browser_session

    page = client.get("/status").data
    assert b"Signed in as prof@uw.edu" in page
    assert b"Lecture.pptx" in page


def test_signing_in_over_someone_elses_account_does_not_take_their_files(client, cognito, monkeypatch):
    # person A signed in on a shared computer and never signed out
    with client.session_transaction() as browser_session:
        browser_session["owner_id"] = "account-a"
        browser_session["email"] = "a@uw.edu"

    add_submission("aaaa0001", owner_id="account-a")
    state = start_sign_in(client)["state"]

    # person B signs in from the same browser
    monkeypatch.setattr(auth, "user_for_code", lambda code, verifier, redirect_uri: {"sub": "account-b", "email": "b@uw.edu"})
    client.get(f"/auth/callback?code=the-code&state={state}")

    assert storage_module.get_submission("aaaa0001")["owner_id"] == "account-a"

    with client.session_transaction() as browser_session:
        assert browser_session["owner_id"] == "account-b"


def test_the_code_is_exchanged_with_the_public_redirect_uri_and_the_verifier(client, cognito):
    state = start_sign_in(client)["state"]

    with client.session_transaction() as browser_session:
        verifier = browser_session["sign_in_verifier"]

    client.get(f"/auth/callback?code=the-code&state={state}")

    assert cognito == [{"code": "the-code", "verifier": verifier, "redirect_uri": PUBLIC_URL + "/auth/callback"}]


def test_a_wrong_state_signs_nobody_in(client, cognito):
    start_sign_in(client)

    client.get("/auth/callback?code=the-code&state=not-the-one")

    assert "expired" in flashes(client)
    assert cognito == []

    with client.session_transaction() as browser_session:
        assert browser_session["owner_id"] == MY_OWNER_ID
        assert "email" not in browser_session


def test_a_callback_nobody_asked_for_signs_nobody_in(client, cognito):
    client.get("/auth/callback?code=the-code&state=anything")

    assert cognito == []


def test_a_cancelled_sign_in_leaves_the_files_where_they_were(client, cognito):
    add_submission("mine0001", owner_id=MY_OWNER_ID)
    state = start_sign_in(client)["state"]

    client.get(f"/auth/callback?error=access_denied&state={state}")

    assert "You were not signed in" in flashes(client)
    assert storage_module.get_submission("mine0001")["owner_id"] == MY_OWNER_ID


def test_a_problem_talking_to_cognito_is_explained(client, cognito, monkeypatch):
    def broken(code, verifier, redirect_uri):
        raise OSError("connection reset")

    monkeypatch.setattr(auth, "user_for_code", broken)
    state = start_sign_in(client)["state"]

    client.get(f"/auth/callback?code=the-code&state={state}")

    assert "Sign-in did not work" in flashes(client)


# --- signing out ---

def test_signing_out_forgets_the_account_and_signs_out_of_cognito(client, cognito):
    state = start_sign_in(client)["state"]
    client.get(f"/auth/callback?code=the-code&state={state}")

    # signing in starts a fresh session, so take its new form token
    with client.session_transaction() as browser_session:
        token = browser_session.get("csrf_token") or "fresh-token"
        browser_session["csrf_token"] = token

    response = client.post_without_token("/auth/sign-out", data={"csrf_token": token})
    query = query_of(response.headers["Location"])

    assert response.headers["Location"].startswith(COGNITO + "/logout?")
    assert query == {"client_id": "client-id", "logout_uri": PUBLIC_URL + "/"}

    with client.session_transaction() as browser_session:
        assert "email" not in browser_session
        assert "owner_id" not in browser_session


def test_signing_out_needs_the_form_token(client, cognito):
    state = start_sign_in(client)["state"]
    client.get(f"/auth/callback?code=the-code&state={state}")

    response = client.post_without_token("/auth/sign-out", data={"csrf_token": "wrong"})

    assert response.status_code == 400


# --- every page needs an account ---

@pytest.mark.parametrize("path", ["/", "/status", "/details/mine0001", "/progress/mine0001",
                                  "/report/mine0001", "/download/mine0001"])
def test_signed_out_visitors_are_sent_to_sign_in(client, cognito, path):
    add_submission("mine0001", owner_id=MY_OWNER_ID)

    response = client.get(path)

    assert response.status_code == 302
    assert response.headers["Location"].endswith("/auth/sign-in")


def test_a_signed_out_upload_is_sent_to_sign_in_and_saves_nothing(client, cognito):
    response = client.post("/upload", data={"presentation": make_upload("control_no_issues.pptx")},
                           content_type="multipart/form-data")

    assert response.headers["Location"].endswith("/auth/sign-in")
    assert client.queued == []


def test_the_health_check_needs_no_account(client, cognito):
    response = client.get("/health")

    assert response.status_code == 200
    assert response.data == b"ok"


def test_signed_in_visitors_can_use_the_site(client, cognito):
    state = start_sign_in(client)["state"]
    client.get(f"/auth/callback?code=the-code&state={state}")

    with client.session_transaction() as browser_session:
        browser_session["csrf_token"] = "token-after-sign-in"

    assert client.get("/").status_code == 200

    client.post_without_token(
        "/upload", data={"csrf_token": "token-after-sign-in", "presentation": make_upload("control_no_issues.pptx")},
        content_type="multipart/form-data",
    )

    assert len(client.queued) == 1


# --- storage ---

def test_moving_submissions_leaves_other_people_alone(client):
    add_submission("mine0001", owner_id="browser-a")
    add_submission("mine0002", owner_id="browser-a")
    add_submission("their001", owner_id="browser-b")

    storage_module.move_submissions("browser-a", "account-1")

    assert sorted(s["id"] for s in storage_module.get_submissions_for_owner("account-1")) == ["mine0001", "mine0002"]
    assert storage_module.get_submission("their001")["owner_id"] == "browser-b"
