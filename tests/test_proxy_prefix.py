import os
import sys

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module


def get_home(prefix=None):
    headers = {"X-Forwarded-Prefix": prefix} if prefix else {}
    return app_module.app.test_client().get("/", headers=headers).get_data(as_text=True)


# The project site proxies this app under /app and sends X-Forwarded-Prefix. Without that,
# url_for() builds /static/style.css, the landing site answers that path, and the proxied
# page loads unstyled with a dead nav and upload form.
def test_the_site_prefix_reaches_generated_urls():
    page = get_home(app_module.PROXY_PREFIX)

    assert 'href="/app/static/style.css"' in page
    assert 'action="/app/upload"' in page


def test_without_the_header_urls_stay_at_the_root():
    page = get_home()

    assert 'href="/static/style.css"' in page
    assert 'action="/upload"' in page


# Render is reachable directly, so the header is attacker-controlled. Honoring any value
# would let a caller decide the prefix every link on the page is built under.
def test_an_unknown_prefix_is_ignored():
    page = get_home("/evil")

    assert "/evil" not in page
    assert 'href="/static/style.css"' in page
