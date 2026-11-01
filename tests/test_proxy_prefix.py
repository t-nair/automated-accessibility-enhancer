import os
import sys

# make sure the tests can import the project files from the folder above
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as app_module


# The project site proxies this app under /app and sends X-Forwarded-Prefix. Without
# ProxyFix, url_for() builds /static/style.css, the landing site answers that path, and
# the proxied page loads unstyled with a dead nav and upload form.
def test_forwarded_prefix_reaches_generated_urls():
    page = app_module.app.test_client().get("/", headers={"X-Forwarded-Prefix": "/app"}).get_data(as_text=True)

    assert 'href="/app/static/style.css"' in page
    assert 'action="/app/upload"' in page


def test_without_the_header_urls_stay_at_the_root():
    page = app_module.app.test_client().get("/").get_data(as_text=True)

    assert 'href="/static/style.css"' in page
    assert 'action="/upload"' in page
