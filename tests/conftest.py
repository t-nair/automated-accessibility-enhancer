import contextlib
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import generate_test_pptx


# the sample decks in "Error Test PPTX" are not tracked, so build them fresh before any test runs
def pytest_sessionstart(session):
    with contextlib.redirect_stdout(io.StringIO()):
        generate_test_pptx.main()
