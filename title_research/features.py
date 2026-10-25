"""Feature extraction lives in the repo root (title_detection.py) so the app and the experiments share one copy."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from title_detection import *  # noqa: E402,F401,F403
