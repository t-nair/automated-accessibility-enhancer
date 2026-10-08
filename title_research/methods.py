"""Title classifiers. Every method maps a slide's record list to (index_or_None).

Records come from features.slide_records. `hide(recs)` removes the placeholder/name cues
so a method can be scored as if the author had typed the title into a plain text box.
"""
import sys
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from title_detection import CONFIG, feature_rows  # noqa: E402  (one copy, shared with the app)

TITLE_PH_NAMES = ("TITLE", "CENTER_TITLE")


def hide(recs):
    return [dict(r, ph=None, name="", furniture=False) for r in recs]


# ---------------------------------------------------------------- deterministic rules
def current_rule(recs):
    """What z_reorder.is_title does today: title placeholder, else 'title' in the shape name."""
    for i, r in enumerate(recs):
        if r["ph"] in TITLE_PH_NAMES:
            return i
        if "title" in r["name"].lower() and "subtitle" not in r["name"].lower():
            return i
    return None


def first_in_order(recs):
    """Text-based/structural: whatever comes first in z-order (= reading order)."""
    return next((i for i, r in enumerate(recs) if not r["digits_only"]), None)


def topmost(recs):
    """Position-based: highest text shape (ties: leftmost)."""
    ok = [i for i, r in enumerate(recs) if not r["digits_only"]]
    return min(ok, key=lambda i: (round(recs[i]["y"], 2), recs[i]["x"])) if ok else None


def largest_font(recs):
    """Typography-based: biggest effective font (ties: topmost)."""
    ok = [i for i, r in enumerate(recs) if not r["digits_only"]]
    return max(ok, key=lambda i: (recs[i]["pt"], -recs[i]["y"])) if ok else None


def hybrid_score(r):
    """Hand-weighted blend of the cues a person uses: big, high up, short, not a sentence."""
    s = 3.0 * r["pt_rel"] + 2.0 * (1 - min(max(r["y"], 0), 1))
    s += 1.0 if r["n_words"] <= 12 else (0.0 if r["n_words"] <= 25 else -1.5)
    s -= 1.5 * r["end_punct"] + 3.0 * r["digits_only"] + 2.0 * r["bullets"]
    s -= 2.0 * (r["n_paras"] >= 3) + 2.0 * (r["y"] > 0.8)
    s += 0.5 * r["bold"]
    if r["ph"] in TITLE_PH_NAMES:
        s += 20
    if r["furniture"]:
        s -= 20
    if r["ph"] == "SUBTITLE":
        s -= 5
    if "title" in r["name"].lower() and "subtitle" not in r["name"].lower():
        s += 5
    return s


def hybrid_rules(recs, threshold=2.0):
    if not recs:
        return None
    scores = [hybrid_score(r) for r in recs]
    best = int(np.argmax(scores))
    return best if scores[best] >= threshold else None


# ---------------------------------------------------------------- learned (structure only)
def build_xy(slides):
    X, y = [], []
    for s in slides:
        recs = s["recs"]
        X.append(feature_rows(recs))
        y.extend(int(i in s["gold"]) for i in range(len(recs)))
    return np.vstack(X), np.array(y)


class LearnedTitle:
    """Per-shape probability of being the title, argmax per slide, None below `threshold`."""

    def __init__(self, kind="gbm", threshold=0.0):
        from sklearn.ensemble import HistGradientBoostingClassifier
        from sklearn.linear_model import LogisticRegression
        from sklearn.pipeline import make_pipeline
        from sklearn.preprocessing import StandardScaler
        self.model = (HistGradientBoostingClassifier(max_depth=4, max_iter=150, learning_rate=0.08, random_state=0)
                      if kind == "gbm" else
                      make_pipeline(StandardScaler(), LogisticRegression(max_iter=2000, C=1.0)))
        self.threshold = threshold

    def fit(self, slides):
        return self.fit_xy(*build_xy(slides))

    def fit_xy(self, X, y):
        self.model.fit(X, y)
        return self

    def proba(self, recs):
        return self.model.predict_proba(feature_rows(recs))[:, 1]

    def __call__(self, recs):
        if not recs:
            return None
        p = self.proba(recs)
        best = int(np.argmax(p))
        return best if p[best] >= self.threshold else None
