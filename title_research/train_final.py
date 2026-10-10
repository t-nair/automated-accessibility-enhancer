"""Train the shipped title model on everything we have and write ../title_model.joblib.

Training data: every slide with a real title placeholder (cues hidden, plus all perturbed copies
from perturb.py) and the hand-labeled natural slides, including the ones with no title.
natural.py scores a model that did NOT see those natural decks; this one does, so do not quote
this model's accuracy on them.
"""
import sys
from pathlib import Path

import joblib
import numpy as np

import methods as M
import natural as N
import perturb as P
from evaluate import load
from robustness import FNS, NAMES

OUT = Path(__file__).resolve().parent.parent / "title_model.joblib"

if __name__ == "__main__":
    slides = load()
    nat = N.natural_set(slides)
    gold = [dict(s, recs=M.hide(s["recs"])) for s in slides if s["gold"]]
    parts = [gold] + [P.apply_all(gold, FNS[n], seed=k + 1) for k, n in enumerate(NAMES)]
    natural = [dict(s, recs=M.hide(s["recs"]), gold=[s["label"]] if s["label"] >= 0 else []) for s in nat]
    # real text-box titles are the case we care about most, so they count several times
    parts += [natural] * 3 + [P.apply_all(natural, P.jitter, seed=100 + k) for k in range(3)]
    X = np.vstack([M.build_xy([s])[0] for part in parts for s in part])
    y = np.concatenate([M.build_xy([s])[1] for part in parts for s in part])
    model = M.LearnedTitle("gbm").fit_xy(X, y).model
    joblib.dump({"model": model}, OUT, compress=3)
    print(f"{len(y)} shapes ({int(y.sum())} titles) -> {OUT} ({OUT.stat().st_size // 1024} KB)")
