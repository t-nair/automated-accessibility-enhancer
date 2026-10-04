"""Score every cheap title classifier on the labeled slides.

Set A ("hidden"): slides whose real title placeholder gives us gold; the placeholder type
and shape name are hidden from every method, i.e. the author typed the title in a text box.
Learned methods are scored with GroupKFold over decks, so no deck is in both train and test.
"""
import pickle
import sys
from collections import defaultdict
from pathlib import Path

import numpy as np
from sklearn.model_selection import GroupKFold

import methods as M

HERE = Path(__file__).parent


def load():
    return pickle.load(open(HERE / "data" / "slides.pkl", "rb"))


def acc(pred_fn, slides, transform=lambda r: r):
    hits = 0
    for s in slides:
        p = pred_fn(transform(s["recs"]))
        hits += p is not None and p in s["gold"]
    return hits / len(slides)


def cv_predictions(slides, make, folds=5):
    """Out-of-fold predicted index per slide (deck-grouped)."""
    hidden = [dict(s, recs=M.hide(s["recs"])) for s in slides]
    groups = [s["deck"] for s in slides]
    preds = [None] * len(slides)
    for tr, te in GroupKFold(folds).split(hidden, groups=groups):
        model = make().fit([hidden[i] for i in tr])
        for i in te:
            preds[i] = model(hidden[i]["recs"])
    return preds


def report(slides):
    by_src = defaultdict(list)
    for s in slides:
        by_src["all"].append(s)
        by_src[s["source"] if s["source"] in ("own", "zenodo") else "fixtures"].append(s)
    names = list(by_src)
    print(f"{'method (cues hidden)':34}" + "".join(f"{n+' ('+str(len(by_src[n]))+')':>16}" for n in names))
    rows = {
        "current rule (ph/name)": lambda sl: acc(M.current_rule, sl, M.hide),
        "first in z-order": lambda sl: acc(M.first_in_order, sl, M.hide),
        "topmost": lambda sl: acc(M.topmost, sl, M.hide),
        "largest font": lambda sl: acc(M.largest_font, sl, M.hide),
        "hybrid hand-weighted": lambda sl: acc(M.hybrid_rules, sl, M.hide),
    }
    for kind in ("logreg", "gbm"):
        preds = cv_predictions(slides, lambda k=kind: M.LearnedTitle(k))
        pmap = {id(s): p for s, p in zip(slides, preds)}
        rows[f"learned {kind} (deck CV)"] = lambda sl, pmap=pmap: np.mean(
            [pmap[id(s)] is not None and pmap[id(s)] in s["gold"] for s in sl])
    for label, fn in rows.items():
        print(f"{label:34}" + "".join(f"{fn(by_src[n]):16.3f}" for n in names))


if __name__ == "__main__":
    report([s for s in load() if s["gold"]])
