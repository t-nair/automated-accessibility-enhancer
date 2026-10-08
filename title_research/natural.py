"""Score methods on real slides that have NO title placeholder (hand-labeled in labels_natural.json).

These are the slides the current rule struggles with: titles typed into text boxes or body
placeholders. Cues stay visible here (this is what the tool really sees), but every
non-baseline method is only given the text shapes that are not footer/date/slide-number
placeholders. Learned models are trained on titled slides from OTHER decks (cues hidden,
all perturbations incl. 'no title').
"""
import json
import random
import time
from pathlib import Path

import numpy as np

import methods as M
import perturb as P
from evaluate import load
from features import add_context
from robustness import FNS, NAMES
from title_detection import THRESHOLD

HERE = Path(__file__).parent


def natural_set(slides):
    labels = {k: v for k, v in json.load(open(HERE / "labels_natural.json", encoding="utf-8")).items() if not k.startswith("_")}
    out = []
    for s in slides:
        v = labels.get(f"{s['deck']}#{s['slide_no']}")
        if v is not None and v != -2 and not s["gold"]:
            out.append(dict(s, label=v))
    return out


def without_furniture(method):
    def run(recs):
        keep = [i for i, r in enumerate(recs) if not r["furniture"]]
        if not keep:
            return None
        sub = add_context([dict(recs[i]) for i in keep])
        p = method(sub)
        return None if p is None else keep[p]
    return run


def train_models(slides, exclude_decks, n=1000, kinds=("logreg", "gbm"), clean_weight=1):
    """clean_weight = how many times the unperturbed slides appear next to one copy of each perturbation."""
    pool = [s for s in slides if s["gold"] and s["deck"] not in exclude_decks]
    random.Random(0).shuffle(pool)
    hidden = [dict(s, recs=M.hide(s["recs"])) for s in pool[:n]]
    parts = [hidden] * clean_weight + [P.apply_all(hidden, FNS[name], seed=k + 1) for k, name in enumerate(NAMES)]
    X = np.vstack([M.build_xy([s])[0] for part in parts for s in part])
    y = np.concatenate([M.build_xy([s])[1] for part in parts for s in part])
    return {k: M.LearnedTitle(k, threshold=THRESHOLD).fit_xy(X, y) for k in kinds}


def score(fn, nat):
    t = time.time()
    hits_t = hits_n = n_t = n_n = 0
    wrong = []
    for s in nat:
        p = fn(s["recs"])
        if s["label"] >= 0:
            n_t += 1
            ok = p == s["label"]
            hits_t += ok
        else:
            n_n += 1
            ok = p is None
            hits_n += ok
        if not ok:
            wrong.append((s["deck"], s["slide_no"], s["label"], p))
    return {"title_acc": hits_t / n_t, "none_acc": hits_n / max(n_n, 1), "n_titled": n_t, "n_none": n_n,
            "overall": (hits_t + hits_n) / (n_t + n_n), "sec_per_slide": (time.time() - t) / len(nat), "wrong": wrong}


def show(results):
    print(f"{'method':32}{'title-acc':>10}{'no-title-acc':>14}{'overall':>9}{'s/slide':>9}")
    for name, r in results.items():
        print(f"{name:32}{r['title_acc']:10.3f}{r['none_acc']:14.3f}{r['overall']:9.3f}{r['sec_per_slide']:9.3f}")


if __name__ == "__main__":
    slides = load()
    nat = natural_set(slides)
    print(f"{len(nat)} natural slides: {sum(s['label'] >= 0 for s in nat)} titled, {sum(s['label'] < 0 for s in nat)} without a title")
    models = train_models(slides, {s["deck"] for s in nat})
    methods = {
        "current rule (ph/name)": M.current_rule,
        "first in z-order": without_furniture(M.first_in_order),
        "topmost": without_furniture(M.topmost),
        "largest font": without_furniture(M.largest_font),
        "hybrid hand-weighted": without_furniture(M.hybrid_rules),
        "learned logreg": without_furniture(models["logreg"]),
        "learned gbm": without_furniture(models["gbm"]),
    }
    results = {k: score(fn, nat) for k, fn in methods.items()}
    show(results)
    for k in ("hybrid hand-weighted", "learned gbm"):
        print(k, "misses:", results[k]["wrong"][:25])
