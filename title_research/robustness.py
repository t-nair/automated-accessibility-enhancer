"""How do the methods hold up when the layout gets awkward?

Every labeled slide is rewritten by each perturbation in perturb.py (cues hidden, gold kept).
Rules are scored directly. Learned models are scored with deck-grouped CV in three ways:
  plain    trained on clean slides only
  aug      trained on clean + every perturbation (tests "seen" layout tricks)
  LOPO     trained on clean + all perturbations EXCEPT the one being tested (unseen trick)
"""
import random
import sys
from pathlib import Path

import numpy as np
from sklearn.model_selection import GroupKFold

import methods as M
import perturb as P
from evaluate import load

TRICKS = list(P.PERTURBATIONS) + list(P.NEGATIVE)
NAMES = TRICKS + ["hard"]
FNS = {**P.PERTURBATIONS, **P.NEGATIVE, "hard": P.hard}
from title_detection import THRESHOLD  # below this the learned models answer "no title"


def hit(pred, gold):
    return pred is None if not gold else pred in gold


def build(slides):
    hidden = [dict(s, recs=M.hide(s["recs"])) for s in slides]
    sets = {"clean": hidden}
    for k, name in enumerate(NAMES):
        sets[name] = P.apply_all(hidden, FNS[name], seed=k + 1)
    return sets


def matrices(sets):
    return {n: [M.build_xy([s]) for s in ss] for n, ss in sets.items()}


def fold_hits(model, mats, sets, idx, name):
    hits = 0
    for i in idx:
        p = model.proba_x(mats[name][i][0])
        hits += hit(int(np.argmax(p)) if p.max() >= THRESHOLD else None, sets[name][i]["gold"])
    return hits


def run(slides, folds=5):
    sets = build(slides)
    mats = matrices(sets)
    groups = [s["deck"] for s in slides]
    idx_all = np.arange(len(slides))
    rule_fns = {"first in z-order": M.first_in_order, "topmost": M.topmost, "largest font": M.largest_font,
                "hybrid hand-weighted": M.hybrid_rules}
    table = {}
    for rn, fn in rule_fns.items():
        table[rn] = {n: np.mean([hit(fn(s["recs"]), s["gold"]) for s in sets[n]]) for n in ["clean"] + NAMES}

    for kind in ("logreg", "gbm"):
        counts = {f"{kind} plain": {}, f"{kind} aug": {}, f"{kind} LOPO": {}}
        for tr, te in GroupKFold(folds).split(idx_all, groups=groups):
            def fit(train_names):
                X = np.vstack([mats[n][i][0] for n in train_names for i in tr])
                y = np.concatenate([mats[n][i][1] for n in train_names for i in tr])
                m = M.LearnedTitle(kind).fit_xy(X, y)
                m.proba_x = lambda x, m=m: m.model.predict_proba(x)[:, 1]
                return m
            plain, aug = fit(["clean"]), fit(["clean"] + NAMES)
            for n in ["clean"] + NAMES:
                for label, model in ((f"{kind} plain", plain), (f"{kind} aug", aug)):
                    counts[label][n] = counts[label].get(n, 0) + fold_hits(model, mats, sets, te, n)
            for n in TRICKS:
                lopo = fit(["clean"] + [x for x in TRICKS if x != n])
                counts[f"{kind} LOPO"][n] = counts[f"{kind} LOPO"].get(n, 0) + fold_hits(lopo, mats, sets, te, n)
        for label, d in counts.items():
            table[label] = {n: v / len(slides) for n, v in d.items()}
    return table


def show(table):
    cols = ["clean"] + NAMES
    print(f"{'method':24}" + "".join(f"{c[:11]:>12}" for c in cols) + f"{'mean(perturbed)':>17}")
    for label, row in table.items():
        vals = [row.get(c) for c in cols]
        pert = [v for c, v in zip(cols, vals) if c != "clean" and v is not None]
        print(f"{label:24}" + "".join(f"{v:12.3f}" if v is not None else f"{'-':>12}" for v in vals)
              + f"{np.mean(pert):17.3f}")


if __name__ == "__main__":
    slides = [s for s in load() if s["gold"]]
    if "--sample" in sys.argv:
        slides = random.Random(0).sample(slides, 400)
    show(run(slides))
