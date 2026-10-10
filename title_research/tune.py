"""How do the clean/perturbed training mix and the abstain threshold trade off on the natural set?"""
import sys

import numpy as np

import natural as N
from evaluate import load
from features import add_context

THRESHOLDS = (0.02, 0.05, 0.1, 0.2, 0.3, 0.5)


def rows(model, nat):
    """(label, predicted index, best probability) per natural slide, footers excluded."""
    out = []
    for s in nat:
        keep = [i for i, r in enumerate(s["recs"]) if not r["furniture"]]
        if not keep:
            out.append((s["label"], -1, 0.0))
            continue
        p = model.proba(add_context([dict(s["recs"][i]) for i in keep]))
        out.append((s["label"], keep[int(np.argmax(p))], float(p.max())))
    return out


def sweep(rs):
    titled = [(l, i, q) for l, i, q in rs if l >= 0]
    none = [(l, i, q) for l, i, q in rs if l < 0]
    print(f"   argmax-only title acc {np.mean([l == i for l, i, q in titled]):.3f}")
    for t in THRESHOLDS:
        a = np.mean([i == l and q >= t for l, i, q in titled])
        b = np.mean([q < t for l, i, q in none])
        print(f"   thr {t:<5} title={a:.3f} none={b:.3f} overall={(a * len(titled) + b * len(none)) / len(rs):.3f}")


if __name__ == "__main__":
    slides = load()
    nat = N.natural_set(slides)
    decks = {s["deck"] for s in nat}
    for w in (1, 4, 12):
        m = N.train_models(slides, decks, n=800, kinds=("gbm",), clean_weight=w)["gbm"]
        print(f"== gbm, clean copies x{w}")
        sweep(rows(m, nat))
        sys.stdout.flush()
