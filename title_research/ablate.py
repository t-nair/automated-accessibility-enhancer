"""Compare feature configs: robustness (perturbed CV, incl. unseen-trick LOPO) and the natural set."""
import random
import sys

import numpy as np

import methods as M
import natural as N
import robustness as R
from evaluate import load

CONFIGS = [
    {"z": True, "text": False},
    {"z": False, "text": False},
    {"z": False, "text": True},
    {"z": True, "text": True},
]

if __name__ == "__main__":
    slides = load()
    nat = N.natural_set(slides)
    decks = {s["deck"] for s in nat}
    gold = [s for s in slides if s["gold"] and s["deck"] not in decks]
    sample = random.Random(0).sample(gold, 300)
    for cfg in CONFIGS:
        M.CONFIG.update(cfg)
        table = R.run(sample, folds=3)
        models = N.train_models(slides, decks, n=800)
        print(f"\n=== {cfg}")
        for label in ("logreg aug", "logreg LOPO", "gbm aug", "gbm LOPO"):
            row = table[label]
            pert = [v for k, v in row.items() if k != "clean"]
            print(f"  robust {label:12} clean={row.get('clean', float('nan')):.3f} mean(perturbed)={np.mean(pert):.3f}")
        for kind in ("logreg", "gbm"):
            r = N.score(N.without_furniture(models[kind]), nat)
            print(f"  natural {kind:7} title={r['title_acc']:.3f} none={r['none_acc']:.3f} overall={r['overall']:.3f}")
        sys.stdout.flush()
