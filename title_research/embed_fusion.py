"""Does a sentence-embedding 'does this wording look like a title?' score help the structural model?

Text-only logistic regression on multilingual MiniLM embeddings (trained on titled slides from
decks OTHER than the natural ones), scored alone and late-fused with the structural GBM:
    fused logit = logit(p_gbm) + w * logit(p_text)
w is picked on the natural set itself (one number), so treat any gain as an upper bound.
"""
import random
import sys

import numpy as np
from sklearn.linear_model import LogisticRegression

import methods as M
import natural as N
from evaluate import load
from features import add_context

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
MODEL = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"


def logit(p):
    p = np.clip(p, 1e-4, 1 - 1e-4)
    return np.log(p / (1 - p))


if __name__ == "__main__":
    from sentence_transformers import SentenceTransformer
    enc = SentenceTransformer(MODEL)
    slides = load()
    nat = N.natural_set(slides)
    decks = {s["deck"] for s in nat}
    pool = [s for s in slides if s["gold"] and s["deck"] not in decks]
    random.Random(0).shuffle(pool)
    pool = pool[:1500]
    texts = [r["text"][:200] for s in pool for r in s["recs"]]
    y = [int(i in s["gold"]) for s in pool for i in range(len(s["recs"]))]
    X = enc.encode(texts, batch_size=64, show_progress_bar=False)
    lr = LogisticRegression(max_iter=2000, C=1.0).fit(X, y)
    gbm = N.train_models(slides, decks, n=800, kinds=("gbm",))["gbm"]

    rows = []  # (label, [(orig_idx, p_gbm, p_text)...])
    for s in nat:
        keep = [i for i, r in enumerate(s["recs"]) if not r["furniture"]]
        if not keep:
            rows.append((s["label"], []))
            continue
        recs = add_context([dict(s["recs"][i]) for i in keep])
        pg = gbm.proba(recs)
        pt = lr.predict_proba(enc.encode([r["text"][:200] for r in recs], show_progress_bar=False))[:, 1]
        rows.append((s["label"], list(zip(keep, pg, pt))))

    def score(w, thr):
        hit_t = hit_n = n_t = n_n = 0
        for label, shapes in rows:
            if shapes:
                z = [logit(g) + w * logit(t) for _, g, t in shapes]
                b = int(np.argmax(z))
                pick, conf = shapes[b][0], 1 / (1 + np.exp(-z[b]))
            else:
                pick, conf = None, 0
            if label >= 0:
                n_t += 1
                hit_t += pick == label and conf >= thr
            else:
                n_n += 1
                hit_n += conf < thr
        return hit_t / n_t, hit_n / n_n, (hit_t + hit_n) / (n_t + n_n)

    print("text-only argmax title acc:", np.mean([
        max(shapes, key=lambda x: x[2])[0] == label for label, shapes in rows if label >= 0 and shapes]))
    for w in (0, 0.25, 0.5, 1.0, 2.0):
        best = max(((score(w, t), t) for t in (0.02, 0.05, 0.1, 0.2, 0.3, 0.5)), key=lambda x: x[0][2])
        print(f"w={w:<4} best thr={best[1]:<5} title={best[0][0]:.3f} none={best[0][1]:.3f} overall={best[0][2]:.3f}")
