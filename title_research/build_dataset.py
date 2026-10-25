"""Parse every deck in corpus/ (+ the user's own decks) into data/slides.pkl.

One entry per slide that has at least one text shape. `gold` lists the indexes of the real
title placeholder(s) with text; slides without one have gold == [] (unlabeled: either no
title or a title typed into a text box).
"""
import pickle
from pathlib import Path

from features import deck_slides

HERE = Path(__file__).parent
OWN_DECKS = Path(r"C:\Users\nairt\Desktop\GitHub Repos\UW_Accessibility_Enhancer")


def decks():
    for p in sorted((HERE / "corpus").rglob("*.pptx")):
        yield p.parent.name, p
    for p in sorted(OWN_DECKS.glob("*.pptx")):
        if not p.name.startswith("~$"):
            yield "own", p


if __name__ == "__main__":
    slides, bad = [], 0
    for source, path in decks():
        try:
            for no, recs in deck_slides(path):
                if not recs:
                    continue
                gold = [i for i, r in enumerate(recs) if r["gold"]]
                slides.append({"deck": path.name, "source": source, "slide_no": no, "gold": gold,
                               "recs": [{k: v for k, v in r.items() if k != "shape"} for r in recs]})
        except Exception as e:  # corpus includes deliberately broken files
            bad += 1
            print("skip", path.name, type(e).__name__, str(e)[:80])
    (HERE / "data").mkdir(exist_ok=True)
    pickle.dump(slides, open(HERE / "data" / "slides.pkl", "wb"))
    lab = sum(bool(s["gold"]) for s in slides)
    print(f"{len(slides)} slides, {lab} labeled, {bad} unreadable decks")
