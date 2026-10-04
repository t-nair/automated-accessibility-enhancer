"""Zero-shot title detection with Laya (convaiinnovations/laya), a local decision model.

Laya reads a state (text or JSON) and answers typed questions in one forward pass, no text
generated. Variants tried here:
  table   one line per shape with numbers (position %, font pt), one `choice` question
  verbal  one sentence per shape in plain words ("the largest text, at the very top ..."),
          one `choice` question
  noul    verbal state, one yes/no (`noul`) question per shape, best yes-probability wins
Each can run on the base checkpoint or on `typed-decisions`.
"""
import warnings

warnings.filterwarnings("ignore")

MAX_SHAPES = 10
MAX_TEXT = 50
QUESTION = ("Which shape is the title of this slide? The title is the short heading that names the slide: "
            "usually the largest text, near the top, not a full sentence and not a bullet list.")


def _trim(text):
    return text[:MAX_TEXT] + ("..." if len(text) > MAX_TEXT else "")


def _keep(recs):
    """Up to MAX_SHAPES shapes (largest font, then highest), back in reading order."""
    keep = sorted(range(len(recs)), key=lambda i: (-recs[i]["pt"], recs[i]["y"]))[:MAX_SHAPES]
    return sorted(keep)


def describe(r, recs):
    where = "at the very top" if r["y"] < 0.12 else "in the upper part" if r["y"] < 0.35 else \
            "in the middle" if r["y"] < 0.7 else "at the bottom"
    size = ("the largest text" if r["pt_rank"] == 0 else "larger than most text" if r["pt_rel"] > 0.7
            else "medium-sized text" if r["pt_rel"] > 0.4 else "small text")
    length = "one or two words" if r["n_words"] <= 2 else "a short phrase" if r["n_words"] <= 8 else \
             "a sentence" if r["n_words"] <= 25 else "a long paragraph"
    style = (", bold" if r["bold"] else "") + (", bulleted" if r["bullets"] else "") + \
            (", centered" if r["align"].lower().startswith("center") else "")
    return f'"{_trim(r["text"])}" is {length}, {size}, {where}{style}.'


def serialize(recs, style="table"):
    keep = _keep(recs)
    if style == "table":
        lines = []
        for i in keep:
            r = recs[i]
            flags = ("bold " if r["bold"] else "") + ("bullets " if r["bullets"] else "")
            lines.append(f'{i + 1} | "{_trim(r["text"])}" | top {r["y"] * 100:.0f}% | left {r["x"] * 100:.0f}% | '
                         f'width {r["w"] * 100:.0f}% | {r["pt"]:.0f}pt {flags}| {r["n_words"]} words')
        return "Slide text shapes (id | text | position | size):\n" + "\n".join(lines), keep
    return "A slide has these text shapes.\n" + "\n".join(f"Shape {i + 1}: {describe(recs[i], recs)}" for i in keep), keep


class LayaTitle:
    def __init__(self, model=None, style="table", mode="choice", threshold=0.5):
        from laya import Router
        self.router, self.model, self.style, self.mode, self.threshold = Router(), model, style, mode, threshold

    def proba(self, recs):
        """{shape_index or None: probability}"""
        state, keep = serialize(recs, self.style)
        kw = {"model": self.model} if self.model else {}
        if self.mode == "choice":
            crit = {str(i + 1): "this shape is the title" for i in keep}
            crit["none"] = "no shape is a title"
            q = {"title": {"type": "choice", "instructions": QUESTION, "criteria": crit}}
            out = self.router.predict(state, q, **kw)["answers"]["title"]["probabilities"]
            return {(None if k == "none" else int(k) - 1): v for k, v in out.items()}
        q = {f"s{i + 1}": {"type": "noul", "instructions": f"Is Shape {i + 1} the title of this slide?"} for i in keep}
        ans = self.router.predict(state, q, **kw)["answers"]
        return {i: ans[f"s{i + 1}"]["noul"] if isinstance(ans[f"s{i + 1}"]["noul"], float)
                else ans[f"s{i + 1}"]["probabilities"].get("yes", 0.0) for i in keep}

    def __call__(self, recs):
        if not recs:
            return None
        probs = self.proba(recs)
        best = max(probs, key=probs.get)
        if self.mode == "noul" and probs[best] < self.threshold:
            return None
        return best
