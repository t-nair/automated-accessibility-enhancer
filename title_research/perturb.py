"""Turn an ordinary labeled slide into a structurally awkward one, keeping the gold label.

Each perturbation takes (recs, gold_indexes, rng) and returns new (recs, gold_indexes).
They imitate real layout habits: kicker text above the title, a giant decorative number,
a side-column title, a title in the middle of a hero slide, all-equal font sizes, sentence
titles, big bold body text and arbitrary z-order.
"""
import random

from features import add_context


def _blank(text, **kw):
    r = {"z": 0, "text": text, "n_chars": len(text), "n_words": len(text.split()), "n_paras": 1,
         "x": 0.05, "y": 0.02, "w": 0.4, "h": 0.05, "pt": 12.0, "bold": False, "align": "left",
         "bullets": False, "caps": text.isupper(), "end_punct": False, "digits_only": text.isdigit(),
         "rot": 0, "is_auto": False, "gold": False, "ph": None, "name": "", "furniture": False,
         "slide_no": 3, "n_slides": 10, "first": False, "last": False}
    r.update(kw)
    return r


def _copy(recs):
    return [dict(r) for r in recs]


def kicker(recs, gold, rng):
    recs = _copy(recs)
    t = recs[gold[0]]
    k = _blank(rng.choice(["CHAPTER 3", "MODULE 2", "Week 4", "SECTION B"]), y=max(t["y"] - 0.08, 0.0),
               x=t["x"], w=0.25, h=0.05, pt=12.0, caps=True, **{k: t[k] for k in ("slide_no", "n_slides", "first", "last")})
    recs.insert(rng.randint(0, len(recs)), k)
    return add_context(recs), [i for i, r in enumerate(recs) if r is t]


def big_number(recs, gold, rng):
    recs = _copy(recs)
    t = recs[gold[0]]
    n = _blank(str(rng.randint(1, 9)), x=0.03, y=0.0, w=0.14, h=0.35, pt=120.0, bold=True,
               **{k: t[k] for k in ("slide_no", "n_slides", "first", "last")})
    recs.insert(rng.randint(0, len(recs)), n)
    return add_context(recs), [i for i, r in enumerate(recs) if r is t]


def side_title(recs, gold, rng):
    recs = _copy(recs)
    for i, r in enumerate(recs):
        if i in gold:
            r.update(x=0.05, y=0.30, w=0.30, h=0.30, cx=0.2, align="left", n_paras=1)
        elif r["y"] > 0.1 and r["w"] > 0.3:
            r.update(x=0.42, w=0.5)
    return add_context(recs), gold


def low_title(recs, gold, rng):
    recs = _copy(recs)
    for i, r in enumerate(recs):
        if i in gold:
            r.update(y=0.38, align="center")
        else:
            r["y"] = max(r["y"], 0.58)
    return add_context(recs), gold


def flat_fonts(recs, gold, rng):
    recs = _copy(recs)
    for r in recs:
        r["pt"] = 20.0
    return add_context(recs), gold


def z_shuffle(recs, gold, rng):
    order = list(range(len(recs)))
    rng.shuffle(order)
    return add_context([dict(recs[i]) for i in order]), [order.index(g) for g in gold]


def sentence_title(recs, gold, rng):
    recs = _copy(recs)
    t = recs[gold[0]]
    text = "Retention improved by twenty percent once the new onboarding flow replaced the old one"
    t.update(text=text, n_chars=len(text), n_words=len(text.split()), caps=False, end_punct=False)
    return add_context(recs), gold


def bold_body(recs, gold, rng):
    recs = _copy(recs)
    top = max(r["pt"] for i, r in enumerate(recs) if i in gold)
    for i, r in enumerate(recs):
        if i not in gold:
            r.update(bold=True, pt=top * 1.15)
    return add_context(recs), gold


def drop_title(recs, gold, rng):
    """Remove the title: the right answer is now 'no title'. A lone shape is swapped for a footer-like line."""
    recs = [dict(r) for i, r in enumerate(recs) if i not in gold]
    if not recs:
        recs = [_blank("Page 3", y=0.93, x=0.4, w=0.2, pt=10.0, digits_only=False)]
    return add_context(recs), []


def jitter(recs, gold, rng):
    """Placeholders sit at exact layout coordinates; text boxes never do. Blur geometry, rescale all fonts."""
    recs = _copy(recs)
    scale = rng.uniform(0.6, 1.5)
    for r in recs:
        r["x"] = min(max(r["x"] + rng.gauss(0, 0.03), -0.02), 0.9)
        r["y"] = min(max(r["y"] + rng.gauss(0, 0.03), -0.02), 0.95)
        r["w"] = min(max(r["w"] * rng.uniform(0.7, 1.3), 0.03), 1.0)
        r["h"] = min(max(r["h"] * rng.uniform(0.7, 1.3), 0.03), 1.0)
        r["pt"] = r["pt"] * scale
        if rng.random() < 0.1:
            r["bold"] = not r["bold"]
        if rng.random() < 0.2:
            r["align"] = rng.choice(["center", "left", "none"])
    return add_context(recs), gold


def shrink_wrap(recs, gold, rng):
    """Text boxes hug their text; placeholders span the layout. Make short shapes narrow like boxes do."""
    recs = _copy(recs)
    for r in recs:
        chars = r["n_chars"] / max(r["n_paras"], 1)
        w = min(r["w"], max(0.08, chars * r["pt"] * 0.0007 * rng.uniform(1.0, 1.4)))
        if r["align"] == "center":
            r["x"] += (r["w"] - w) / 2
        r["w"] = w
        r["h"] = min(r["h"], max(0.05, r["n_paras"] * r["pt"] * 0.0026 * rng.uniform(1.0, 1.5)))
    return add_context(recs), gold


def title_only(recs, gold, rng):
    """Section headers, 'Questions?', 'Thank you': a slide that is only a title."""
    return add_context([dict(recs[i]) for i in gold]), list(range(len(gold)))


PERTURBATIONS = {"kicker": kicker, "big_number": big_number, "side_title": side_title,
                 "low_title": low_title, "flat_fonts": flat_fonts, "z_shuffle": z_shuffle,
                 "sentence_title": sentence_title, "bold_body": bold_body, "jitter": jitter,
                 "shrink_wrap": shrink_wrap, "title_only": title_only}


NEGATIVE = {"drop_title": drop_title}


def hard(recs, gold, rng, k=None):
    """Stack 2-3 random perturbations."""
    for name in rng.sample(list(PERTURBATIONS), k or rng.choice([2, 3])):
        recs, gold = PERTURBATIONS[name](recs, gold, rng)
    return recs, gold


def apply_all(slides, fn, seed=0):
    """Perturbed copy of each labeled slide (cues already hidden by the caller)."""
    rng = random.Random(seed)
    out = []
    for s in slides:
        recs, gold = fn(s["recs"], s["gold"][:1], rng)
        out.append(dict(s, recs=recs, gold=gold))
    return out
