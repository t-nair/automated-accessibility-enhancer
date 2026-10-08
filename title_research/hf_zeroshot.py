"""Text-only zero-shot baselines from Hugging Face: does the WORDING alone say 'this is a title'?

Each non-footer shape's text is scored with an NLI zero-shot classifier; the slide's title is
the highest-scoring shape (None if the best score is below `threshold`). Layout is ignored on
purpose: this measures what a plain text classifier can do without geometry.
"""
HYPOTHESIS = "This text is the short heading of a presentation slide."


class ZeroShotTitle:
    def __init__(self, model="MoritzLaurer/deberta-v3-base-zeroshot-v2.0", threshold=0.5):
        from transformers import pipeline
        self.pipe = pipeline("zero-shot-classification", model=model)
        self.threshold = threshold

    def proba(self, recs):
        texts = [r["text"][:200] for r in recs]
        out = self.pipe(texts, candidate_labels=["slide heading", "body text"], hypothesis_template="This text is {}.")
        if isinstance(out, dict):
            out = [out]
        return [o["scores"][o["labels"].index("slide heading")] for o in out]

    def __call__(self, recs):
        if not recs:
            return None
        p = self.proba(recs)
        best = max(range(len(p)), key=p.__getitem__)
        return best if p[best] >= self.threshold else None
