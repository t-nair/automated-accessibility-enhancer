# Finding a slide's title: what we tried and what we shipped

**Result.** `z_reorder.py` used to find a title only through the title placeholder type or a shape name containing "title". It now falls back to `title_detection.py`: a 110 KB gradient-boosted model that reads each text shape's size, position, formatting and wording. On 129 hand-labeled slides with no title placeholder it finds the title on **91%** of the slides that have one (the old rule: 2%), and correctly says "no title" on 70% of the slides that have none.

## 1. The problem with the old rule

`z_reorder.is_title` returns true for a title/center-title placeholder, or any shape whose name contains "title" (not "subtitle"). Decks whose author built the title some other way are invisible to it. The real deck `test.pptx` shows this: 18 of 22 slides keep their title in a *body* placeholder named "Text Placeholder 1", so the old rule found 4 titles and missed the rest.

## 2. Data

| Source | Slides | Use |
|---|---:|---|
| LibreOffice, Apache POI and python-pptx test fixtures (`fetch_corpus.py`) | ~440 titled | odd layouts, mostly simple |
| 160 random decks from the CC-licensed Zenodo10K set on Hugging Face (`fetch_zenodo.py`) | ~1,790 titled | real-world decks, several languages |
| the repo owner's own decks | 36 titled | real use |

A slide is **labeled** when it has a title placeholder with text; that shape is the gold title (2,266 slides). Two evaluation setups follow from this:

* **Cues hidden** (labeled slides): placeholder type and shape name are removed from every method, as if the title had been typed into a text box. Without this the old rule would score 100% by construction.
* **Natural set** (`labels_natural.json`): 129 real slides with *no* title placeholder, hand-labeled with the title shape, or "no title" (30 slides, e.g. diagrams and quotes). 7 ambiguous slides were left out. **Labels were written by Claude from a compact text dump of each slide, not by the deck authors**, so they carry some judgment noise. With 129 slides, one slide is 0.8 points.
* **Perturbations** (`perturb.py`): each labeled slide is rewritten with an awkward-layout trick and the gold label kept: kicker text above the title, a giant decorative number, a side-column title, a title in the middle of the slide, all fonts equal, a sentence-length title, big bold body text, jittered geometry, shrink-wrapped boxes, a title-only slide, scrambled z-order, and "title removed" (right answer: none). `hard` stacks 2-3 of them.

## 3. Methods and results

Natural set, held out (models never saw these decks). "title" = share of slides with a title answered correctly; "none" = share of title-less slides correctly left alone.

| Method | title | none | overall | s/slide (CPU) |
|---|---:|---:|---:|---:|
| Current rule (placeholder type / name) | 0.02 | 1.00 | 0.25 | ~0 |
| First text shape in z-order | 0.77 | 0.20 | 0.64 | ~0 |
| Topmost text shape | 0.92 | 0.20 | 0.75 | ~0 |
| Largest font | 0.92 | 0.20 | 0.75 | ~0 |
| Hand-weighted score (size, height, length, punctuation, ...) | 0.91 | 0.27 | 0.76 | ~0 |
| Logistic regression on shape features | 0.83 | 0.67 | 0.79 | ~0 |
| **Gradient-boosted trees on shape features (shipped)** | **0.91** | **0.70** | **0.86** | 0.01-0.08 |
| Laya (base), shape table + one `choice` question | 0.44 | 0.40 | 0.43 | 1.5 |
| Laya (base), plain-words description + `choice` | 0.44 | 0.30 | 0.41 | 1.5 |
| Laya (base), plain-words + one yes/no question per shape | 0.15 | 0.53 | 0.24 | 4.4 |
| Laya typed-decisions checkpoint, table / words | 0.26 / 0.22 | 0.33 / 0.30 | 0.28 / 0.24 | 0.9-1.7 |
| HF zero-shot NLI (deberta-v3-base-zeroshot), text only, always answers | 0.29 | 0.20 | 0.27 | 9.5 |
| MiniLM multilingual embedding + logistic regression, text only | 0.73 (argmax) | - | - | ~0.01 |
| GBM + that text score fused (weight picked on this set) | - | - | 0.868 -> 0.876-0.884 | needs a ~470 MB model |

Layout stress test (400 labeled slides, cues hidden, deck-grouped 5-fold CV, `robustness.py`), share correct:

| Method | clean | mean over 13 tricks | worst single trick |
|---|---:|---:|---|
| First in z-order | 0.91 | 0.78 | scrambled z-order 0.48 |
| Topmost | 0.97 | 0.76 | kicker text above title 0.04 |
| Largest font | 0.96 | 0.81 | big bold body text 0.12 |
| Hand-weighted score | 0.97 | 0.87 | giant number 0.69 |
| GBM trained on clean slides only | 0.99 | 0.81 | giant number 0.33 |
| **GBM trained on clean + perturbed copies (shipped)** | **0.99** | **0.98** | none below 0.86 |
| Same, but the tested trick was held out of training | - | 0.86 | giant number 0.24, no-title 0.43 |

### What we learned

* Every single-cue rule has a layout that breaks it. "Topmost" dies when a small line sits above the title; "largest font" when body text is big and bold; "first in order" when z-order is scrambled, and none of them can say "this slide has no title".
* A model on shape features ranks well (96% argmax on the natural titled slides) but needs two things to be trusted: training on perturbed copies (a model trained on clean slides alone collapses on tricks it never saw), and an abstain threshold for slides without a title.
* Training data made from title placeholders teaches the model habits that text boxes do not have (full-width boxes at exact layout coordinates, title first in z-order). The jitter and shrink-wrap perturbations exist to unlearn that. Z-order position stayed in as a feature: on the natural set it lifted argmax accuracy from 92% to 96% (overall 0.845 to 0.860, same training and threshold).
* **Unseen tricks still hurt.** Holding a trick out of training drops the model from 0.98 to 0.86 (a giant decorative number above the title is the worst). The augmentation list is the model's real coverage.
* **Laya does not work for this task zero-shot.** It is a decision model for routing/scoring text and JSON state, not a layout model: 0.41-0.43 overall with the best prompt, behind even "topmost", and 0.9-4.4 s per slide on CPU. Four prompt/checkpoint variants all landed in the same range. The pip package only ships inference and calibration code (no training script), so fine-tuning it on slide shapes was not tried.
* Wording alone is a real signal (73% argmax from sentence embeddings) but adds under 2 points on top of geometry and costs a ~470 MB model, so it was left out. The shipped model uses seven cheap wording features instead (share of capitalized words, trailing "?", numbered heading, ...).

### Not tried

* **Document-layout detectors on a rendered slide image** (DocLayout-YOLO, PP-DocLayout, Docling's layout model). They need each slide rendered first, and this machine has no LibreOffice and no GPU. They would also add a slow per-slide step to recover information (font size, position) that is already in the XML. Worth a test if a rendering step is added for another reason.
* Anything using a hosted LLM.

## 4. How the shipped model works

`title_detection.guess_title(slide)` runs only when a slide has no real title placeholder and no shape named "title". It skips footer/date/slide-number placeholders, builds one feature row per text shape, and returns the shape with the highest title probability, or `None` if that probability is under `THRESHOLD` (0.1; overall accuracy was flat from 0.05 to 0.2 on the natural set, so this was not tuned tightly). It never raises: a slide it cannot read, or a missing model file, means "no guess" and the pipeline behaves as before.

`title_model.joblib` is a pickled scikit-learn model, so `requirements.txt` pins `scikit-learn~=1.8.0`. Retrain with `python train_final.py` after changing features or `perturb.py`.

## 5. Reproducing

```
cd title_research
python fetch_corpus.py && python fetch_zenodo.py   # downloads decks into corpus/ (git-ignored)
python build_dataset.py                            # parses them into data/slides.pkl
python evaluate.py                                 # rules + learned, cues hidden
python robustness.py --sample                      # layout stress test
python natural.py                                  # the hand-labeled natural set, held out
python models_eval.py laya table-choice-base verbal-choice-base   # slow
python models_eval.py hf                                          # slow
python embed_fusion.py                             # sentence-embedding experiment
python train_final.py                              # writes ../title_model.joblib
```

## 6. Limits

* The natural set is small (129 slides) and hand-labeled by one reviewer; treat differences under ~3 points as noise. The abstain threshold and fusion weight were picked on it.
* `train_final.py` trains on the natural slides too, so the shipped model's accuracy on them is not meaningful. The held-out numbers above come from `natural.py`.
* Zenodo decks skew toward research talks, and the labeled slides skew toward conventional layouts. A set of labeled slides from the decks this tool will actually see is the most useful next addition.
* Only text shapes are candidates. A title that is an image or WordArt-style graphic is not found.
