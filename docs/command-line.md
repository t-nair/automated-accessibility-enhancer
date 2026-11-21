# Command line

The web app is the main way to use the tool. These run the same pipelines on files on your own machine, after the setup in the README's quick start.

## A folder of decks

```bash
python z_reorder.py                       # reads pptx_input/, writes pptx_output/
python z_reorder.py my_decks fixed_decks  # or name the folders
```

Each deck gets an `_updated.pptx` and an `_alt_text` report. Originals are left where they are; add `--move-originals` to move each one into the output folder once it is done (an existing file there is never replaced).

## A PDF

```bash
python -m pipeline.cli "input.pdf" --output-dir output/remediation
```

This writes a tagged PDF, an HTML companion and a JSON manifest. PDFs with images stop until each description is approved in a review file ([pipeline/README.md](../pipeline/README.md)). On the website the same approval happens on the submission's review page. On Windows, `./pipeline/setup.ps1 -Python python` sets up the environment in one step.
