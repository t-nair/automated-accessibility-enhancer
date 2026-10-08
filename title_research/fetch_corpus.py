"""Download public .pptx files to use as a labeled title-detection corpus.

Sources are test fixtures from LibreOffice and Apache POI (deliberately odd layouts)
plus python-pptx's own fixtures. Files land in title_research/corpus/<source>/ and are
git-ignored; rerun this script to rebuild them.
"""
import json
import time
import subprocess
import urllib.request
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

CORPUS = Path(__file__).parent / "corpus"
MAX_BYTES = 6_000_000  # skip giant decks, they only slow the experiments down

SOURCES = {
    "libreoffice": ("LibreOffice/core", "sd/qa/unit/data/pptx"),
    "poi": ("apache/poi", "test-data/slideshow"),
    "pythonpptx": ("scanny/python-pptx", "features/steps/test_files"),
}


def list_pptx(repo, path):
    out = subprocess.run(["gh", "api", f"repos/{repo}/contents/{path}"],
                         capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
    try:
        items = json.loads(out)
    except json.JSONDecodeError:
        return []
    return [i for i in items if i["name"].lower().endswith(".pptx") and i["size"] <= MAX_BYTES]


def grab(args):
    dest, url = args
    if dest.exists():
        return
    try:
        dest.write_bytes(urllib.request.urlopen(url, timeout=30).read())
    except Exception as e:
        print("fail", url, e)


REAL_QUERIES = ["lecture", "slides", "introduction", "course", "workshop", "overview", "tutorial", "week",
                "chapter", "seminar", "presentation", "training", "class", "lab", "module", "unit", "talk",
                "project", "report", "proposal", "review", "agenda", "intro", "summary", "research", "teaching"]


def fetch_real(per_query=60):
    """Real-world decks found through GitHub code search (deduped by blob sha)."""
    dest_dir = CORPUS / "real"
    dest_dir.mkdir(parents=True, exist_ok=True)
    seen = set()
    for q in REAL_QUERIES:
        time.sleep(7)  # code search is rate limited
        out = subprocess.run(["gh", "search", "code", "--extension", "pptx", q, "--limit", str(per_query),
                              "--json", "path,repository,sha"], capture_output=True, text=True, encoding="utf-8", errors="replace").stdout
        try:
            hits = json.loads(out)
        except json.JSONDecodeError:
            continue
        for h in hits:
            if h["sha"] in seen:
                continue
            seen.add(h["sha"])
            dest = dest_dir / f"{h['sha'][:10]}.pptx"
            if dest.exists():
                continue
            url = f"https://github.com/{h['repository']['nameWithOwner']}/raw/HEAD/{quote(h['path'])}"
            try:
                data = urllib.request.urlopen(url, timeout=30).read()
            except Exception:
                continue
            if data[:2] == b"PK" and 20_000 < len(data) <= MAX_BYTES:  # a real zip, not an HTML error page
                dest.write_bytes(data)
    print(len(list(dest_dir.glob("*.pptx"))), "real decks")


if __name__ == "__main__":
    fetch_real()
    jobs = []
    for name, (repo, path) in SOURCES.items():
        (CORPUS / name).mkdir(parents=True, exist_ok=True)
        for item in list_pptx(repo, path):
            jobs.append((CORPUS / name / item["name"], item["download_url"]))
    with ThreadPoolExecutor(8) as ex:
        list(ex.map(grab, jobs))
    print(len(jobs), "files requested")
