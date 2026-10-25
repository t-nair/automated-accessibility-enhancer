"""Sample real-world decks (CC-licensed Zenodo uploads) from the Forceless/Zenodo10K dataset on Hugging Face."""
import random
import shutil
from pathlib import Path

from huggingface_hub import HfApi, hf_hub_download

REPO, WANT, MAX_BYTES = "Forceless/Zenodo10K", 160, 6_000_000
DEST = Path(__file__).parent / "corpus" / "zenodo"

if __name__ == "__main__":
    DEST.mkdir(parents=True, exist_ok=True)
    api = HfApi()
    paths = [f for f in api.list_repo_files(REPO, repo_type="dataset") if f.startswith("pptx/") and f.endswith(".pptx")]
    random.Random(0).shuffle(paths)
    got = 0
    for i in range(0, len(paths), 50):
        batch = paths[i:i + 50]
        for info in api.get_paths_info(REPO, batch, repo_type="dataset"):
            if info.size > MAX_BYTES or got >= WANT:
                continue
            try:
                src = hf_hub_download(REPO, info.path, repo_type="dataset")
            except Exception as e:
                print("fail", info.path, str(e)[:60])
                continue
            shutil.copy(src, DEST / f"{got:03d}.pptx")
            got += 1
        if got >= WANT:
            break
    print(got, "decks")
