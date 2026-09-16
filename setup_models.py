"""One-time download of the pinned English model; inference never downloads."""
from __future__ import annotations

import hashlib
from pathlib import Path
from urllib.request import urlopen

MODEL_PATH = Path(__file__).parent / "ocr_models/en_PP-OCRv4_rec_mobile.onnx"
MODEL_URL = "https://www.modelscope.cn/models/RapidAI/RapidOCR/resolve/v3.9.2/onnx/PP-OCRv4/rec/en_PP-OCRv4_rec_mobile.onnx"
MODEL_SHA256 = "e8770c967605983d1570cdf5352041dfb68fa0c21664f49f47b155abd3e0e318"


def verified_model() -> bool:
    return MODEL_PATH.is_file() and hashlib.sha256(MODEL_PATH.read_bytes()).hexdigest() == MODEL_SHA256


def main() -> None:
    if verified_model():
        print("English OCR model is installed and verified.")
        return
    with urlopen(MODEL_URL, timeout=120) as response:
        data = response.read(30_000_001)
    if len(data) > 30_000_000 or hashlib.sha256(data).hexdigest() != MODEL_SHA256:
        raise RuntimeError("OCR model download failed its size/checksum verification.")
    MODEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    temporary = MODEL_PATH.with_suffix(".download")
    temporary.write_bytes(data)
    temporary.replace(MODEL_PATH)
    print("English OCR model installed and verified.")


if __name__ == "__main__":
    main()
