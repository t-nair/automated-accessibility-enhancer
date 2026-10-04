"""Print unlabeled slides compactly so a human (or I) can hand-label the title: python label_dump.py own|zenodo START COUNT"""
import json
import random
import sys
from pathlib import Path

from evaluate import load

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

src, start, count = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
pool = [s for s in load() if not s["gold"] and s["source"] == src and (len(s["recs"]) >= 2 or src == "own")]
random.Random(1).shuffle(pool)
if src == "own":
    pool.sort(key=lambda s: (s["deck"], s["slide_no"]))
for s in pool[start:start + count]:
    print(f"## {s['deck']}#{s['slide_no']}")
    for i, r in enumerate(s["recs"][:9]):
        print(f"  {i}: y{r['y']:.2f} x{r['x']:.2f} w{r['w']:.2f} {r['pt']:.0f}pt {r['ph'] or '-'} | {r['text'][:55]}")
