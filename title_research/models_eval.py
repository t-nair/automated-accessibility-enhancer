"""Score the model-based methods (Laya variants, HF zero-shot) on the natural set. Slow: ~2-4 min per variant on CPU.

  python models_eval.py hf
  python models_eval.py laya table-choice-base verbal-choice-base verbal-noul-base verbal-choice-td ...
"""
import sys

import natural as N
from evaluate import load

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

if __name__ == "__main__":
    nat = N.natural_set(load())
    results = {}
    if sys.argv[1] == "hf":
        from hf_zeroshot import ZeroShotTitle
        # threshold 0: always answer, so this measures ranking alone (the 0.5 default never fires)
        results["HF zero-shot NLI (text only, argmax)"] = N.score(N.without_furniture(ZeroShotTitle(threshold=0.0)), nat)
    else:
        from laya_method import LayaTitle
        for v in sys.argv[2:]:
            style, mode, ckpt = v.split("-")
            m = LayaTitle(model="typed-decisions" if ckpt == "td" else None, style=style, mode=mode)
            results[f"Laya {v}"] = N.score(N.without_furniture(m), nat)
            N.show({f"Laya {v}": results[f"Laya {v}"]})
            sys.stdout.flush()
    N.show(results)
