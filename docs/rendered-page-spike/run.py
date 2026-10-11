# Spike for docs/rendered-page.md. Needs: pip install playwright; playwright install chromium; npm pack axe-core (extract to package/).
import json, time, glob, pathlib
from playwright.sync_api import sync_playwright
axe = pathlib.Path("package/axe.min.js").read_text(encoding="utf-8")
t0=time.time()
with sync_playwright() as p:
    b = p.chromium.launch(); t_launch=time.time()-t0
    for f in sorted(glob.glob("pages/*.html")):
        pg = b.new_page(); t=time.time()
        pg.goto(pathlib.Path(f).resolve().as_uri())
        pg.evaluate(axe)
        r = pg.evaluate("axe.run(document,{runOnly:{type:'tag',values:['wcag2a','wcag2aa','wcag21a','wcag21aa']}})")
        dt=time.time()-t
        print(f"\n== {f}  ({dt:.2f}s)")
        for k in ("violations","incomplete"):
            print(" ",k)
            for v in r[k]:
                print(f"    {v['id']:32} {len(v['nodes'])} node(s)  {[t for t in v['tags'] if t.startswith('wcag') and t[4:].isdigit()]}")
        # keyboard trap probe: tab through
        seen=[]
        for i in range(8):
            pg.keyboard.press("Tab")
            seen.append(pg.evaluate("document.activeElement.tagName+(document.activeElement.id?'#'+document.activeElement.id:'')+':'+(document.activeElement.textContent||'').trim().slice(0,12)"))
        print("  tab order:", seen)
    b.close()
print(f"\nlaunch {t_launch:.2f}s total {time.time()-t0:.2f}s")
