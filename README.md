# Rio Economic Development Slides

15–20 minute investor presentation for German commerce board members.

**Topic:** Municipal strategies that drove Rio de Janeiro's economic recovery (2021–2028), outpacing Brazil and the state of RJ.

**Presenter:** Otávio Bopp, SMDEIS

## Build

```bash
pip install python-pptx pillow
python scripts/build_pptx.py
# → output/rio_econ_slides.pptx
```

## Edit content

All slide text lives in `scripts/slide_content.py` — edit there, then rebuild.

## Docs

- `longtermplan.md` — axes, slide map, future iterations
- `logs.md` — round-by-round action log
