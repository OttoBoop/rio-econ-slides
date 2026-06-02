# Logs — Rio Economic Development Slides

---

## Round 1 — 2026-06-02

**Goal:** Initial build — 17-slide PPTX from paper + WhatsApp notes.

**Actions:**
- Created folder: `/home/otavio/Documents/vscode/rio-econ-slides/`
- Copied assets: P06, C2, C11, C12, C13 graphs from PSD-Rio project; Otávio headshot; SMDEIS logos
- Wrote `scripts/slide_content.py` — all 17 slides as Python dicts (easy to edit without touching layout)
- Wrote `scripts/build_pptx.py` — python-pptx builder with 11 slide type renderers
- Built `output/rio_econ_slides.pptx` — 17 slides, 679K

**Slide types implemented:**
- title, content, stats, chart_slide, axes_grid, axis_detail, remaining_axes, break_slide, forward, why_slide, bio_slide

**Known gaps / TODOs for Round 2:**
- [ ] Confirm 12 axes with Otávio (OGG audio not transcribable via text tools)
- [ ] Visual review: open PPTX and check layout on each slide
- [ ] Possibly add C2 (PIB growth) as a second chart slide
- [ ] Verify photo orientation (headshot JPEG may need rotation check)
- [ ] Confirm break slide (#13) is visually distinct enough in LibreOffice / PowerPoint

**Build command:**
```bash
cd /home/otavio/Documents/vscode/rio-econ-slides
python scripts/build_pptx.py
```

**Output:** `output/rio_econ_slides.pptx`
