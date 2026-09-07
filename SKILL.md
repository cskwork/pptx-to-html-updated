---
name: pptx-to-html
description: Convert a PowerPoint .pptx into a browser presentation using the bundled converter, preserving supported formatting and media while reporting conversion gaps. Use when the requested output is HTML, not when editing the source deck.
---

# PowerPoint to HTML Converter

Convert the supplied presentation and verify it against the original. Keep the PPTX authoritative; do not promise complete or percentage-based visual fidelity without measured comparison.

## Convert

Use the user's input and output paths. Resolve the script relative to this skill directory, not a hardcoded host installation:

```bash
python3 <skill-dir>/scripts/convert_pptx_to_html_v2.py /abs/input.pptx /abs/output-dir 150
```

The final argument is optional image DPI (default 150). Layout uses fixed 96 DPI independently; increasing export DPI must not change element coordinates. Check the installed interpreter and dependencies in [requirements.txt](requirements.txt) before running; this package is not standard-library-only.

The converter writes `<input-stem>.html`, a report, conversion log, and an `assets/` directory. Inspect the reported paths instead of assuming a fixed filename. Keep the assets beside the HTML when delivering or copying it. Charts use Chart.js from a CDN; this is not a single offline file unless those dependencies are explicitly bundled.

## Supported paths and limits

Inspect the result for features actually present in this deck:

| Feature | Converter path | What to verify |
|---|---|---|
| Text, shapes, backgrounds, images, tables, hyperlinks, media | Main converter | Position, layering, formatting, missing content, links, and browser media playback |
| Charts | `scripts/chart_extractor.py` | Extracted series/labels, chart type, and CDN availability |
| Custom shapes | `scripts/shape_geometry.py` | SVG geometry and positioning |
| SmartArt | `scripts/smartart_parser.py` | Extracted text and simplified hierarchy; do not assume the original visual layout |
| Animations and effects | `scripts/animation_handler.py` | Supported effects and timing; complex PowerPoint behavior may differ |
| Fonts | `scripts/font_manager.py` | Availability, fallback, embedding permissions, and resulting line wraps |

Complex master inheritance, advanced geometry, custom fonts, animations, and unsupported actions can degrade. The converter may continue after an element fails; exit 0 alone is not proof that every element survived. Read `conversion.log` and the generated report for warnings. Do not run embedded macros or publish the deck merely to convert it.

## Verify and deliver

- Compare slide count, order, visible text, media, tables/charts, and any warned-about elements with the source.
- Inspect representative slides and every known conversion gap in a browser, including dense/overlapping content and the requested viewport sizes.
- Exercise navigation and the hyperlinks/media relevant to the deck. Report unavailable codecs, font substitutions, broken assets, or network requirements.
- If only static checks ran, say so; do not claim browser fidelity or successful playback.

Return the actual HTML path and any required assets, plus material limitations. Use the host's normal file links; do not assume `computer://` URLs are supported.

For a converter implementation change, read [docs/architecture.md](docs/architecture.md) and repository instructions first. Keep layout/image DPI separate, XML namespace handling, relationship caching, relative asset links, and element-level error reporting intact. Validate changed behavior with a deck that exercises it.
