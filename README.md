# PPTX to HTML Converter

Turn a PowerPoint deck into HTML you can open in any browser. Point the converter at a `.pptx` file and you get back a small folder of HTML, CSS and JavaScript that looks and behaves like the original: same text styling, charts, media, animations, even the fonts embedded in the file. It reads the `.pptx` directly (it's just a ZIP of XML), so PowerPoint itself is never needed.

Phase 2. Python 3.7+.

🌐 **[Live site](https://cskwork.github.io/pptx-to-html-updated/)** · [Architecture](docs/architecture.md) · [Changelog](docs/changelog.md)

## Quick start

```bash
pip install -r requirements.txt
python scripts/convert_pptx_to_html_v2.py deck.pptx output/
```

Open `output/deck.html` in a browser and you're done. Image quality defaults to 150 DPI; add a third argument to change it:

```bash
python scripts/convert_pptx_to_html_v2.py deck.pptx output/ 300
```

If you'd rather call it from code:

```python
from scripts.convert_pptx_to_html_v2 import EnhancedPPTXToHTMLV2

converter = EnhancedPPTXToHTMLV2("deck.pptx", output_dir="output", dpi=150)
html_path = converter.convert()   # path to the HTML file, or None if something broke
```

There are also `convert.sh` and `convert.bat` wrappers that handle filename quoting for you.

## What survives the trip

Most of a slide comes through intact. Text keeps its font, size, color, bold/italic/underline, alignment, bullets and numbered lists. Shapes keep their position, size, rotation, fills (solid and gradient) and borders. Images export at your chosen DPI, video and audio play inline with HTML5 players, tables keep their borders and cell styling, and hyperlinks stay clickable whether they're on text or on a whole shape. Backgrounds resolve the way PowerPoint resolves them: slide first, then layout, then master, with theme colors carried over.

A few things deserve their own mention:

**Charts become live Chart.js canvases**, not screenshots. Bar, line, pie, doughnut, area, scatter, radar and bubble all work, including the 3D variants. Chart.js loads from a CDN, so viewing charts needs an internet connection.

**Custom shapes convert to SVG.** About two dozen common presets (arrows, flowchart blocks, polygons, a star) have hand-drawn paths, and freeform geometry is traced into SVG paths. Anything more exotic falls back to a plain rectangle.

**Embedded fonts are extracted and re-served as WOFF webfonts**, so your typefaces survive on machines that don't have them. Decks without embedded fonts fall back to web-safe stacks.

**Entrance animations** (fade, fly, zoom, bounce and friends) become CSS keyframes that replay each time the slide appears.

**SmartArt is text-only.** You get the words in a labeled outline, not the diagram layout.

## What doesn't work yet

- Macros/VBA, slide notes, and transition sounds are ignored
- Connectors and uncommon preset shapes render as rectangles
- `arcTo` curves inside custom paths are skipped (logged as warnings)
- `wipe` and `split` entrance effects map to keyframes that aren't generated yet, so affected elements may stay hidden; re-save the deck without those effects as a workaround
- Reflections render in Chromium and Safari, not Firefox

## What you get

```
output/
├── deck.html            # the presentation
├── deck.css             # styles: slides, animations, viewer chrome
├── deck.js              # navigation, animation replay, chart rendering
├── deck_report.md       # what was converted, counts and warnings
├── conversion.log       # full log
├── assets/              # images, videos, audio pulled from the deck
└── fonts/               # converted embedded fonts, when present
```

The bundle ships with a small viewer: arrow keys and on-screen buttons to navigate, a progress bar, and scaling that fits any screen size.

## Good to know

- Expect roughly 1–2 seconds per slide and around 100 MB of memory for typical decks.
- The only library actually required is `fonttools` (embedded font conversion). `requirements.txt` also lists `python-pptx` and `openpyxl`, but the converter parses the OOXML itself and never imports them.
- One broken shape won't kill a conversion. Failures are logged per element and summarized in the report file.

## Troubleshooting

**Filenames with spaces or non-Latin characters?** Quote them:
```bash
./convert.sh "deck (final).pptx"
./convert.sh "(한글) 파일명.pptx"
```

**Charts blank?** They need the Chart.js CDN at view time. Check the network and the browser console.

**Fonts look wrong?** The deck probably doesn't embed its fonts. Embed them in PowerPoint before converting.

**Big deck running out of memory?** Lower the DPI (`96` works fine for screens) or convert in batches.

## Project layout

```
pptx-to-html-updated/
├── scripts/
│   ├── convert_pptx_to_html_v2.py   # main converter, generates the HTML/CSS/JS bundle
│   ├── chart_extractor.py           # chart XML → Chart.js configs
│   ├── shape_geometry.py            # preset + custom geometry → SVG paths
│   ├── smartart_parser.py           # SmartArt text extraction
│   ├── animation_handler.py         # animations, shadows, reflections → CSS
│   ├── font_manager.py              # embedded fonts (ODTTF) → WOFF + @font-face
│   └── logger.py                    # logging, summary, markdown report
├── docs/                            # architecture notes + this project's website
├── convert.sh / convert.bat         # platform wrappers
├── QUICKSTART.md
├── SKILL.md                         # agent-skill manifest (Claude Code & friends)
└── README.md
```

## Using it as an agent skill

The repo ships a `SKILL.md`, so coding agents that discover skills (Claude Code et al.) can pick it up directly: point the agent at this repository and ask it to convert a deck. No extra setup.

## More docs

- [Landing page](https://cskwork.github.io/pptx-to-html-updated/) for the feature tour
- [SKILL.md](SKILL.md) for agent-facing usage
- [docs/architecture.md](docs/architecture.md) for how the pipeline works
- [docs/changelog.md](docs/changelog.md) for version history
