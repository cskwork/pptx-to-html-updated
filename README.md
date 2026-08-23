# PPTX to HTML Converter — Phase 2 (Production Ready)

Convert PowerPoint presentations (.pptx) into faithful, self-contained HTML bundles — parsed straight from the OOXML source, no PowerPoint required.

🌐 **[Live site & docs](https://cskwork.github.io/pptx-to-html-updated/)** · [Quick Start](#quick-start) · [Architecture](docs/architecture.md) · [Changelog](docs/changelog.md)

## How It Works

The converter reads the .pptx ZIP directly (`zipfile` + `ElementTree`) and translates each slide's DrawingML into an HTML/CSS/JS bundle:

- **Pixel-accurate layout** — EMU coordinates converted to percentage-based positioning, so slides scale responsively while keeping exact placement, rotation, layering (z-order), and aspect ratio.
- **Theme & inheritance aware** — theme color palettes (with lumMod/alpha modifiers), slide backgrounds resolved through the slide → layout → master chain, and placeholder geometry/inheritance pulled from layouts and masters.
- **Group shapes** — nested group transforms (offset, scale, rotation) are flattened correctly onto child elements.

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Convert (150 DPI by default)
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/

# Custom DPI
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/ 300
```

Shell wrappers `convert.sh` (macOS/Linux) and `convert.bat` (Windows) do the same with filename quoting handled.

> **Requirements:** Python 3.7+. The only hard dependency is `fonttools` (embedded font conversion). `requirements.txt` also declares `python-pptx`/`openpyxl`, but the v2 pipeline parses OOXML directly and does not import them.

### Python API

```python
from scripts.convert_pptx_to_html_v2 import EnhancedPPTXToHTMLV2

converter = EnhancedPPTXToHTMLV2(
    'presentation.pptx',
    output_dir='./output',
    dpi=150,
    log_file='./output/conversion.log',
)

html_path = converter.convert()   # returns Path to the HTML file, or None on failure
```

## What Gets Converted

| Category | Support |
|---|---|
| Text formatting | ✅ Font family/size/color, bold/italic/underline, alignment, bullets & numbered lists (incl. roman/alpha formats), indentation, line spacing, mixed formatting per run |
| Shapes | ✅ Solid & multi-stop gradient fills, borders with dash patterns, rotation, z-order |
| Custom geometry | ✅ `custGeom` paths → SVG (`moveTo`/`lnTo`/bezier/`close`), plus 24 preset shapes (rect, ellipse, polygons, star, 9 arrow variants, flowchart process/decision/data/terminator/document) |
| Charts | ✅ 12 PowerPoint chart types → live Chart.js renderings (see below) |
| Tables | ✅ Structure, cell fills, cell borders, text formatting |
| Images | ✅ Extracted at 150 DPI (configurable), exact placement |
| Video / Audio | ✅ Extracted to `assets/`, embedded HTML5 players |
| Hyperlinks | ✅ Text-level and shape-level, open in new tab |
| Backgrounds | ✅ Solid & gradient, resolved through slide → layout → master |
| Embedded fonts | ✅ ODTTF de-obfuscation → WOFF conversion → `@font-face` (Regular/Bold/Italic/BoldItalic) |
| Animations | ✅ Entrance effects (fade, fly, zoom, grow, bounce, swivel, wheel…) → CSS keyframes + JS controller, replayed on slide change |
| Shadows | ✅ Outer/inner shadows → CSS `box-shadow` |
| Reflections | ⚠️ Mapped to `-webkit-box-reflect` (Chromium/Safari only) |
| SmartArt | ⚠️ Text extraction only — rendered as a labeled text outline, layout not reconstructed |

### Charts

All of these are extracted from the chart XML parts (cached series values, names, and colors) and rendered as interactive [Chart.js](https://www.chartjs.org/) canvases:

bar (2D/3D) · line (2D/3D) · pie (2D/3D) · doughnut · area (2D/3D, as filled line) · scatter · radar · bubble

Chart.js loads from a CDN in the generated HTML, so charts need network access at view time.

### Embedded Fonts

If the presentation embeds fonts (PowerPoint's "Embed fonts in the file" option), the converter:

1. Reads `ppt/fontTable.xml` and its relationships,
2. De-obfuscates ODTTF font data using the embed GUID key,
3. Converts each variant to WOFF with fontTools,
4. Emits `@font-face` rules (with `font-display: swap`) pointing at `fonts/`.

Presentations without embedded fonts fall back to web-safe font stacks.

## Output Structure

```
output-directory/
├── presentation.html        # Main presentation file
├── presentation.css         # Styles (slides, animations, viewer chrome)
├── presentation.js          # Viewer logic (navigation, animations, charts)
├── presentation_report.md   # Detailed conversion report (counts, durations, warnings)
├── conversion.log           # Full conversion log
├── assets/                  # Extracted images, videos, audio
└── fonts/                   # Converted embedded fonts (when present)
```

### Viewer Features

The generated bundle ships with a slide viewer:

- Keyboard navigation (arrow keys, space) and on-screen controls
- Progress bar and slide numbering
- Per-slide animation replay
- Responsive scaling for desktop and mobile

## CLI Reference

```
python scripts/convert_pptx_to_html_v2.py <input.pptx> [output_directory] [dpi]
```

| Argument | Required | Default |
|---|---|---|
| `input.pptx` | yes | — |
| `output_directory` | no | alongside the input file |
| `dpi` | no | `150` |

## Known Limitations

- **Macros/VBA, slide notes, transition sounds** — never converted.
- **SmartArt** — text content only; the visual diagram layout is not reconstructed.
- **`arcTo` segments** in custom geometry are skipped (logged as approximations).
- **Connectors and uncommon preset shapes** fall back to a plain rectangle.
- **`wipe`/`split` animation effects** currently map to keyframes that aren't generated — affected elements may stay hidden; re-save without those effects as a workaround.
- **Reflections** render only in WebKit-based browsers.
- **Charts** require Chart.js CDN access at view time.

## Performance

Typical behavior on a modern machine:

- ~1–2 seconds per slide
- ~100 MB memory for standard decks
- HTML 50–300 KB plus assets proportional to media content
- Handles decks of 100+ slides

## Troubleshooting

**Filenames with special characters**
```bash
# macOS/Linux — quote the filename
./convert.sh "presentation (with spaces).pptx"
./convert.sh "(한글) 파일명.pptx"

# Windows — quote in Command Prompt or PowerShell
convert.bat "presentation (with spaces).pptx"

# Or call Python directly
python scripts/convert_pptx_to_html_v2.py "(한글) 파일명.pptx" output/
```

**Charts not rendering** — the generated HTML loads Chart.js from a CDN; check network access and the browser console.

**Custom shapes appear as rectangles** — uncommon presets and skipped `arcTo` segments fall back to rectangles; check `conversion.log` for warnings.

**Fonts look different** — the source deck likely doesn't embed its fonts; the converter falls back to web-safe stacks. Embed fonts in PowerPoint before converting to preserve typefaces.

**High memory usage on large decks** — lower the DPI (`python … presentation.pptx output/ 96`) or convert in batches.

## Project Structure

```
pptx-to-html-updated/
├── scripts/
│   ├── convert_pptx_to_html_v2.py   # Main converter + HTML/CSS/JS bundle generator
│   ├── chart_extractor.py           # Chart XML → Chart.js configs
│   ├── shape_geometry.py            # Preset + custom geometry → SVG paths
│   ├── smartart_parser.py           # SmartArt text extraction
│   ├── animation_handler.py         # Animations / shadows / reflections → CSS
│   ├── font_manager.py              # Embedded fonts (ODTTF) → WOFF + @font-face
│   └── logger.py                    # Logging, summary, markdown report
├── docs/                            # Architecture docs + landing page (GitHub Pages)
├── convert.sh / convert.bat         # Platform wrappers
├── QUICKSTART.md                    # Fast-path setup guide
├── SKILL.md                         # Agent-skill manifest (Claude Code & friends)
└── README.md                        # This file
```

## Using as an Agent Skill

The repo ships a `SKILL.md` manifest, so coding agents that discover skills (Claude Code et al.) can pick it up directly: point the agent at this repository, and "convert this .pptx to HTML" dispatches the v2 converter with no extra setup.

## Documentation

- **🌐 Landing page** — [cskwork.github.io/pptx-to-html-updated](https://cskwork.github.io/pptx-to-html-updated/)
- **[SKILL.md](SKILL.md)** — agent-facing usage and triggers
- **[docs/architecture.md](docs/architecture.md)** — technical architecture
- **[docs/changelog.md](docs/changelog.md)** — version history

## Contributing

1. Keep the modular pipeline (extractor modules stay independent).
2. Log comprehensively — the report/log files are the debugging surface.
3. Handle failures per-element, never per-deck (one bad shape shouldn't kill a conversion).
4. Update `docs/changelog.md` with every change.
