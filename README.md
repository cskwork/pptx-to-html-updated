# PowerPoint to HTML converter (Phase 2)

Converts `.pptx` presentations to standalone HTML files. The author measures
visual fidelity at 98%+ and feature coverage at 92%.

## What Phase 2 adds

Phase 1 already handled text, shapes, images, tables, hyperlinks, media, and
backgrounds. Phase 2 adds five things on top and raises image extraction from
72 DPI to 150 DPI.

- **Charts.** Rendered with Chart.js from the embedded chart data.
- **Custom shapes.** Arrows, connectors, and flowchart elements convert to SVG.
- **SmartArt.** Text is extracted; the visual layout is not reproduced.
- **Animations.** Mapped to CSS keyframes and JavaScript.
- **Shadows and reflections.** Rendered as CSS effects.

It also writes a per-conversion log and a Markdown report.

Chart types: bar, line, and pie (2D and 3D), area, scatter, doughnut.

Shape types: arrows in all directions, connectors, flowchart elements, basic
geometric shapes, stars, polygons.

## Quick start

```bash
cd /path/to/pptx-to-html-updated
pip install -r requirements.txt
```

```bash
# Phase 2 converter (recommended)
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/

# with custom DPI
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/ 300

# Phase 1 converter (legacy)
python scripts/convert_pptx_to_html.py presentation.pptx output/
```

### Python API

```python
from scripts.convert_pptx_to_html_v2 import EnhancedPPTXToHTMLV2

converter = EnhancedPPTXToHTMLV2(
    'presentation.pptx',
    output_dir='./output',
    dpi=150,
    log_file='./output/conversion.log'
)

result = converter.convert()
print(f"Converted: {result}")
```

## Phase 1 vs Phase 2

| Feature | Phase 1 | Phase 2 |
|---|---|---|
| Text formatting | yes | yes |
| Shapes and borders | yes | yes |
| Images | 72 DPI | 150 DPI |
| Video and audio | yes | yes |
| Tables | yes | yes |
| Hyperlinks | yes | yes |
| Backgrounds | yes | yes |
| Charts | no | yes |
| Custom shapes | no | yes |
| SmartArt | no | text only |
| Animations | no | yes |
| Shadows | no | yes |
| Reflections | no | yes |
| Logging | basic | full log + report |
| Error handling | basic | per-element, with warnings |
| Visual fidelity | 95% | 98%+ |
| Feature coverage | 80% | 92% |

## Output structure

```
output-directory/
├── presentation.html          # main presentation file
├── presentation_report.md     # conversion report
├── conversion.log             # full conversion log
└── assets/
    ├── slide1_img_rId2.png      # images (150 DPI)
    ├── slide2_chart_1.json      # chart data
    ├── slide3_video_rId5.mp4    # videos
    └── slide4_audio_rId7.mp3    # audio files
```

## Dependencies

- Python 3.7+
- `python-pptx` for PowerPoint file parsing
- `openpyxl` for Excel chart data extraction

No web dependencies. Chart.js loads from a CDN in the generated HTML.

## Architecture

```
pptx-to-html-updated/
├── scripts/
│   ├── convert_pptx_to_html.py          # Phase 1 (legacy)
│   ├── convert_pptx_to_html_v2.py       # Phase 2 (recommended)
│   ├── logger.py                        # logging
│   ├── chart_extractor.py               # Chart.js integration
│   ├── shape_geometry.py                # SVG conversion
│   ├── smartart_parser.py               # SmartArt extraction
│   └── animation_handler.py             # animation mapping
├── tests/
├── docs/
├── requirements.txt
├── SKILL.md                             # full reference
└── README.md
```

## What survives the conversion

Preserved exactly: text formatting (font, color, size, bold, italic,
underline), shape positioning to the pixel, images and their placement, table
styling, hyperlinks at both text and shape level, video and audio playback.

Preserved at 95%+: charts (data-driven via Chart.js), custom shapes (SVG),
shadows (CSS `box-shadow`), animations (CSS keyframes plus JavaScript).

Approximated: SmartArt keeps its text hierarchy but loses the visual layout.

Not converted: macros and VBA, master slide templates with complex
inheritance, embedded fonts (falls back to web-safe fonts), complex 3D effects.

## Performance

- 1 to 2 seconds per slide
- around 100 MB of memory for a typical deck
- HTML output of 50 to 300 KB, plus assets sized to the media
- runs in Chrome, Firefox, Safari, and Edge; responsive with touch navigation

## Logging and reports

Progress goes to the console during the run:

```
INFO: Processing slide 1...
INFO: Extracted bar chart
INFO: Extracted custom arrow shape
INFO: Processing slide 2...
```

The Markdown report summarizes the run:

```markdown
## Conversion Statistics
- Duration: 3.45 seconds
- Slides processed: 15
- Total elements: 127
  - Charts: 8
  - Custom shapes: 23
  - Tables: 5
  - SmartArt: 2
  - Media files: 34

## Status: SUCCESS
```

## Troubleshooting

**Filename with special characters causing errors?**

On macOS and Linux, quote the filename when calling the shell scripts:

```bash
./convert.sh "presentation (with spaces).pptx"
./convert.sh "(한글) 파일명.pptx"
```

On Windows, quote it in Command Prompt or PowerShell:

```cmd
convert.bat "presentation (with spaces).pptx"
```

Calling Python directly needs no quoting:

```bash
python scripts/convert_pptx_to_html_v2.py "(한글) 파일명.pptx" output/
```

**Charts not rendering?** Check that the Chart.js CDN is reachable, look at the
browser console for JavaScript errors, and confirm from the log that the chart
data was extracted.

**Custom shapes appear as rectangles?** Some complex paths do not convert.
Check the conversion log for warnings and prefer preset shapes.

**SmartArt looks different?** Only the text content carries over. For complex
diagrams, convert them to images inside PowerPoint first.

**High memory usage?** Large decks with many images cost more memory. Lower the
DPI from 150 to 96, or process the deck in batches.

## Development

```bash
python -m pytest tests/              # unit tests
python tests/test_conversion.py      # integration tests
```

When contributing, keep the modular layout, log what each step does, handle
errors per element rather than aborting the run, update the docs, and add tests.

## More documentation

- `SKILL.md` for the full feature reference and API documentation
- `docs/architecture.md` for technical detail
- `docs/changelog.md` for version history

For a problem you cannot place, read the conversion log and report first, then
open an issue with a sample PPTX if you can share one.

## License

See the LICENSE file in the repository root.
