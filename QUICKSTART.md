# Quick start: Phase 2 PPTX to HTML converter

## Five-minute setup

### 1. Install dependencies

```bash
python3 -m venv path/to/venv
source path/to/venv/bin/activate
pip install -r requirements.txt
```

Expected output:
```
Successfully installed python-pptx-0.6.23 openpyxl-3.1.2
```

### 2. Test the converter

```bash
# Create a test output directory
mkdir -p test_output

# Convert a presentation (use your own PPTX file)
python scripts/convert_pptx_to_html_v2.py /path/to/your/presentation.pptx test_output/

# For filenames with spaces or special characters, use quotes:
python scripts/convert_pptx_to_html_v2.py "presentation (with spaces).pptx" test_output/
python scripts/convert_pptx_to_html_v2.py "(한글) 파일명.pptx" test_output/
```

### 3. View results

Open the generated HTML file in your browser:
```bash
open test_output/presentation.html
```

## What you'll see

### Console output
```
INFO: Initializing Enhanced PPTX Converter (DPI: 150)
INFO: Starting conversion: presentation.pptx
INFO: Found 10 slide(s)
INFO: Processing slide 1...
INFO: Processing slide 2...
...
INFO: Conversion complete: test_output/presentation.html

============================================================
📊 CONVERSION SUMMARY
============================================================
⏱️  Duration: 12.34 seconds
📄 Slides processed: 10
🔧 Total elements: 85
   ├─ 📊 Charts: 4
   ├─ 📋 Tables: 2
   ├─ 🎨 Custom shapes: 12
   ├─ 🔷 SmartArt: 1
   └─ 🖼️  Media files: 23

✅ Conversion completed successfully!
============================================================
```

### Generated files

```
test_output/
├── presentation.html              # Open this in browser
├── presentation_report.md         # Detailed statistics
├── conversion.log                 # Full logs
└── assets/
    ├── slide1_img_rId2.png
    ├── slide2_video_rId5.mp4
    └── ...
```

## Navigation

- **Arrow Keys** (← →): Previous/Next slide
- **Space Bar**: Next slide
- **On-screen Buttons**: Click Previous/Next
- **Progress Bar**: Shows current position

## Features to look for

### Charts
Bar, line, and pie charts render through Chart.js.

### Custom shapes
Arrows, flowchart elements, and complex shapes render as SVG.

### SmartArt
SmartArt diagrams show their text content with the hierarchy intact.

### Animations
Elements may fade or slide in when you navigate to a slide.

### Effects
Shapes with shadows and reflections keep them.

## Customization

### Change DPI quality

```bash
# Higher quality (larger files)
python scripts/convert_pptx_to_html_v2.py input.pptx output/ 300

# Standard quality (smaller files)
python scripts/convert_pptx_to_html_v2.py input.pptx output/ 96
```

### Enable detailed logging

```bash
python scripts/convert_pptx_to_html_v2.py input.pptx output/
# Check output/conversion.log for details
```

## Troubleshooting

### Shell script errors with special characters
If you see errors like `zsh: unknown file attribute` or `bash: syntax error`:

```bash
# ❌ Wrong (causes errors with special characters)
./convert.sh ./(동아출판) 파일명.pptx

# ✅ Correct (use quotes)
./convert.sh "(동아출판) 파일명.pptx"
./convert.sh "presentation (with spaces).pptx"

# ✅ Or use Python directly (no quoting issues)
python scripts/convert_pptx_to_html_v2.py "(동아출판) 파일명.pptx" output/
```

### "Module not found" error
```bash
# Make sure you installed dependencies
pip install -r requirements.txt
```

### Charts not rendering
- Check internet connection (Chart.js loads from CDN)
- Open browser console (F12) for JavaScript errors
- Verify chart was extracted (check logs for "Extracted chart")

### Shapes look wrong
- Some complex custom shapes may approximate
- Use preset shapes (arrows, flowcharts) for best results
- Check conversion log for warnings

### Python version issues
```bash
# Check Python version (need 3.7+)
python --version

# Use specific Python version if needed
python3.9 scripts/convert_pptx_to_html_v2.py input.pptx output/
```

## Next steps

- `README.md` covers the full feature set.
- `docs/architecture.md` explains how the pipeline works.
- `docs/changelog.md` lists what changed in each version.
- `SKILL.md` is the API reference and the longer troubleshooting guide.

## Quick examples

### Python API

```python
from scripts.convert_pptx_to_html_v2 import EnhancedPPTXToHTMLV2

# Basic conversion
converter = EnhancedPPTXToHTMLV2('input.pptx', 'output/')
result = converter.convert()

# With options
converter = EnhancedPPTXToHTMLV2(
    'input.pptx',
    output_dir='output/',
    dpi=150,
    log_file='output/conversion.log'
)
result = converter.convert()
```

### Batch conversion

```bash
# Convert multiple presentations
for file in presentations/*.pptx; do
    python scripts/convert_pptx_to_html_v2.py "$file" output/
done
```

## Performance tips

- Pick the DPI for the job: 72 for speed, 150 for balance, 300 for quality.
- Delete the output directory between runs.
- Read `conversion.log` for warnings.
- Convert a single slide first when you are trying something new.

## When something goes wrong

Read `conversion.log` for the detailed error and `presentation_report.md` for
the per-slide statistics. `README.md` and `SKILL.md` cover the rest. If you can
reproduce the problem with a sample presentation, that makes it much easier to
diagnose.

---

Start with a simple presentation. The Phase 2 additions to look for are charts
via Chart.js, custom shapes as SVG, animations in CSS, 150 DPI images, and a
full log plus report for every run.
