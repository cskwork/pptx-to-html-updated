# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PowerPoint to HTML Converter (Phase 2) - A production-ready Python tool that converts .pptx presentations to standalone HTML files with 98%+ visual fidelity. Supports charts, custom shapes, SmartArt, animations, tables, videos, and more.

## Essential Commands

### Installation & Setup
```bash
# Install dependencies (required first time)
pip install -r requirements.txt

# Dependencies: python-pptx, openpyxl, fonttools
```

### Running Conversions

```bash
# Phase 2 converter (recommended - full features)
python scripts/convert_pptx_to_html_v2.py <input.pptx> <output_dir> [dpi]

# Examples:
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/ 300  # Higher quality
python scripts/convert_pptx_to_html_v2.py presentation.pptx output/ 96   # Smaller files
```

### No Test Framework Currently
There is no test suite configured. When adding tests:
- Use pytest framework
- Place tests in a `tests/` directory
- Follow pattern: `test_<module_name>.py`

## Architecture Overview

### Modular Design (Phase 2)

The converter follows a **separation of concerns** architecture with specialized modules:

```
EnhancedPPTXToHTMLV2 (Main Orchestrator)
    ├── ConversionLogger         → Logging & statistics tracking
    ├── ChartExtractor           → PowerPoint charts → Chart.js
    ├── ShapeGeometryConverter   → Custom shapes → SVG paths
    ├── SmartArtParser           → SmartArt diagrams → HTML hierarchy
    ├── AnimationHandler         → Animations → CSS/JavaScript
    └── FontManager              → Font extraction & embedding
```

**Key principle:** Each module handles ONE responsibility. The main converter (`convert_pptx_to_html_v2.py`) orchestrates, but delegates all specialized tasks to dedicated modules.

### Critical Architectural Patterns

1. **Two-Phase DPI Handling**
   - `layout_dpi = 96`: Fixed for browser rendering (positioning, sizing)
   - `image_dpi = 150`: Configurable for image export quality
   - **Never mix these in calculations!** Use `emu_to_layout_px()` for positioning

2. **Graceful Degradation**
   ```python
   try:
       # Attempt Phase 2 feature (chart, custom shape, etc.)
       chart = extract_chart(...)
   except Exception as e:
       logger.warning("Feature extraction failed", exception=e)
       # Continue conversion, skip failed element
   ```
   **Critical:** Never let a single element failure stop the entire conversion

3. **Relationship File Caching**
   ```python
   self.rels_cache = {}  # Prevents repeated XML parsing
   ```
   PPTX files have complex relationship webs - cache aggressively

4. **Z-index Management**
   - Elements layer using incrementing z-index
   - Counter in `self.z_index_counter`
   - Call `_next_z_index()` for each element

### Data Flow (Conversion Process)

```
1. PPTX Unpacking (ZipFile) → XML parsing
2. For each slide:
   - Extract background (solid/gradient)
   - Extract animations
   - Process shapes:
     - Regular shapes (p:sp) → position, fill, text, hyperlinks
     - Graphic frames (p:graphicFrame):
       - Charts → ChartExtractor → Chart.js canvas
       - Tables → HTML table elements
       - SmartArt → SmartArtParser → text hierarchy
     - Custom geometries → ShapeGeometryConverter → SVG
3. Generate HTML:
   - Slide containers with backgrounds
   - Positioned elements (absolute positioning)
   - Chart.js initialization scripts
   - CSS animations
   - Navigation controls
4. Output:
   - presentation.html
   - presentation_report.md (statistics)
   - conversion.log
   - assets/ directory (images, videos, audio)
```

## File Organization

```
scripts/
├── convert_pptx_to_html_v2.py    # Main converter (Phase 2) - START HERE
├── logger.py                      # Logging system with statistics
├── chart_extractor.py             # Chart → Chart.js conversion
├── shape_geometry.py              # DrawingML → SVG path conversion
├── smartart_parser.py             # SmartArt text extraction
├── animation_handler.py           # Animations + shadow/reflection
└── font_manager.py                # Font extraction (Phase 3)

docs/
├── architecture.md                # Detailed technical architecture (Korean)
├── changelog.md                   # Version history
└── high_fidelity_plan.md          # Feature planning

SKILL.md                           # Complete feature reference for Claude Code skill
README.md                          # User-facing documentation
QUICKSTART.md                      # 5-minute setup guide
```

## XML Namespaces (Critical Reference)

```python
ns = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
}
```

**Always use namespace prefixes in XPath queries:**
```python
shape.find('.//a:rPr', ns)  # ✅ Correct
shape.find('.//rPr')        # ❌ Will fail silently
```

## Core Conversion Concepts

### 1. EMU to Pixel Conversion

```python
# EMU (English Metric Units) = PowerPoint's coordinate system
# Formula: px = (emu / 914400) * dpi

# Layout positioning (always 96 DPI)
self.emu_to_layout_px(emu)  # Use this for position, width, height

# Image export (configurable DPI)
self.emu_to_px(emu, self.image_dpi)  # Use this for image extraction
```

### 2. Element Extraction Pattern

Every element follows this pattern in `convert_pptx_to_html_v2.py`:
1. Extract position (`extract_shape_position`)
2. Extract fill color/gradient (`extract_shape_fill`)
3. Extract border (`extract_shape_border`)
4. Extract text with formatting (`extract_text_with_formatting`)
5. Check for special features (hyperlinks, images, videos)
6. Generate HTML with absolute positioning

### 3. Relationship Resolution

```python
# PPTX files store media references in relationship files
# Pattern: ppt/slides/_rels/slide1.xml.rels
# Links shapes to images/videos/charts via rId (relationship IDs)

def get_relationships(rels_path):
    if rels_path in self.rels_cache:  # Always cache!
        return self.rels_cache[rels_path]
    # Parse XML, cache result
```

### 4. Color Handling

PowerPoint uses multiple color systems:
- **RGB Colors**: `<a:srgbClr val="FF5733"/>` → #FF5733
- **Theme Colors**: `<a:schemeClr val="accent1"/>` → Map to `self.theme_colors`
- **Gradients**: Multiple `<a:gs>` stops with positions

### 5. Chart.js Integration

Charts use Chart.js 4.4.1 loaded from CDN:
```javascript
// Generated in HTML output
<canvas id="chart_1"></canvas>
<script>
new Chart(document.getElementById('chart_1'), {
    type: 'bar',  // Detected from PowerPoint chart type
    data: { /* extracted from chart XML */ }
});
</script>
```

**Supported chart types:**
- Bar (2D/3D) → `barChart`
- Line (2D/3D) → `lineChart`
- Pie/Doughnut → `pieChart`, `doughnutChart`
- Area, Scatter, Radar, Bubble

## Phase 2 Features (Current Implementation)

### Fully Implemented ✅
- Text formatting (fonts, colors, bold, italic, underline)
- Shape positioning (pixel-accurate with absolute positioning)
- Images (150 DPI, up from 72 DPI in Phase 1)
- Tables (borders, cell colors, text formatting)
- Hyperlinks (text and shape level)
- Videos & Audio (HTML5 players)
- Backgrounds (solid and gradient)
- **Charts** → Chart.js rendering
- **Custom Shapes** → SVG path conversion
- **SmartArt** → Text hierarchy extraction
- **Animations** → CSS keyframes + JavaScript
- **Shadows/Reflections** → CSS box-shadow

### Known Limitations ⚠️
- SmartArt: Only text content, visual layout simplified
- Custom fonts: Fall back to web-safe fonts
- Complex 3D effects: Not preserved
- Master slide templates: Complex inheritance not fully supported
- Macros/VBA: Never supported

## Development Guidelines

### Adding New Features

1. **Create a new module** for complex features (e.g., `new_feature.py`)
2. **Add logging** using `self.logger.info/warning/error`
3. **Implement graceful degradation** with try/except
4. **Update statistics** in `ConversionLogger` if trackable
5. **Document** in both SKILL.md and README.md
6. **Key Insights** - maintain essential insights in CLAUDE.md

### Error Handling Levels

```python
# INFO: Normal operation
self.logger.info("Processing slide 5...")

# WARNING: Feature failed, conversion continues
self.logger.warning("Chart extraction failed", exception=e)

# ERROR: Element processing failed, slide continues
self.logger.error("Shape processing error", exception=e)

# CRITICAL: Conversion must stop
self.logger.critical("PPTX file corrupted")
```

### Code Style Notes

- **Comments in Korean** (per project standard)
- **Docstrings in English or Korean**
- **Variable names in English**
- Use type hints: `def method(arg: str) -> Dict:`

### DPI Configuration

Default: 150 DPI (balance quality/performance)
- 72 DPI: Fast, smaller files
- 96 DPI: Standard web quality
- 150 DPI: **Recommended** default
- 300 DPI: High quality, larger files

## Common Pitfalls

1. **Mixing layout_dpi and image_dpi**
   - Always use `emu_to_layout_px()` for positioning
   - Only use `emu_to_px(emu, image_dpi)` for image extraction

2. **Forgetting namespace prefixes in XPath**
   ```python
   shape.find('.//a:rPr', self.ns)  # Required!
   ```

3. **Not caching relationship files**
   - PPTX files have many relationships
   - Always check `self.rels_cache` first

4. **Hardcoding paths in HTML**
   - Use relative paths: `assets/slide1_img.png`
   - Never absolute paths

5. **Assuming all elements exist**
   - PowerPoint XML is sparse (missing = default)
   - Always use `.get()` or `find()` with None checks

## Output Structure

```
output-directory/
├── presentation.html              # Main file (open in browser)
├── presentation_report.md         # Statistics and warnings
├── conversion.log                 # Detailed logs
└── assets/
    ├── slide1_img_rId2.png        # Images (150 DPI)
    ├── slide2_chart_1.json        # Chart data (if needed)
    ├── slide3_video_rId5.mp4      # Videos
    └── slide4_audio_rId7.mp3      # Audio files
```

## When Making Changes

1. **Read architecture.md first** - Has detailed Korean documentation
2. **Test with diverse presentations** - Different shapes, charts, layouts
3. **Check conversion.log** - Warnings indicate partial failures
4. **Verify browser compatibility** - Test in Chrome, Firefox, Safari
5. **Update statistics tracking** - Add counters in `ConversionLogger`

## Dependencies

- **python-pptx** (0.6.23+): PowerPoint file parsing
- **openpyxl** (3.1.2+): Chart data extraction from embedded Excel
- **fonttools** (4.51.0+): Font embedding (Phase 3 - in progress)

Chart.js 4.4.1 loads from CDN in generated HTML (no Python dependency).

## Python Version

Requires Python 3.7+. Uses standard library heavily (zipfile, xml.etree, pathlib).

<insight>
**SVG Transform & Overflow Handling**:
- When converting shapes (especially lines/connectors), PowerPoint rotations are around the center of the bounding box. Always use `transform-origin: center` (CSS default) or explicitly set it, avoiding `top left`.
- Flipped shapes (flipH/flipV) must be handled via CSS transforms (`scaleX(-1)`, `scaleY(-1)`) to preserve directional semantics (e.g., arrows).
- Horizontal/Vertical lines often have 0 height or width. To ensure they render, SVG containers must use `overflow: visible` and strokes must be allowed to extend beyond the viewBox.

**Text & Layout Precision**:
- PowerPoint default text padding (internal margins) is **0.1" (left/right)** and **0.05" (top/bottom)**. Ignoring these and defaulting to 0px causes text misalignment.
- Slide background scaling should use `background-size: 100% 100%` instead of `cover` to ensure exact fit without cropping, preserving alignment with absolute-positioned foreground elements.

**Group Shape Transformations**:
- Group flips (`flipH`, `flipV`) must be applied to the group container itself using CSS `scaleX(-1)` or `scaleY(-1)`. This correctly flips and positions all children within the group context without needing complex individual coordinate re-calculation for every child.
</insight>
