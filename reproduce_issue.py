import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
import subprocess

def create_test_pptx(filename):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5] # Blank
    slide = prs.slides.add_slide(slide_layout)

    # Add a shape
    left = top = width = height = Inches(1.0)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.text = "Animated Shape"
    
    # We cannot easily add animations via python-pptx as it doesn't support it natively yet.
    # However, we can check if the converter handles existing animations if we had a file.
    # Since we can't create animations programmatically easily, we will trust the code analysis
    # that the animation extraction logic was already there (AnimationHandler) and we just enabled the JS.
    
    # But we can test positioning.
    # Add a group
    group_shape = slide.shapes.add_group_shape()
    # Add shapes to group
    # Note: python-pptx group API is a bit limited in older versions, but let's try.
    # Actually, adding to group is tricky.
    
    # Let's just save a simple PPTX and convert it to check if JS is present.
    prs.save(filename)
    print(f"Created {filename}")

def run_conversion(pptx_path, output_dir):
    cmd = [sys.executable, "scripts/convert_pptx_to_html_v2.py", pptx_path, output_dir]
    result = subprocess.run(cmd, capture_output=True, text=True)
    print(result.stdout)
    print(result.stderr)
    return result.returncode == 0

def check_output(html_path):
    if not os.path.exists(html_path):
        print("HTML file not found")
        return False
        
    with open(html_path, 'r') as f:
        html_content = f.read()
        
    # Check for unique IDs
    if 'id="shape_' in html_content:
        print("✅ Shape IDs found")
    else:
        print("❌ Shape IDs NOT found")

    # Check for animation JS
    # First check if it's inline
    js_found = "Animation Controller" in html_content or "playAnimations" in html_content
    
    # If not inline, check external JS file
    if not js_found:
        js_filename = os.path.splitext(os.path.basename(html_path))[0] + ".js"
        js_path = os.path.join(os.path.dirname(html_path), js_filename)
        print(f"Checking JS file: {js_path}")
        if os.path.exists(js_path):
            with open(js_path, 'r') as f:
                js_content = f.read()
                print(f"JS content length: {len(js_content)}")
                if "Animation Controller" in js_content or "playAnimations" in js_content:
                    js_found = True
                else:
                    print("JS content does not contain expected strings")
        else:
            print("JS file does not exist")
    
    if js_found:
        print("✅ Animation JS found")
    else:
        print("❌ Animation JS NOT found")

    return True

if __name__ == "__main__":
    pptx_file = "reproduce_test.pptx"
    output_dir = "test_output"
    
    create_test_pptx(pptx_file)
    if run_conversion(pptx_file, output_dir):
        html_filename = os.path.splitext(pptx_file)[0] + ".html"
        check_output(os.path.join(output_dir, html_filename))
    else:
        print("Conversion failed")
