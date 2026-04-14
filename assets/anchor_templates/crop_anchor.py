"""
Crop an anchor template from a screenshot for Star Citizen Scanner Tool.

Usage:
    python crop_anchor.py screenshot.png x y w h output.png

Where:
    screenshot.png - path to your screenshot
    x, y           - top-left coordinates of the anchor in the screenshot
    w, h           - width and height of the anchor region
    output.png     - path to save the cropped anchor template
"""
import sys
from PIL import Image

def crop_anchor(screenshot_path, x, y, w, h, output_path):
    img = Image.open(screenshot_path)
    cropped = img.crop((x, y, x + w, y + h))
    cropped.save(output_path)
    print(f"Anchor template saved to {output_path}")

if __name__ == "__main__":
    if len(sys.argv) != 7:
        print("Usage: python crop_anchor.py screenshot.png x y w h output.png")
        sys.exit(1)
    screenshot_path = sys.argv[1]
    x = int(sys.argv[2])
    y = int(sys.argv[3])
    w = int(sys.argv[4])
    h = int(sys.argv[5])
    output_path = sys.argv[6]
    crop_anchor(screenshot_path, x, y, w, h, output_path)
