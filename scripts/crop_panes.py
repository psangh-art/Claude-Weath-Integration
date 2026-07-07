#!/usr/bin/env python3
"""Crop individual chart-pane regions out of one full-layout screenshot.
Usage: python crop_panes.py <source_png> <crops.json> <output_dir>
crops.json: [{"x":.., "y":.., "width":.., "height":.., "filename":"..."}, ...] in
source-image pixel coordinates (already scaled to match the screenshot's actual
pixel dimensions, not CSS px).
"""
import sys
import os
import json
from PIL import Image

def main():
    source_png, crops_path, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(crops_path, 'r', encoding='utf-8') as f:
        crops = json.load(f)

    im = Image.open(source_png)
    os.makedirs(out_dir, exist_ok=True)

    for c in crops:
        box = (int(c['x']), int(c['y']), int(c['x'] + c['width']), int(c['y'] + c['height']))
        im.crop(box).save(os.path.join(out_dir, c['filename']))

    print(f"Cropped {len(crops)} pane(s) from {source_png}")

if __name__ == '__main__':
    main()
