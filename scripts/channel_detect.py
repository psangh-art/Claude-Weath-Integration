#!/usr/bin/env python3
"""Detect parallel-channel (trendline) boundaries drawn on a TradingView chart
screenshot. Port of the algorithm specified in Claude_Code_Handoff_Instructions.md
section 4 — this is the "actual working method" from prior manual sessions, refined
to cross-validate and reject rather than guess.

Requires the Tesseract OCR binary (not just the `pytesseract` pip package) on PATH,
or at TESSERACT_CMD env var. Install: https://github.com/UB-Mannheim/tesseract/wiki
(on Windows, a one-time interactive install — this cannot be scripted unattended
because the installer needs a UAC prompt).

Usage:
  python channel_detect.py <image_path>                  -> one JSON result
  python channel_detect.py --batch <manifest.json>        -> JSON list of results
    manifest: [{"ticker": str, "screenshot": str}, ...]
    output:   [{"ticker": str, "screenshot": str, "lower": float|None,
                "upper": float|None, "x_frac": float|None, "reason": str|None}, ...]
"""
import sys
import os
import json

try:
    from PIL import Image
    import pytesseract
    import numpy as np
except ImportError as e:
    print(json.dumps({"error": f"missing dependency: {e}. Run: pip install pillow pytesseract numpy"}), file=sys.stderr)
    sys.exit(1)

if os.environ.get('TESSERACT_CMD'):
    pytesseract.pytesseract.tesseract_cmd = os.environ['TESSERACT_CMD']


def check_tesseract_available():
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def is_channel_blue(r, g, b):
    # TradingView's channel-line blue, empirically: R 15-60, G 60-110, B 190-255.
    # Distinct from candle colours (teal ~0,150,120 and red ~240,50,60).
    return 15 <= r <= 60 and 60 <= g <= 110 and 190 <= b <= 255


def read_channel(image_path):
    """Returns (lower_boundary, upper_boundary, x_frac_used, reason) — reason is
    None on success, or a short string explaining why detection was rejected/failed.
    Never guess — a rejection is a valid, expected, and safe outcome."""
    img = Image.open(image_path).convert('RGB')
    arr = np.array(img)
    h, w, _ = arr.shape

    # 1. OCR the right-hand price axis.
    axis_crop = img.crop((int(w * 0.85), 0, w, h))
    data = pytesseract.image_to_data(axis_crop, output_type=pytesseract.Output.DICT)
    labels = []
    for i in range(len(data['text'])):
        txt = data['text'][i].strip().replace(',', '')
        try:
            val = float(txt)
            y_center = data['top'][i] + data['height'][i] / 2
            labels.append((val, y_center))
        except ValueError:
            continue
    labels.sort(key=lambda x: x[1])  # sort top-to-bottom

    if not labels:
        return None, None, None, 'no OCR-readable axis labels'

    # 2. Filter stray OCR noise — axis labels should be strictly monotonic
    #    decreasing top-to-bottom. A common failure: an x-axis year label (e.g.
    #    "2028") gets caught in the crop and breaks a naive endpoint-only fit.
    clean = [labels[0]]
    for val, y in labels[1:]:
        if val < clean[-1][0]:
            clean.append((val, y))
    if len(clean) < 3:
        return None, None, None, 'fewer than 3 clean axis labels after noise filtering'

    # 3. Least-squares fit across ALL clean labels, not just two endpoints.
    vals = np.array([c[0] for c in clean])
    ys = np.array([c[1] for c in clean])
    a, b = np.polyfit(ys, vals, 1)  # price = a*y + b

    # 4/5. Scan multiple x-positions right-to-left. Require EXACTLY 2 line
    #    clusters — more means an extra overlay is present (ambiguous), fewer
    #    means no channel is drawn.
    for x_frac in [0.85, 0.80, 0.75, 0.70, 0.60, 0.50, 0.40, 0.30, 0.20, 0.10]:
        x = int(w * x_frac)
        rows = [y for y in range(h) if is_channel_blue(*arr[y, x])]
        clusters = []
        for y in rows:
            if clusters and y - clusters[-1][-1] <= 3:
                clusters[-1].append(y)
            else:
                clusters.append([y])
        centers = sorted(sum(c) / len(c) for c in clusters)
        if len(centers) == 2:
            lower_price = a * centers[-1] + b
            upper_price = a * centers[0] + b
            if lower_price > 0 and lower_price < upper_price:
                return round(lower_price, 2), round(upper_price, 2), x_frac, None

    return None, None, None, 'no x-position found with exactly 2 channel-blue line clusters'


def plausibility_filter(lower, upper, known_price=None):
    """Returns None if the reading passes all checks, else a rejection reason string.
    Applied AFTER read_channel() succeeds, before the caller trusts the result."""
    if lower is None or upper is None:
        return 'no channel detected'
    width_pct = (upper - lower) / lower
    if width_pct < 0.08:
        return f'width filter: {width_pct:.1%} < 8% (likely noise, not the real channel)'
    if width_pct > 1.50:
        return f'width filter: {width_pct:.1%} > 150% (likely unrelated lines picked up)'
    if known_price is not None:
        # Real price should fall within or very near the detected channel.
        margin = (upper - lower) * 0.15
        if known_price < lower - margin or known_price > upper + margin:
            return f'known price {known_price} is physically implausible vs detected channel [{lower}, {upper}]'
    return None


def process_one(ticker, screenshot_path, known_price=None):
    if not screenshot_path or not os.path.exists(screenshot_path):
        return {'ticker': ticker, 'screenshot': screenshot_path, 'lower': None, 'upper': None,
                 'x_frac': None, 'reason': 'screenshot file not found'}
    lower, upper, x_frac, reason = read_channel(screenshot_path)
    if reason is None:
        reason = plausibility_filter(lower, upper, known_price)
        if reason is not None:
            lower, upper, x_frac = None, None, None
    return {'ticker': ticker, 'screenshot': screenshot_path, 'lower': lower, 'upper': upper,
             'x_frac': x_frac, 'reason': reason}


def main():
    if not check_tesseract_available():
        print(json.dumps({
            "error": "Tesseract OCR binary not found on PATH (pytesseract is just a wrapper "
                     "around it). One-time manual install required: "
                     "https://github.com/UB-Mannheim/tesseract/wiki — this needs an interactive "
                     "UAC prompt and cannot be scripted from an unattended run. After installing, "
                     "either add it to PATH or set the TESSERACT_CMD environment variable to its "
                     "full exe path."
        }), file=sys.stderr)
        sys.exit(1)

    if len(sys.argv) >= 3 and sys.argv[1] == '--batch':
        with open(sys.argv[2], 'r', encoding='utf-8') as f:
            manifest = json.load(f)
        results = [process_one(item.get('ticker'), item.get('screenshot'), item.get('known_price')) for item in manifest]
        print(json.dumps(results, indent=2))
    elif len(sys.argv) >= 2:
        result = process_one(None, sys.argv[1])
        print(json.dumps(result, indent=2))
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == '__main__':
    main()
