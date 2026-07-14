#!/usr/bin/env python3
"""Detect parallel-channel (trendline) boundaries drawn on a TradingView chart
screenshot. Port of the algorithm specified in Claude_Code_Handoff_Instructions.md
section 4 — this is the "actual working method" from prior manual sessions, refined
to cross-validate and reject rather than guess.

Refinement (user decision 2026-07-13): boundaries are read AT TODAY'S DATE — the
rightmost candle's x-position — by straight-line-fitting each drawn boundary across
many sample x-positions and evaluating the fit there. The original scan took the
first clean sample right-to-left, which on these charts was the blank future space
right of the last candle, i.e. a projected-forward boundary that overstated Alert
Low/High on ascending channels.

Requires the Tesseract OCR binary (not just the `pytesseract` pip package) on PATH,
or at TESSERACT_CMD env var. Install: https://github.com/UB-Mannheim/tesseract/wiki
(on Windows, a one-time interactive install — this cannot be scripted unattended
because the installer needs a UAC prompt).

Also handles the case where a chart has a single trendline rather than a parallel
channel: whichever side of the current price ("known_price") it falls on decides
whether it's used as Alert Low or Alert High (kind: "single_low"/"single_high").
Never guesses direction without a known_price to compare against.

Usage:
  python channel_detect.py <image_path>                  -> one JSON result
  python channel_detect.py --batch <manifest.json>        -> JSON list of results
    manifest: [{"ticker": str, "screenshot": str, "known_price": float|None}, ...]
    output:   [{"ticker": str, "screenshot": str,
                "kind": "parallel"|"single_low"|"single_high"|None,
                "lower": float|None, "upper": float|None,
                "x_frac": float|None, "reason": str|None}, ...]
"""
import sys
import os
import re
import json

# Pin Tesseract's OpenMP threading to a single thread BEFORE it runs. Tesseract's
# LSTM engine is otherwise non-deterministic across runs on marginal (faint/small)
# axis text: the exact same screenshot flipped between "readable" and "no OCR-
# readable labels" run-to-run, and once even emitted a wrong channel. Single-
# threaded LSTM inference is reproducible. setdefault() so a caller can override.
os.environ.setdefault('OMP_THREAD_LIMIT', '1')
os.environ.setdefault('OMP_NUM_THREADS', '1')

try:
    from PIL import Image
    import pytesseract
    import numpy as np
except ImportError as e:
    print(json.dumps({"error": f"missing dependency: {e}. Run: pip install pillow pytesseract numpy"}), file=sys.stderr)
    sys.exit(1)

if os.environ.get('TESSERACT_CMD'):
    pytesseract.pytesseract.tesseract_cmd = os.environ['TESSERACT_CMD']


# Macro/reference charts (stock-index, government-bond-yield, FX-pair) carry a
# price axis Tesseract can't read (small decimals / index scales) AND are not
# buy-list instruments that need channel alerts — a channel read on them is
# neither achievable nor wanted, so they're skipped rather than counted as OCR
# failures. Their last price is still captured separately by the export step.
_CURRENCIES = {'USD', 'EUR', 'GBP', 'JPY', 'CHF', 'AUD', 'CAD', 'NZD',
               'CNY', 'CNH', 'HKD', 'SEK', 'NOK', 'SGD'}
_EXPLICIT_MACRO = {'NASDAQ', 'DJI', 'DXY', 'VIX'}


def is_macro_reference(ticker):
    """True for index / bond-yield / FX-pair symbols that should be excluded from
    channel detection. Deliberately narrow — equity indices actually tracked for a
    channel (e.g. SPX) and commodities (GOLD/SILVER/…) are NOT excluded."""
    t = (ticker or '').upper().strip()
    if not t:
        return False
    if re.match(r'^[A-Z]{2,4}\d{1,2}Y$', t):          # bond yields: US10Y, JP10Y, DE30Y
        return True
    if re.fullmatch(r'[A-Z]{6}', t) and t[:3] in _CURRENCIES and t[3:] in _CURRENCIES:
        return True                                    # FX pairs: GBPUSD, EURUSD, USDJPY
    return t in _EXPLICIT_MACRO


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


def candle_mask(arr):
    """Boolean mask of candle-coloured pixels (TradingView up-candle teal ~#089981,
    down-candle red ~#F23645, with antialiasing tolerance). Deliberately tight so it
    does NOT match the pale-green drawn-label text, muted volume bars, or the teal
    dotted last-price line's axis chip."""
    r = arr[..., 0].astype(np.int32)
    g = arr[..., 1].astype(np.int32)
    b = arr[..., 2].astype(np.int32)
    green = (r <= 60) & (g >= 120) & (g <= 185) & (b >= 100) & (b <= 160)
    red = (r >= 200) & (g >= 30) & (g <= 95) & (b >= 40) & (b <= 105)
    return green | red


def find_today_x(arr, w):
    """x of the rightmost candle column == today's date on the chart. A column must
    have >= 5 candle-coloured pixels so the 1-2px-tall dotted 'last price' line that
    runs from the last candle to the right edge can't masquerade as a candle.

    Scans the FULL width, NOT a hardcoded 0.85w cap. That cap silently assumed the
    last candle sits at ~0.85w, but the right-offset varies per chart: some run
    candles up to the axis (~0.95w), others leave a wide blank future band so the
    last candle is at ~0.65w. The old cap read the channel at frac 0.849 regardless
    — understating ascending channels that actually reach further right (CCH: lower
    rail 3679 vs true 3810) and, where there's blank space, reading the projected-
    forward rail instead of today's.

    Two things on the right are candle-COLOURED but are NOT candles and must be
    excluded: (a) the last-price LABEL CHIP (teal up / red down) that sits in the
    price-axis panel flush to the far edge — it's a fixed ~119px block reaching
    frac ~0.99 on every crop; (b) nothing else matches the tight candle mask. So:
    if the rightmost candle-coloured run reaches the far edge (frac >= 0.975) it's
    the chip — walk left across that solid block to its left edge (the plot/axis
    boundary) and take today as the rightmost real candle strictly left of it.
    Returns None when no candle column is found (blank or non-candle chart)."""
    counts = candle_mask(arr).sum(axis=0)
    xs = np.nonzero(counts >= 5)[0]
    if not len(xs):
        return None
    plot_right = int(xs[-1]) + 1
    if xs[-1] >= 0.975 * w:
        # Rightmost run reaches the far edge -> it's the last-price chip. Walk left
        # across the solid block (internal gaps <=3px) to its left edge; the plot
        # ends there, separated from real candles by the axis-panel margin.
        k = len(xs) - 1
        while k > 0 and xs[k - 1] >= xs[k] - 3:
            k -= 1
        plot_right = int(xs[k])           # chip's left edge
    plot_xs = xs[xs < plot_right]
    if not len(plot_xs):
        return None
    return int(plot_xs[-1])               # rightmost real candle in the plot


def read_channel(image_path, known_price=None):
    """Returns a dict {kind, lower, upper, single_price, x_frac, reason}.
    kind is 'parallel' (two channel-blue lines -> lower+upper both set),
    'single' (exactly one line -> single_price set, direction not yet decided),
    or None (nothing usable found -> reason explains why). Never guess — a
    rejection is a valid, expected, and safe outcome.

    known_price (the current price from the master sheet, if available) hardens the
    axis OCR: it bounds which numeric tokens are plausible price labels and rejects
    an axis whose read doesn't bracket the current price (a systematic misread)."""
    img = Image.open(image_path).convert('RGB')
    arr = np.array(img)
    h, w, _ = arr.shape

    def fail(reason):
        return {'kind': None, 'lower': None, 'upper': None, 'single_price': None, 'x_frac': None, 'reason': reason}

    # 1. OCR the right-hand price axis (full height — the date labels below the plot
    #    are handled by the calendar-year token filter below, so we don't crop them
    #    out here: cropping the bottom strip also dropped legitimate low price ticks
    #    on some charts and shifted the axis fit).
    axis_crop = img.crop((int(w * 0.85), 0, w, h))
    data = pytesseract.image_to_data(axis_crop, output_type=pytesseract.Output.DICT)
    labels = []
    for i in range(len(data['text'])):
        txt = data['text'][i].strip().replace(',', '')
        try:
            val = float(txt)
        except ValueError:
            continue
        if not np.isfinite(val):   # "nan"/"inf" tokens parse as float but aren't prices
            continue
        # Drop calendar-year tokens (2015-2035 integers) that survive the crop,
        # unless a real price genuinely sits in that band (within 20% of known).
        if val == int(val) and 2015 <= val <= 2035 and (
                known_price is None or abs(val - known_price) / known_price > 0.20):
            continue
        # Drop gross outliers far from the current price — these are mis-OCR'd
        # on-chart annotation text (alert-label boxes, ids) rather than axis ticks.
        if known_price is not None and not (known_price / 5.0 <= val <= known_price * 5.0):
            continue
        y_center = data['top'][i] + data['height'][i] / 2
        labels.append((val, y_center))
    labels.sort(key=lambda x: x[1])  # sort top-to-bottom

    if not labels:
        return fail('no OCR-readable axis labels')

    # 2. The price axis is perfectly linear in y, so every genuine label lies on a
    #    single price=a*y+b line, evenly spaced. Fit that line ROBUSTLY (Theil-Sen:
    #    the median of all pairwise slopes) so a couple of junk reads can't leverage
    #    it — an ordinary least-squares fit gets dragged toward an outlier and then
    #    reports a small residual for it, which is exactly how a stray "40"/"124"
    #    misread among real 280..520 ticks used to survive. Keep only labels within
    #    tol of the robust line; tol scales to the axis tick spacing.
    ally = [l[1] for l in labels]
    allv = [l[0] for l in labels]
    pair_slopes = [(allv[j] - allv[i]) / (ally[j] - ally[i])
                   for i in range(len(labels)) for j in range(i + 1, len(labels))
                   if ally[j] != ally[i]]
    if len(labels) >= 3 and pair_slopes:
        slope = float(np.median(pair_slopes))
        intercept = float(np.median([v - slope * y for y, v in zip(ally, allv)]))
        med_gap = float(np.median(np.abs(np.diff(sorted(allv)))))
        tol = max(3.0, 0.3 * med_gap)
        clean = [(v, y) for v, y in labels if abs(slope * y + intercept - v) <= tol]
    else:
        clean = list(labels)

    if len(clean) < 3:
        return fail('fewer than 3 clean axis labels after noise filtering')

    # 2c. A trustworthy read must bracket the current price: the last candle is
    #     on-screen, so its price falls between the top and bottom axis labels. If
    #     it doesn't, the OCR'd axis is on the wrong scale (systematic misread) and
    #     must not produce a channel — this is what stops a false channel off an
    #     axis that OCR'd as 280..520 against a true price of 197 (BT.A).
    lbl_vals = [c[0] for c in clean]
    lo_lbl, hi_lbl = min(lbl_vals), max(lbl_vals)
    if known_price is not None and not (lo_lbl * 0.9 <= known_price <= hi_lbl * 1.1):
        return fail(f'OCR axis labels [{lo_lbl:g}-{hi_lbl:g}] do not bracket known price '
                    f'{known_price:g} — axis read untrustworthy')

    # 3. Least-squares fit across ALL clean labels, not just two endpoints.
    vals = np.array([c[0] for c in clean])
    ys = np.array([c[1] for c in clean])
    a, b = np.polyfit(ys, vals, 1)  # price = a*y + b

    # 4. Sample the drawn line(s) at many x-positions. At each x, cluster the
    #    channel-blue pixels: EXACTLY 2 clusters = both boundaries cleanly visible
    #    there, EXACTLY 1 = a single trendline (or one boundary). More than 2
    #    (extra overlay / the channel's dashed midline) and 0 are skipped.
    #    Sampling deliberately stays in the CLEAN region (<=0.85w): nearer today the
    #    candle mass occludes the rails and the surviving blue fragments mispair
    #    (midline paired with a rail), which poisons the fit. The rails are straight
    #    lines, so we fit here and EXTRAPOLATE to today's x (step 5) — the target x
    #    is corrected below; the sample region is not.
    today_x = find_today_x(arr, w)
    samples2 = []  # (x, upper_y, lower_y)
    samples1 = []  # (x, y)
    for i in range(17):
        x = int(w * (0.85 - i * 0.05))
        rows = [y for y in range(h) if is_channel_blue(*arr[y, x])]
        clusters = []
        for y in rows:
            if clusters and y - clusters[-1][-1] <= 3:
                clusters[-1].append(y)
            else:
                clusters.append([y])
        centers = sorted(sum(c) / len(c) for c in clusters)
        # A dashed/antialiased single line can split into two clusters a few px
        # apart — merge them back into one line; two real channel boundaries are
        # never this close (a <=10px-wide "channel" would fail the 8% width
        # filter anyway, and losing the single-trendline read with it).
        if len(centers) == 2 and centers[1] - centers[0] <= 10:
            centers = [(centers[0] + centers[1]) / 2]
        if len(centers) == 2:
            samples2.append((x, centers[0], centers[-1]))
        elif len(centers) == 1:
            samples1.append((x, centers[0]))

    # 5. Read the boundaries AT TODAY'S DATE (the last candle's x-position), not
    #    wherever a clean sample happened to be. The old first-clean-hit scan read
    #    the channel in the blank future space right of the last candle — i.e. a
    #    projected-forward boundary, overstating Alert Low/High on ascending
    #    channels (user decision 2026-07-13: reads must be at today's date). The
    #    lines are usually occluded by candles AT today's x itself, so fit each
    #    boundary's straight line through its clean samples and evaluate the fit
    #    at today's x. (today_x was computed above, before sampling.)

    # Extrapolating a fit is only trustworthy when the samples genuinely pin down
    # one straight line: at least 3 of them (2 points always fit perfectly, so a
    # mispaired 2-sample "fit" extrapolates anywhere — real captures produced
    # negative prices this way), a small residual (inconsistent cluster pairing
    # across x-positions poisons the fit), and a fitted y that lands in/near the
    # frame (a drawn line can't be outside the pane it was drawn on). Otherwise
    # fall back to the actual sample nearest today — a real pixel read, just at
    # the closest position we could cleanly see the line.
    MAX_RESIDUAL_PX = 6.0

    def read_line_at(pts, target_x):
        """(x, y) samples of one drawn line -> (y at target_x or nearest-sample y,
        x actually used)."""
        if len(pts) >= 3:
            xs = np.array([p[0] for p in pts], dtype=float)
            ys = np.array([p[1] for p in pts], dtype=float)
            slope, intercept = np.polyfit(xs, ys, 1)
            residual = float(np.max(np.abs(slope * xs + intercept - ys)))
            y = float(slope * target_x + intercept)
            if residual <= MAX_RESIDUAL_PX and -0.1 * h <= y <= 1.1 * h:
                return y, target_x
        nearest = min(pts, key=lambda p: abs(p[0] - target_x))
        return float(nearest[1]), nearest[0]

    if samples2:
        target_x = today_x if today_x is not None else max(s[0] for s in samples2)
        upper_y, used_ux = read_line_at([(s[0], s[1]) for s in samples2], target_x)
        lower_y, used_lx = read_line_at([(s[0], s[2]) for s in samples2], target_x)
        lower_price = a * lower_y + b
        upper_price = a * upper_y + b
        if lower_price > 0 and lower_price < upper_price:
            return {'kind': 'parallel', 'lower': round(lower_price, 2), 'upper': round(upper_price, 2),
                    'single_price': None, 'x_frac': round((used_ux + used_lx) / 2 / w, 3), 'reason': None}
        # Non-positive or inverted prices mean the blue marks we clustered are not
        # a drawn channel at all (typically bottom-of-pane UI icons that happen to
        # match channel blue) — same charts the old scan rejected as "no
        # x-position found".
        return fail(f'channel-blue marks do not resolve to a plausible channel at today\'s date '
                    f'(lower {lower_price:.2f} vs upper {upper_price:.2f} — likely UI elements, not a drawn channel)')

    if samples1:
        target_x = today_x if today_x is not None else max(s[0] for s in samples1)
        y, used_x = read_line_at(samples1, target_x)
        price = a * y + b
        if price > 0:
            return {'kind': 'single', 'lower': None, 'upper': None,
                    'single_price': round(price, 2), 'x_frac': round(used_x / w, 3), 'reason': None}

    return fail('no x-position found with exactly 1 or 2 channel-blue line clusters')


def plausibility_filter(kind, lower, upper, known_price=None):
    """Returns None if the reading passes all checks, else a rejection reason string.
    Applied AFTER read_channel()/direction-resolution succeeds, before the caller
    trusts the result."""
    if kind == 'parallel':
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
    if kind in ('single_low', 'single_high'):
        price = lower if kind == 'single_low' else upper
        if known_price is None:
            return 'single trendline detected but no current price available to determine Alert Low vs Alert High'
        # A single trendline should sit reasonably close to the current price —
        # this is a looser sanity check than the parallel-channel width filter
        # (there's no second boundary to cross-validate against), just enough
        # to catch a line that's obviously unrelated to this instrument.
        ratio = price / known_price
        if ratio < 0.3 or ratio > 3.0:
            return f'single trendline price {price} is implausibly far from known price {known_price} (ratio {ratio:.2f})'
        return None
    return 'no channel or trendline detected'


def process_one(ticker, screenshot_path, known_price=None):
    if is_macro_reference(ticker):
        return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': None, 'lower': None, 'upper': None,
                 'x_frac': None,
                 'reason': 'macro/reference instrument (index/yield/FX) — not a buy-list channel, skipped'}
    if not screenshot_path or not os.path.exists(screenshot_path):
        return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': None, 'lower': None, 'upper': None,
                 'x_frac': None, 'reason': 'screenshot file not found'}

    raw = read_channel(screenshot_path, known_price)
    kind, lower, upper, x_frac, reason = raw['kind'], raw['lower'], raw['upper'], raw['x_frac'], raw['reason']

    if reason is None and kind == 'single':
        # A lone trendline: decide Alert Low vs Alert High by which side of the
        # current price it falls on. Never guess when we don't have a price to
        # compare against — reject instead.
        if known_price is None:
            kind, reason = None, 'single trendline detected but no current price available to determine Alert Low vs Alert High'
        elif raw['single_price'] < known_price:
            kind, lower, upper = 'single_low', raw['single_price'], None
        else:
            kind, lower, upper = 'single_high', None, raw['single_price']

    if reason is None:
        reason = plausibility_filter(kind, lower, upper, known_price)
        if reason is not None:
            kind, lower, upper, x_frac = None, None, None, None

    return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': kind, 'lower': lower, 'upper': upper,
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
