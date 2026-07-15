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


# A drawn line within this fraction of the live price is treated as REACHED rather
# than assigned a side (see process_one). 0.25% is about the resolution limit of
# these screenshots: on a ~1300px-tall pane one pixel is ~0.1% of price, so a line
# 2-3px from the last-price line cannot be honestly called "above" or "below" it.
# User decision 2026-07-15.
ON_ALERT_TOL = 0.0025
# See detect_wedge. Horizon is the load-bearing gate; the gap floor only drops lines
# double-detected on top of each other.
WEDGE_HORIZON_FRAC = 0.15   # apex at most 0.15 pane-widths past today (~3 months)
WEDGE_MIN_GAP_FRAC = 0.03   # lines still >=3% of price apart at today

# Minimum span (fraction of pane width) for a CHANNEL-BLUE line to count as a rail.
# Well above the yellow lines' 0.12: rails are drawn across the chart, while the
# SELL/BUY buttons' 1px blue border is ~100px wide and clears 0.12 on a narrow pane.
# See read_blue_rails. Real rails measured 0.40-0.90; button borders 0.12.
BLUE_MIN_SPAN_FRAC = 0.25


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


def trend_yellow_mask(arr):
    """Boolean mask of the user's hand-drawn yellow trend-line pixels (line core is
    ~#FDEA3B, R 250-255 / G 230-235 / B 55-60, with antialiasing tolerance). Tight
    on the low end so it does NOT match the pale-yellow app icon top-left
    (~#EAC102, G 193) or the amber event-marker chips at the bottom (~#CDA05C,
    R 205). Candles never match (up-green has R<=60; red has G<=95)."""
    r = arr[..., 0].astype(np.int32)
    g = arr[..., 1].astype(np.int32)
    b = arr[..., 2].astype(np.int32)
    return (r >= 220) & (g >= 205) & (b <= 95)


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


def fit_price_axis(img, arr, w, h, known_price=None):
    """OCR the right-hand price axis and fit price = a*y + b. Returns (a, b, None)
    on success, or (None, None, reason) when the axis can't be trusted. Shared by
    blue-channel and yellow-trendline detection so both price the same axis reads
    identically. known_price (if given) hardens the OCR: it bounds which numeric
    tokens are plausible price labels and rejects an axis that doesn't bracket the
    current price (a systematic misread)."""
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
        return None, None, 'no OCR-readable axis labels'

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
        return None, None, 'fewer than 3 clean axis labels after noise filtering'

    # 2c. A trustworthy read must bracket the current price: the last candle is
    #     on-screen, so its price falls between the top and bottom axis labels. If
    #     it doesn't, the OCR'd axis is on the wrong scale (systematic misread) and
    #     must not produce a channel — this is what stops a false channel off an
    #     axis that OCR'd as 280..520 against a true price of 197 (BT.A).
    lbl_vals = [c[0] for c in clean]
    lo_lbl, hi_lbl = min(lbl_vals), max(lbl_vals)
    if known_price is not None and not (lo_lbl * 0.9 <= known_price <= hi_lbl * 1.1):
        return None, None, (f'OCR axis labels [{lo_lbl:g}-{hi_lbl:g}] do not bracket known price '
                            f'{known_price:g} — axis read untrustworthy')

    # 3. Least-squares fit across ALL clean labels, not just two endpoints.
    vals = np.array([c[0] for c in clean])
    ys = np.array([c[1] for c in clean])
    a, b = np.polyfit(ys, vals, 1)  # price = a*y + b
    return float(a), float(b), None


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

    # 1-3. OCR + robust fit of the price axis (shared with yellow-trendline reads).
    a, b, axis_reason = fit_price_axis(img, arr, w, h, known_price)
    if axis_reason is not None:
        return fail(axis_reason)

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


def _extract_straight_lines(mask, w, h, max_lines=6, resid_px=3.0,
                            min_span_frac=0.12, min_coverage=0.6):
    """Extract distinct STRAIGHT drawn lines from a colour mask. Returns a list of
    (slope, intercept, xmin, xmax) in pixel space, one per accepted line.

    Used for the user's hand-drawn yellow trend lines, which are dead-straight
    (measured residual <2px over the full span on USOIL/CRDA/DGE). RANSAC finds the
    dominant line, then a COVERAGE check (fraction of columns across the line's span
    that actually carry a mask pixel within resid_px of it) rejects wavy indicators:
    a yellow EMA is only locally linear, so no single straight line covers a long
    span of it. Inliers are removed and the search repeats for further lines (a
    chart can carry several trend lines, e.g. a drawn channel's two rails)."""
    ys, xs = np.nonzero(mask)
    if len(xs) < 40:
        return []
    pts = np.column_stack([xs, ys]).astype(float)
    rng = np.random.RandomState(0)
    if len(pts) > 4000:
        pts = pts[rng.choice(len(pts), 4000, replace=False)]
    min_span = min_span_frac * w
    lines = []
    for _ in range(max_lines):
        if len(pts) < 40:
            break
        X, Y = pts[:, 0], pts[:, 1]
        best_inl, best_count = None, 0
        for _ in range(400):
            i, j = rng.randint(0, len(pts)), rng.randint(0, len(pts))
            x1, y1 = pts[i]; x2, y2 = pts[j]
            if abs(x2 - x1) < 5:            # need a horizontal span, not two stacked px
                continue
            slope = (y2 - y1) / (x2 - x1)
            intercept = y1 - slope * x1
            inl = np.abs(slope * X + intercept - Y) <= resid_px
            c = int(inl.sum())
            if c > best_count:
                best_count, best_inl = c, inl
        if best_inl is None or best_count < 40:
            break
        xin = X[best_inl]
        span = float(xin.max() - xin.min())
        slope, intercept = np.polyfit(xin, Y[best_inl], 1)   # refit on inliers
        pts = pts[~best_inl]
        if span < min_span:
            continue
        # Coverage: does the line trace a real drawn stroke across its span, or is it
        # just the best fit through a diffuse/ wavy blob?
        cols = range(int(xin.min()), int(xin.max()) + 1, 3)
        hit = tot = 0
        for x in cols:
            tot += 1
            yl = slope * x + intercept
            y0 = max(0, int(yl - resid_px)); y1 = min(h, int(yl + resid_px) + 1)
            if y0 < y1 and mask[y0:y1, x].any():
                hit += 1
        if tot and hit / tot >= min_coverage:
            lines.append((float(slope), float(intercept), float(xin.min()), float(xin.max())))
    return lines


def channel_blue_mask(arr):
    """Boolean mask of channel-blue pixels — the array-wide form of is_channel_blue,
    for the straight-line extractor."""
    r = arr[..., 0].astype(np.int32)
    g = arr[..., 1].astype(np.int32)
    b = arr[..., 2].astype(np.int32)
    return (r >= 15) & (r <= 60) & (g >= 60) & (g <= 110) & (b >= 190) & (b <= 255)


def read_blue_rails(image_path, known_price=None):
    """Prices (at today's date) of the SOLID channel-blue rails, each extracted and
    fitted independently — the same RANSAC treatment as the yellow trend lines.

    Supersedes read_channel()'s column-pairing for ALERT SELECTION (user rule
    2026-07-15). Pairing clustered the blue pixels in each sampled column and
    required exactly 1 or 2 clusters, which assumed both rails are in-frame at the
    same x and nothing else blue shares the column. Neither holds: on a steep
    channel one rail leaves the frame long before today (SILVER's lower rail exits
    at x=586 of 1452), and TradingView's dashed midline / projection lines and the
    blue SMA share columns with the rails. A single mispaired sample then poisons
    the straight-line fit — on SILVER that produced a bogus 52.54/65.56 "channel"
    fitted through the dashed line, so the REAL upper rail (~58.1, and the nearest
    support below price) never became a candidate at all and a distant yellow line
    at 56.97 won by default.

    Extracting each line independently fixes both: the rails are found by their own
    collinearity, the dashed midline is rejected by the coverage check (a dashed
    stroke covers ~half its span), the wavy blue SMA is rejected the same way, and a
    rail that has left the frame simply fails the reach-today test instead of
    corrupting the other rail's fit. Rails are validated exactly like yellow lines.
    Returns a list of prices (see _read_lines_at_today for the geometry behind them);
    direction is decided by the caller from which side of the price each falls on.

    Unlike a yellow trend line, a rail is NOT required to be drawn as far as today:
    a TradingView channel is a projected construct, and its rails are meaningful
    extrapolated forward (that is what the old paired read did, and requiring reach
    here dropped 56 otherwise-good charts). The in-frame test still applies, so a
    rail that has genuinely left the pane by today — SILVER's lower rail, off the
    bottom — is still discarded rather than reported at an off-chart price.

    A rail must also span at least BLUE_MIN_SPAN_FRAC of the pane. The SELL/BUY
    price buttons top-left are drawn with a 1px blue BORDER, and RANSAC happily fits
    that border's horizontal edge as a "line": ~100px long, which on a narrow pane
    (WIZZ's is 820px wide) clears the 12% default the yellow lines use. Thickness
    cannot separate them — the border is 1px where a real rail is ~4px — but span
    can: measured across 60 charts, real rails span 40-90% of the pane and the only
    lines below 25% were exactly the four button borders (WIZZ, TFIF, VOF, JDW).
    WIZZ's button was being published as Alert High 2,984 against a price of 1,106."""
    return [r['price'] for r in _read_lines_at_today(
        image_path, channel_blue_mask, known_price,
        require_reach=False, min_span_frac=BLUE_MIN_SPAN_FRAC)[0]]


def read_yellow_trendlines(image_path, known_price=None):
    """Prices (at today's date) of the user's hand-drawn yellow trend lines. Returns
    a list of floats — empty when the axis can't be read or no straight yellow line
    is found. Each line is validated straight (via _extract_straight_lines), must
    reach near today's x (so its value there is a real read, not a long
    extrapolation), land inside the pane, and sit within a plausible ratio of the
    current price. Direction (support vs resistance) is decided by the caller from
    which side of the price it falls on."""
    return [r['price'] for r in read_yellow_trendlines_geom(image_path, known_price)[0]]


def read_yellow_trendlines_geom(image_path, known_price=None):
    """As read_yellow_trendlines, but keeps each line's fitted geometry — returns
    (records, meta). Wedge detection needs the SLOPES to find where two lines meet,
    which the price-only read throws away. process_one uses this and derives the
    prices from it, so the wedge costs no extra OCR pass (fit_price_axis is the
    expensive part and already runs twice per chart, once per colour)."""
    return _read_lines_at_today(image_path, trend_yellow_mask, known_price)


def _read_lines_at_today(image_path, mask_fn, known_price=None, require_reach=True,
                         min_span_frac=0.12):
    """Shared by the yellow-trendline and blue-rail reads so both price the same
    axis identically: extract straight lines from mask_fn's colour, then value each
    at today's x under the same validity gates. require_reach=False lets a line be
    extrapolated to today from wherever it was drawn, and min_span_frac sets how much
    of the pane a line must cross to count (both differ for rails — see
    read_blue_rails).

    Returns (records, meta): a record per accepted line, carrying both its price at
    today and the pixel-space fit it came from, and meta with today_x / pane width so
    a caller can reason about the lines geometrically. Callers wanting only prices
    take r['price']; meta is {} when the axis or today's x could not be read."""
    img = Image.open(image_path).convert('RGB')
    arr = np.array(img)
    h, w, _ = arr.shape
    a, b, axis_reason = fit_price_axis(img, arr, w, h, known_price)
    if axis_reason is not None:
        return [], {}
    today_x = find_today_x(arr, w)
    if today_x is None:
        return [], {}
    out = []
    for slope, intercept, xmin, xmax in _extract_straight_lines(mask_fn(arr), w, h,
                                                                min_span_frac=min_span_frac):
        # Only trust the value at today's x if the drawn line actually reaches there
        # (allow a modest 0.15w extrapolation past its drawn extent — trend lines are
        # often drawn a little short of, or into, the future edge).
        if require_reach and (today_x < xmin - 0.15 * w or today_x > xmax + 0.15 * w):
            continue
        y_today = slope * today_x + intercept
        if not (-0.05 * h <= y_today <= 1.05 * h):
            continue
        price = a * y_today + b
        if price <= 0:
            continue
        if known_price is not None and not (0.3 <= price / known_price <= 3.0):
            continue
        out.append({'price': round(price, 2), 'slope': slope, 'intercept': intercept,
                    'xmin': xmin, 'xmax': xmax})
    return out, {'today_x': today_x, 'w': w}


def detect_wedge(records, meta, known_price):
    """A WEDGE: two yellow trend lines converging in the near future (user rule
    2026-07-15). Returns a dict describing it, or None.

    A wedge does NOT change Alert Low / Alert High — the trend-line rules still stand
    and the nearest line each side of price wins as always. It is a classification of
    the chart's shape, plus the break signal the user reads off it: price breaking
    ABOVE a wedge is a potential buy.

    Both gates matter, because 'two lines that meet somewhere to the right' is not a
    pattern — any two non-parallel lines do that, and the bare test fired on 31 of the
    353 charts:
      * HORIZON. The apex must be at most WEDGE_HORIZON_FRAC of the pane width ahead
        of today — 'near future', measured on the charts at roughly three months. The
        horizon is the whole rule: relaxing it to 0.40w doubles the hits to 14.
      * GAP FLOOR. The lines must still be WEDGE_MIN_GAP_FRAC apart at today. Without
        it, one thick line double-detected reads as a wedge that has already closed —
        PALLADIUM's three 'lines' within 2% of each other, apex 0.001w away, plus
        PLATINUM/LAND/ADM.
    Where several pairs qualify the most imminent apex wins."""
    lines = [r for r in records if r.get('price')]
    if len(lines) < 2 or not meta or known_price is None:
        return None
    today_x, w = meta['today_x'], meta['w']
    best = None
    for i in range(len(lines)):
        for j in range(i + 1, len(lines)):
            l1, l2 = lines[i], lines[j]
            if abs(l1['slope'] - l2['slope']) < 1e-9:
                continue                      # parallel: they never meet
            apex_x = (l2['intercept'] - l1['intercept']) / (l1['slope'] - l2['slope'])
            ahead = (apex_x - today_x) / w
            if not (0 < ahead <= WEDGE_HORIZON_FRAC):
                continue                      # already crossed, or too far out to be 'near'
            gap = abs(l1['price'] - l2['price'])
            if gap / known_price < WEDGE_MIN_GAP_FRAC:
                continue                      # same line found twice, not a wedge
            if best is None or ahead < best['apex_ahead_frac']:
                lo, hi = sorted([l1['price'], l2['price']])
                best = {'apex_ahead_frac': round(ahead, 4), 'lower_line': lo, 'upper_line': hi,
                        'broken_above': known_price > hi}
    return best


def process_one(ticker, screenshot_path, known_price=None):
    if is_macro_reference(ticker):
        return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': None, 'lower': None, 'upper': None,
                 'x_frac': None,
                 'reason': 'macro/reference instrument (index/yield/FX) — not a buy-list channel, skipped'}
    if not screenshot_path or not os.path.exists(screenshot_path):
        return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': None, 'lower': None, 'upper': None,
                 'x_frac': None, 'reason': 'screenshot file not found'}

    raw = read_channel(screenshot_path, known_price)
    x_frac = raw['x_frac']

    # Gather candidate boundary lines from BOTH sources, then pick the line closest
    # to today's price on each side (user rule 2026-07-14): Alert Low = nearest line
    # below price, Alert High = nearest line above price, choosing among the blue
    # parallel-channel boundaries AND the user's hand-drawn yellow trend lines. A
    # yellow line further from price than the blue boundary is naturally out-competed
    # (a yellow below the channel's lower rail loses to that rail; a yellow between
    # the rail and price wins). below/above hold prices strictly on that side.
    candidates = []                # every boundary / trend line price found
    blue_reason = raw['reason']

    # Blue rails are read line-by-line (read_blue_rails) rather than taken from
    # read_channel's column-pairing, which silently dropped in-frame rails whenever
    # the other rail had left the frame or a dashed/SMA line shared the column —
    # see read_blue_rails' docstring for the SILVER case this fixes. read_channel is
    # still consulted for x_frac and for its rejection reason when nothing is found.
    blue = read_blue_rails(screenshot_path, known_price) if known_price is not None else []
    # Read the yellow lines WITH their geometry and derive the prices from it, rather
    # than calling read_yellow_trendlines and then re-reading for the wedge: the axis
    # OCR inside each read is what makes a batch run take minutes.
    if known_price is not None:
        yellow_geom, yellow_meta = read_yellow_trendlines_geom(screenshot_path, known_price)
    else:
        yellow_geom, yellow_meta = [], {}
    yellow = [r['price'] for r in yellow_geom]
    wedge = detect_wedge(yellow_geom, yellow_meta, known_price)
    candidates += blue
    candidates += yellow

    # Split EVERY candidate by which side of today's price it sits on (not by its
    # blue lower/upper role — when price has broken just outside the channel, a rail
    # can be on the far side of price, and forcing it to its old role produced a
    # degenerate alert_high < alert_low). Nearest on each side wins.
    #
    # A line closer than ON_ALERT_TOL to price is NOT assigned a side at all (user
    # rule 2026-07-15). At that distance the screenshot genuinely cannot resolve
    # which side it falls on — SILVER's top rail reads 58.09 against a 58.0602 close,
    # a 2.6px gap, so a strict comparison made the rail "resistance" when the user
    # reads it as support that price has broken above. Rather than guess the side
    # from noise, the line is reported as the level price has REACHED: it becomes
    # alert_low (the buy trigger) and the row is flagged on_alert for the Stocks of
    # Interest block. This is not an axis error — the axis was verified against the
    # dotted last-price line to within 0.01% on this chart; price is simply sitting
    # on the line.
    at = []
    below = []
    above = []
    for c in candidates:
        if known_price is None:
            continue
        if abs(c - known_price) / known_price <= ON_ALERT_TOL:
            at.append(c)
        elif c < known_price:
            below.append(c)
        else:
            above.append(c)
    on_alert = bool(at)
    # The reached line is the operative support; fall back to the nearest line
    # genuinely below price when nothing is sitting on it.
    alert_low = max(at + below) if (at or below) else None
    alert_high = min(above) if above else None     # nearest resistance above price

    # BELOW A PARALLEL CHANNEL, THE BAND STILL GOVERNS (user rule 2026-07-15).
    # The side-of-price split leaves a broken-down chart with no support at all — both
    # rails sit above price, so alert_low comes back None and the row ships a bare
    # resistance while a stale alert_low stays in the sheet above it. The user reads a
    # parallel channel as the trading band itself: price dropping out of the bottom is
    # the buy signal, not a reason to re-cast the bottom rail as resistance. So the
    # rails keep the roles they were drawn with.
    #
    # Deliberately narrow:
    #   - parallel ONLY. Where yellow lines are drawn too they were put nearer to price
    #     on purpose, and still win on the nearest-line rule (CTEC, ADM).
    #   - a real channel (two distinct rails), not a lone blue line (AV., MTRO, MKS).
    #   - a breakout ABOVE is untouched: the top rail still flips to support.
    below_parallel = (not on_alert and not yellow and known_price is not None
                      and len(set(blue)) >= 2 and all(c > known_price for c in blue))
    if below_parallel:
        alert_low, alert_high = min(blue), max(blue)

    def src(val):
        if val is None:
            return None
        return 'blue' if val in blue else 'yellow'

    # Plain-English shape of the chart, for the user to check a run against the
    # actual TradingView drawings. Describes where price sits relative to the blue
    # channel — which is what decides whether a rail acts as support or resistance.
    if on_alert:
        pattern = 'price ON a drawn line (%s) — reached, alert triggered' % (
            'blue rail' if at[0] in blue else 'trend line')
    elif wedge:
        # Named ahead of the blue classification: a wedge is the shape the user reads
        # off the chart, and it is the yellow lines that make it (CNA and DGE sit in a
        # blue channel AND carry a wedge). The alert levels are unaffected either way.
        pattern = ('wedge (trend lines converging) — price broke ABOVE, potential buy'
                   if wedge['broken_above'] else
                   'wedge (trend lines converging) — price still inside')
    elif not blue:
        pattern = 'trend lines only (no blue channel read)' if yellow else 'no lines read'
    elif known_price is None:
        pattern = 'blue channel (no price to compare)'
    elif all(c < known_price for c in blue):
        pattern = 'price ABOVE channel (broken out) — top rail is support'
    elif all(c > known_price for c in blue):
        pattern = ('price BELOW channel (broken down) — band still governs' if below_parallel
                   else 'price BELOW channel (broken down) — bottom rail is resistance')
    else:
        pattern = 'price INSIDE channel'

    if alert_low is not None and alert_high is not None:
        kind, lower, upper, reason = 'parallel', alert_low, alert_high, None
    elif alert_low is not None:
        kind, lower, upper, reason = 'single_low', alert_low, None, None
    elif alert_high is not None:
        kind, lower, upper, reason = 'single_high', None, alert_high, None
    else:
        kind, lower, upper = None, None, None
        reason = blue_reason or 'no channel or trend line found near price'

    return {'ticker': ticker, 'screenshot': screenshot_path, 'kind': kind, 'lower': lower, 'upper': upper,
             'x_frac': x_frac, 'reason': reason, 'pattern': pattern,
             'blue_lines': blue, 'yellow_lines': yellow, 'wedge': wedge,
             'on_alert': on_alert, 'on_alert_line': (max(at) if at else None),
             'alert_low_src': src(alert_low), 'alert_high_src': src(alert_high)}


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
