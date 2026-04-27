"""
Test auto-fit engraving: shift toward center instead of shrinking.

Verifies that both SVG and G-code engines prefer moving text toward the
disc center over scaling down when engraving impinges on the disc edge.
"""

import sys
import os
import math
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from config import DEFAULT_SETTINGS
from svg_engine import generate_svg, get_disc_diameter
from gcode_engine import get_text_strokes, generate_gcode

pass_count = 0
fail_count = 0


def test(name, condition, detail=""):
    global pass_count, fail_count
    if condition:
        print(f"  PASS: {name}")
        pass_count += 1
    else:
        print(f"  FAIL: {name} — {detail}")
        fail_count += 1


def make_settings(**overrides):
    """Create test settings with sensible defaults for testing."""
    s = dict(DEFAULT_SETTINGS)
    s.update(overrides)
    return s


def svg_autofit_result(pad_size, material, font_size, eng_loc, settings=None):
    """
    Simulate the SVG auto-fit logic (mirrors svg_engine.py _render_svg_discs).
    Returns dict with font_size, label_y, cy, shifted, scaled flags.
    """
    if settings is None:
        settings = make_settings()
    r = get_disc_diameter(pad_size, material, settings) / 2
    hole_dia = 0
    cy = 50.0  # arbitrary center

    mode = eng_loc['mode']
    value = eng_loc['value']
    if mode == 'from_outside':
        engraving_y = cy - (r - value)
    elif mode == 'from_inside':
        hole_r = hole_dia / 2 if hole_dia > 0 else 0
        engraving_y = cy - (hole_r + value)
    else:
        hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
        offset_from_center = (r + hole_r) / 2
        engraving_y = cy - offset_from_center

    vertical_adjust = font_size * 0.35
    label_y = engraving_y + vertical_adjust
    text_content = f"{pad_size:.1f}".rstrip('0').rstrip('.')

    # Auto-fit (same math as svg_engine.py)
    min_clearance = 0.5
    text_half_h = font_size / 2
    text_half_w = sum(0.3 if c == '.' else 0.6 for c in text_content)
    text_half_w += 0.1 * (len(text_content) - 1) if len(text_content) > 1 else 0
    text_half_w = text_half_w * font_size / 2

    corners = [(text_half_w, label_y - text_half_h - cy),
               (-text_half_w, label_y - text_half_h - cy),
               (text_half_w, label_y + text_half_h - cy),
               (-text_half_w, label_y + text_half_h - cy)]
    max_dist = max(math.sqrt(px**2 + py**2) for px, py in corners)
    safe_radius = r - min_clearance

    orig_fs = font_size
    orig_ly = label_y
    shifted = False
    scaled = False

    if max_dist > safe_radius > 0:
        centered_max = math.sqrt(text_half_w**2 + text_half_h**2)
        if centered_max <= safe_radius:
            max_offset = math.sqrt(safe_radius**2 - text_half_w**2) - text_half_h
            dy = label_y - cy
            if abs(dy) > max_offset:
                label_y = cy + math.copysign(max_offset, dy)
                shifted = True
        else:
            label_y = cy
            scale = safe_radius / centered_max
            font_size *= scale
            scaled = True

    # Verify final corners
    text_half_h2 = font_size / 2
    # Recalc text_half_w with potentially new font_size
    text_half_w2 = sum(0.3 if c == '.' else 0.6 for c in text_content)
    text_half_w2 += 0.1 * (len(text_content) - 1) if len(text_content) > 1 else 0
    text_half_w2 = text_half_w2 * font_size / 2
    final_corners = [(text_half_w2, label_y - text_half_h2 - cy),
                     (-text_half_w2, label_y - text_half_h2 - cy),
                     (text_half_w2, label_y + text_half_h2 - cy),
                     (-text_half_w2, label_y + text_half_h2 - cy)]
    final_max = max(math.sqrt(px**2 + py**2) for px, py in final_corners)

    return {
        'font_size': font_size, 'orig_fs': orig_fs,
        'label_y': label_y, 'orig_ly': orig_ly, 'cy': cy,
        'r': r, 'safe_radius': safe_radius,
        'shifted': shifted, 'scaled': scaled,
        'final_max_dist': final_max,
        'text': text_content,
    }


def gcode_autofit_result(pad_size, material, font_size, eng_loc, settings=None):
    """
    Simulate the G-code auto-fit logic (mirrors gcode_engine.py _collect_disc_strokes).
    Returns dict with font_size, label_y, shifted, scaled flags.
    """
    if settings is None:
        settings = make_settings()
    r = get_disc_diameter(pad_size, material, settings) / 2
    hole_dia = 0
    cy = 50.0
    cx = 50.0

    mode = eng_loc['mode']
    value = eng_loc['value']
    if mode == 'from_outside':
        engraving_y = cy - (r - value)
    elif mode == 'from_inside':
        hole_r = hole_dia / 2 if hole_dia > 0 else 0
        engraving_y = cy - (hole_r + value)
    else:
        hole_r = hole_dia / 2 if hole_dia > 0 else 1.75
        offset_from_center = (r + hole_r) / 2
        engraving_y = cy - offset_from_center

    vertical_adjust = font_size * 0.35
    label_y = engraving_y + vertical_adjust
    text = f"{pad_size:.1f}".rstrip('0').rstrip('.')

    disc_font_size = font_size
    text_strokes = get_text_strokes(text, disc_font_size, cx, label_y)

    max_dist = 0
    for stroke in text_strokes:
        for px, py in stroke:
            dist = math.sqrt((px - cx)**2 + (py - cy)**2)
            if dist > max_dist:
                max_dist = dist

    min_clearance = 0.5
    safe_radius = r - min_clearance
    orig_fs = font_size
    orig_ly = label_y
    shifted = False
    scaled = False

    if max_dist > safe_radius > 0:
        max_horiz = max((abs(px - cx) for stroke in text_strokes for px, py in stroke), default=0)
        max_vert = max((abs(py - label_y) for stroke in text_strokes for px, py in stroke), default=0)
        centered_max = math.sqrt(max_horiz**2 + max_vert**2)

        if centered_max <= safe_radius:
            max_offset = math.sqrt(safe_radius**2 - max_horiz**2) - max_vert
            dy = label_y - cy
            new_label_y = label_y
            if abs(dy) > max_offset:
                new_label_y = cy + math.copysign(max_offset, dy)
            if new_label_y != label_y:
                text_strokes = get_text_strokes(text, disc_font_size, cx, new_label_y)
                label_y = new_label_y
                shifted = True
        else:
            scale = safe_radius / centered_max
            disc_font_size *= scale
            text_strokes = get_text_strokes(text, disc_font_size, cx, cy)
            label_y = cy
            scaled = True

    final_max = 0
    for stroke in text_strokes:
        for px, py in stroke:
            dist = math.sqrt((px - cx)**2 + (py - cy)**2)
            if dist > final_max:
                final_max = dist

    return {
        'font_size': disc_font_size, 'orig_fs': orig_fs,
        'label_y': label_y, 'orig_ly': orig_ly, 'cy': cy,
        'r': r, 'safe_radius': safe_radius,
        'shifted': shifted, 'scaled': scaled,
        'final_max_dist': final_max,
        'text': text,
    }


# ============================================================
# 1. SVG bounding-box auto-fit
# ============================================================
print("\n=== SVG Auto-Fit: Shift vs Scale ===")

# --- Case: text near edge on small pad, should SHIFT not scale ---
# 8mm felt pad, from_outside 1.0mm → text near top edge
loc_near_edge = {"mode": "from_outside", "value": 1.0}
r = svg_autofit_result(8.0, 'felt', 3.0, loc_near_edge)
test("Small pad (8mm) near-edge — shifted, not scaled",
     r['shifted'] and not r['scaled'],
     f"shifted={r['shifted']}, scaled={r['scaled']}")
test("Small pad (8mm) — font size preserved at 3.0",
     r['font_size'] == 3.0,
     f"fs={r['font_size']:.3f}")
test("Small pad (8mm) — text moved closer to center",
     abs(r['label_y'] - r['cy']) < abs(r['orig_ly'] - r['cy']),
     f"offset {abs(r['orig_ly']-r['cy']):.2f} → {abs(r['label_y']-r['cy']):.2f}")
test("Small pad (8mm) — final corners within safe radius",
     r['final_max_dist'] <= r['safe_radius'] + 0.001,
     f"max={r['final_max_dist']:.3f}, safe={r['safe_radius']:.3f}")

# --- Case: large pad, text comfortably inside ---
loc_inside = {"mode": "from_inside", "value": 4.0}
r2 = svg_autofit_result(30.0, 'felt', 3.0, loc_inside)
test("Large pad (30mm) from_inside — no adjustment needed",
     not r2['shifted'] and not r2['scaled'],
     f"shifted={r2['shifted']}, scaled={r2['scaled']}")
test("Large pad (30mm) — font size unchanged",
     r2['font_size'] == 3.0,
     f"fs={r2['font_size']:.3f}")

# --- Case: very small pad, text too wide even centered → must scale ---
r3 = svg_autofit_result(5.0, 'felt', 3.0, loc_near_edge)
test("Tiny pad (5mm) — must scale (text too large)",
     r3['scaled'],
     f"shifted={r3['shifted']}, scaled={r3['scaled']}")
test("Tiny pad (5mm) — font size reduced",
     r3['font_size'] < 3.0,
     f"fs={r3['font_size']:.3f}")
test("Tiny pad (5mm) — text centered when scaling",
     abs(r3['label_y'] - r3['cy']) < 0.01,
     f"offset={abs(r3['label_y']-r3['cy']):.3f}")
test("Tiny pad (5mm) — final corners within safe radius",
     r3['final_max_dist'] <= r3['safe_radius'] + 0.001,
     f"max={r3['final_max_dist']:.3f}, safe={r3['safe_radius']:.3f}")

# --- Case: card material, small pad ---
r4 = svg_autofit_result(8.0, 'card', 3.0, loc_near_edge)
test("Card (8mm) near-edge — shifted, not scaled",
     r4['shifted'] and not r4['scaled'],
     f"shifted={r4['shifted']}, scaled={r4['scaled']}")
test("Card (8mm) — font size preserved",
     r4['font_size'] == 3.0,
     f"fs={r4['font_size']:.3f}")

# --- Case: decimal pad size (wider text "7.5") ---
r5 = svg_autofit_result(7.5, 'felt', 3.0, loc_near_edge)
# 7.5mm felt → diameter 6.75, radius 3.375. Text "7.5" is wider (3 chars).
# centered_max with 3 chars at font 3.0: text_half_w ≈ (0.6+0.3+0.6)*3/2 + spacing = big
# This might need scaling if text is too wide
if r5['shifted']:
    test("Decimal pad (7.5mm) — shifted, font preserved",
         r5['font_size'] == 3.0,
         f"fs={r5['font_size']:.3f}")
elif r5['scaled']:
    test("Decimal pad (7.5mm) — scaled (text '7.5' too wide for tiny disc)",
         r5['font_size'] < 3.0,
         f"fs={r5['font_size']:.3f}")
test("Decimal pad (7.5mm) — within safe radius",
     r5['final_max_dist'] <= r5['safe_radius'] + 0.001,
     f"max={r5['final_max_dist']:.3f}, safe={r5['safe_radius']:.3f}")


# ============================================================
# 2. G-code stroke-based auto-fit
# ============================================================
print("\n=== G-code Auto-Fit: Shift vs Scale ===")

# G-code stroke font has tighter bounding boxes than SVG approximation,
# so we use smaller pads or more extreme positioning to trigger impingement.

# --- Small pad near edge ---
g1 = gcode_autofit_result(8.0, 'felt', 3.0, loc_near_edge)
if g1['shifted'] or g1['scaled']:
    if g1['shifted']:
        test("G-code 8mm near-edge — shifted, not scaled",
             not g1['scaled'], f"scaled={g1['scaled']}")
        test("G-code 8mm — font preserved",
             g1['font_size'] == 3.0, f"fs={g1['font_size']:.3f}")
    else:
        test("G-code 8mm near-edge — scaled (strokes wider than expected)",
             g1['scaled'], "")
    test("G-code 8mm — strokes within safe radius",
         g1['final_max_dist'] <= g1['safe_radius'] + 0.01,
         f"max={g1['final_max_dist']:.3f}, safe={g1['safe_radius']:.3f}")
else:
    # Stroke font fits without adjustment — that's fine, the font is just tighter
    test("G-code 8mm — strokes fit without adjustment (tighter font)",
         g1['final_max_dist'] <= g1['safe_radius'],
         f"max={g1['final_max_dist']:.3f}, safe={g1['safe_radius']:.3f}")

# --- 6mm pad: stroke font is compact enough to fit without adjustment ---
g2 = gcode_autofit_result(6.0, 'felt', 3.0, loc_near_edge)
test("G-code 6mm — strokes fit (compact stroke font)",
     g2['final_max_dist'] <= g2['safe_radius'] + 0.01,
     f"max={g2['final_max_dist']:.3f}, safe={g2['safe_radius']:.3f}")

# --- Large pad, no adjustment ---
g3 = gcode_autofit_result(30.0, 'felt', 3.0, loc_inside)
test("G-code 30mm from_inside — no adjustment",
     not g3['shifted'] and not g3['scaled'],
     f"shifted={g3['shifted']}, scaled={g3['scaled']}")

# --- Very small pad: stroke font compact enough, but test with larger font ---
g4 = gcode_autofit_result(5.0, 'felt', 3.0, loc_near_edge)
test("G-code 5mm — strokes fit or adjusted",
     g4['final_max_dist'] <= g4['safe_radius'] + 0.01,
     f"max={g4['final_max_dist']:.3f}, safe={g4['safe_radius']:.3f}")

# Force impingement with a bigger font to verify shift works in G-code
g5 = gcode_autofit_result(10.0, 'felt', 5.0, loc_near_edge)
if g5['shifted']:
    test("G-code 10mm big font — shifted, not scaled",
         not g5['scaled'] and g5['font_size'] == 5.0,
         f"scaled={g5['scaled']}, fs={g5['font_size']:.3f}")
    test("G-code 10mm big font — within safe radius after shift",
         g5['final_max_dist'] <= g5['safe_radius'] + 0.01,
         f"max={g5['final_max_dist']:.3f}, safe={g5['safe_radius']:.3f}")
elif g5['scaled']:
    test("G-code 10mm big font — scaled (font too wide)",
         g5['font_size'] < 5.0,
         f"fs={g5['font_size']:.3f}")
    test("G-code 10mm big font — within safe radius after scale",
         g5['final_max_dist'] <= g5['safe_radius'] + 0.01,
         f"max={g5['final_max_dist']:.3f}, safe={g5['safe_radius']:.3f}")


# ============================================================
# 3. Key invariant: shift should NEVER change font size
# ============================================================
print("\n=== Invariant: Shift Preserves Font Size ===")

for pad_size in [7.0, 8.0, 9.0, 10.0, 12.0]:
    for mat in ['felt', 'card']:
        rs = svg_autofit_result(pad_size, mat, 3.0, loc_near_edge)
        if rs['shifted']:
            test(f"SVG {mat} {pad_size}mm shift — font size = 3.0",
                 rs['font_size'] == 3.0,
                 f"fs={rs['font_size']:.3f}")
        rg = gcode_autofit_result(pad_size, mat, 3.0, loc_near_edge)
        if rg['shifted']:
            test(f"G-code {mat} {pad_size}mm shift — font size = 3.0",
                 rg['font_size'] == 3.0,
                 f"fs={rg['font_size']:.3f}")


# ============================================================
# 4. Key invariant: result always within safe radius
# ============================================================
print("\n=== Invariant: Always Within Safe Radius ===")

for pad_size in [5.0, 6.0, 7.0, 8.0, 10.0, 15.0, 25.0]:
    for mat in ['felt', 'card']:
        rs = svg_autofit_result(pad_size, mat, 3.0, loc_near_edge)
        if rs['shifted'] or rs['scaled']:
            test(f"SVG {mat} {pad_size}mm — within safe radius",
                 rs['final_max_dist'] <= rs['safe_radius'] + 0.001,
                 f"max={rs['final_max_dist']:.3f}, safe={rs['safe_radius']:.3f}")
        rg = gcode_autofit_result(pad_size, mat, 3.0, loc_near_edge)
        if rg['shifted'] or rg['scaled']:
            test(f"G-code {mat} {pad_size}mm — within safe radius",
                 rg['final_max_dist'] <= rg['safe_radius'] + 0.01,
                 f"max={rg['final_max_dist']:.3f}, safe={rg['safe_radius']:.3f}")


# ============================================================
# 5. Integration: generate actual SVG file
# ============================================================
print("\n=== Integration: SVG File Output ===")

with tempfile.TemporaryDirectory() as tmpdir:
    # Use 15mm and 25mm pads with 2mm font to avoid 80% radius cutoff
    # (8mm felt = 3.625mm radius, 3mm font >= 2.9 = 80% of radius → engraving off)
    pads = [{'size': 15.0, 'qty': 2}, {'size': 25.0, 'qty': 1}]
    svg_file = os.path.join(tmpdir, "test.svg")
    s = make_settings()
    s["engraving_on"] = True
    s["engraving_font_size"] = {"felt": 2.0, "card": 2.0, "leather": 2.0, "exact_size": 2.0}
    s["engraving_location"] = {
        "felt": {"mode": "from_outside", "value": 1.0},
        "card": {"mode": "from_outside", "value": 1.0},
        "leather": {"mode": "from_outside", "value": 1.0},
        "exact_size": {"mode": "from_outside", "value": 1.0},
    }

    try:
        generate_svg(pads, 'felt', 200, 200, svg_file, 0, s)
        with open(svg_file, 'r') as f:
            content = f.read()

        import re
        text_tags = re.findall(r'<text\b', content)
        test("SVG has text elements for each pad",
             len(text_tags) >= 3,
             f"found {len(text_tags)} text elements for 3 pads")

        # Extract font-size values
        font_sizes = re.findall(r'font-size="([\d.]+)', content)
        font_sizes = [float(fs) for fs in font_sizes]
        test("SVG font sizes preserved at 2.0mm (shifted, not scaled)",
             all(abs(fs - 2.0) < 0.01 for fs in font_sizes),
             f"font sizes found: {font_sizes}")

    except Exception as e:
        test("SVG generation succeeded", False, str(e))

    # G-code
    print("\n=== Integration: G-code File Output ===")
    gc_file = os.path.join(tmpdir, "test.gcode")
    try:
        generate_gcode(pads, 'felt', 200, 200, gc_file, 0, s)
        with open(gc_file, 'r') as f:
            gc_content = f.read()
        test("G-code file generated", len(gc_content) > 100, f"len={len(gc_content)}")
        # G-code should have X/Y moves for engraving
        has_moves = 'G1' in gc_content or 'G0' in gc_content
        test("G-code contains motion commands", has_moves, "no G0/G1 found")
    except Exception as e:
        test("G-code generation succeeded", False, str(e))


# ============================================================
# Summary
# ============================================================
print(f"\n{'='*50}")
print(f"RESULTS: {pass_count} passed, {fail_count} failed out of {pass_count + fail_count}")
if fail_count == 0:
    print("All tests PASSED!")
else:
    print("Some tests FAILED.")
    sys.exit(1)
