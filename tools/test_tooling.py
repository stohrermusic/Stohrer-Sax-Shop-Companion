"""
Test script for Tooling tab features: die insert generation, die holder generation,
size parsing, nesting, and scrap mode logic.
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
from svg_engine import (
    get_disc_diameter, _nest_discs, can_all_pads_fit,
    try_nest_partial, generate_die_svg, generate_die_svg_from_placed,
    generate_holder_svg, generate_kerf_test_svg,
    generate_die_organizer_svg,
    HOLDER_OUTER_R, HOLDER_MAGNET_HOLE_R, HOLDER_PIN_HOLE_R,
    HOLDER_LARGE_INNER_R, HOLDER_SMALL_INNER_R,
)
from gcode_engine import (
    generate_die_gcode_from_placed, generate_holder_gcode, can_generate_gcode,
    generate_kerf_test_gcode,
)

passed = 0
failed = 0

def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1

settings = DEFAULT_SETTINGS.copy()

# ============================================================
print("\n=== Die Ring Material: get_disc_diameter ===")
# ============================================================

check("Small die 7.0mm -> 50mm dia",
      get_disc_diameter(7.0, 'die_ring', settings) == 50.0)

check("Small die 20.0mm -> 50mm dia",
      get_disc_diameter(20.0, 'die_ring', settings) == 50.0)

check("Small die 39.5mm -> 50mm dia",
      get_disc_diameter(39.5, 'die_ring', settings) == 50.0)

check("Large die 40.0mm -> 70mm dia",
      get_disc_diameter(40.0, 'die_ring', settings) == 70.0)

check("Large die 50.0mm -> 70mm dia",
      get_disc_diameter(50.0, 'die_ring', settings) == 70.0)

check("Large die 60.0mm -> 70mm dia",
      get_disc_diameter(60.0, 'die_ring', settings) == 70.0)

# Boundary: 39.5 is small, 40.0 is large
check("Boundary: 39.5 is small (50mm)",
      get_disc_diameter(39.5, 'die_ring', settings) == 50.0)
check("Boundary: 40.0 is large (70mm)",
      get_disc_diameter(40.0, 'die_ring', settings) == 70.0)

# ============================================================
print("\n=== can_generate_gcode ===")
# ============================================================

check("die_ring is supported for gcode", can_generate_gcode('die_ring'))
check("acrylic is supported for gcode", can_generate_gcode('acrylic'))
check("felt still supported", can_generate_gcode('felt'))

# ============================================================
print("\n=== Die Nesting: _nest_discs ===")
# ============================================================

# Small dies: 50mm diameter each. On a 300x300mm sheet, should fit many
small_pads = [{'size': 10.0, 'qty': 4}]
placed, fp, ft = _nest_discs(small_pads, 'die_ring', 300, 300, settings)
check("4 small dies on 300x300: all placed", fp == 4 and ft == 4)
check("4 small dies on 300x300: placed list len", len(placed) == 4)

# Verify radius is correct (50mm diameter = 25mm radius)
if placed:
    check("Small die radius is 25mm", placed[0][3] == 25.0)

# Large dies: 70mm diameter each. On 300x300mm, should fit several
large_pads = [{'size': 45.0, 'qty': 4}]
placed, fp, ft = _nest_discs(large_pads, 'die_ring', 300, 300, settings)
check("4 large dies on 300x300: all placed", fp == 4 and ft == 4)
if placed:
    check("Large die radius is 35mm", placed[0][3] == 35.0)

# Mixed sizes: some small, some large
mixed_pads = [{'size': 10.0, 'qty': 2}, {'size': 45.0, 'qty': 2}]
placed, fp, ft = _nest_discs(mixed_pads, 'die_ring', 300, 300, settings)
check("Mixed dies on 300x300: all placed", fp == 4 and ft == 4)

# Tiny sheet - nothing fits
tiny_pads = [{'size': 10.0, 'qty': 1}]
placed, fp, ft = _nest_discs(tiny_pads, 'die_ring', 30, 30, settings)
check("Die on 30x30mm sheet: none fit", fp == 0)

# Can all pads fit?
check("can_all_pads_fit: 4 small on 300x300",
      can_all_pads_fit([{'size': 10.0, 'qty': 4}], 'die_ring', 300, 300, settings))
check("can_all_pads_fit: die on tiny sheet is False",
      not can_all_pads_fit([{'size': 10.0, 'qty': 1}], 'die_ring', 30, 30, settings))

# ============================================================
print("\n=== Die Scrap Mode: try_nest_partial & compute_remaining ===")
# ============================================================

# 10 small dies on a small sheet that can only hold 4
scrap_pads = [{'size': s, 'qty': 1} for s in [10.0, 11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0]]
placed, remaining, any_placed = try_nest_partial(scrap_pads, 'die_ring', 200, 120, settings)
check("Scrap: some placed on 200x120", any_placed and len(placed) > 0)
check("Scrap: some remaining", len(remaining) > 0)
check("Scrap: placed + remaining = original",
      len(placed) + sum(p['qty'] for p in remaining) == 10)

# Second scrap with remaining
if remaining:
    placed2, remaining2, any_placed2 = try_nest_partial(remaining, 'die_ring', 200, 120, settings)
    check("Scrap round 2: some placed", any_placed2)
    total_placed = len(placed) + len(placed2)
    total_remaining = sum(p['qty'] for p in remaining2)
    check("Scrap round 2: total accounted for", total_placed + total_remaining == 10)

# ============================================================
print("\n=== Die SVG Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    # Generate die SVG
    svg_path = os.path.join(tmpdir, "test_dies.svg")
    pads = [{'size': 10.0, 'qty': 1}, {'size': 20.0, 'qty': 1}, {'size': 40.0, 'qty': 1}]
    placed = generate_die_svg(pads, 300, 300, svg_path, settings)
    check("Die SVG created", os.path.exists(svg_path))
    check("Die SVG has 3 placed", len(placed) == 3)

    # Read SVG and verify content
    with open(svg_path, 'r') as f:
        svg_content = f.read()
    # Should have circles (outer + inner for each die = 6 circles minimum)
    circle_count = svg_content.count('<circle')
    check(f"Die SVG has >= 6 circles (found {circle_count})", circle_count >= 6)
    # Should have text elements for engravings
    text_count = svg_content.count('<text')
    check(f"Die SVG has text engravings (found {text_count})", text_count > 0)

    # Generate from placed (scrap mode)
    svg_path2 = os.path.join(tmpdir, "test_dies_scrap.svg")
    generate_die_svg_from_placed(placed, 300, 300, svg_path2, settings)
    check("Die SVG from_placed created", os.path.exists(svg_path2))

# ============================================================
print("\n=== Die G-code Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    gcode_path = os.path.join(tmpdir, "test_dies.gcode")
    pads = [{'size': 10.0, 'qty': 1}, {'size': 40.0, 'qty': 1}]
    placed, _, _ = _nest_discs(pads, 'die_ring', 300, 300, settings)
    generate_die_gcode_from_placed(placed, 300, 300, gcode_path, settings)
    check("Die G-code created", os.path.exists(gcode_path))

    with open(gcode_path, 'r') as f:
        gcode_content = f.read()
    check("Die G-code has header (G90)", "G90" in gcode_content)
    check("Die G-code has moves", "G1" in gcode_content)
    # Should have both inner and outer cuts
    check("Die G-code non-trivial (>100 lines)", len(gcode_content.split('\n')) > 100)

# ============================================================
print("\n=== Die Holder SVG Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    # 6-layer Large holder (default): solid + magnet + 3x pin + ring = 6 pieces
    svg_path = os.path.join(tmpdir, "holder_large.svg")
    generate_holder_svg("large", svg_path, settings)
    check("Holder SVG (large, 6-layer) created", os.path.exists(svg_path))
    with open(svg_path, 'r') as f:
        content = f.read()
    # 6 outer + 1 magnet hole + 3 pin holes + 1 ring inner = 11 circles
    circle_count = content.count('<circle')
    check(f"Holder large 6-layer: 11 circles (found {circle_count})", circle_count == 11)

    # 5-layer Large: solid + magnet + 2x pin + ring = 5 pieces
    svg_path = os.path.join(tmpdir, "holder_large5.svg")
    generate_holder_svg("large", svg_path, settings, layer_count=5)
    check("Holder SVG (large, 5-layer) created", os.path.exists(svg_path))
    with open(svg_path, 'r') as f:
        content = f.read()
    # 5 outer + 1 magnet hole + 2 pin holes + 1 ring inner = 9 circles
    circle_count = content.count('<circle')
    check(f"Holder large 5-layer: 9 circles (found {circle_count})", circle_count == 9)

    # Small holder
    svg_path = os.path.join(tmpdir, "holder_small.svg")
    generate_holder_svg("small", svg_path, settings)
    check("Holder SVG (small) created", os.path.exists(svg_path))

    # Both holders, 6-layer = 12 pieces (two complete independent holders)
    svg_path = os.path.join(tmpdir, "holder_both6.svg")
    generate_holder_svg("both", svg_path, settings, layer_count=6,
                       sheet_width_mm=400, sheet_height_mm=400)
    check("Holder SVG (both, 6-layer) created", os.path.exists(svg_path))
    with open(svg_path, 'r') as f:
        content = f.read()
    # 12 outer + 2 magnet holes + 6 pin holes + 2 ring inners = 22 circles
    circle_count = content.count('<circle')
    check(f"Holder both 6-layer: 22 circles (found {circle_count})", circle_count == 22)
    # Sheet dims show up in the SVG header
    check("Holder both 6-layer SVG width = 400mm",
          'width="400mm"' in content)

    # Both holders, 5-layer = 10 pieces
    svg_path = os.path.join(tmpdir, "holder_both5.svg")
    generate_holder_svg("both", svg_path, settings, layer_count=5,
                       sheet_width_mm=400, sheet_height_mm=400)
    with open(svg_path, 'r') as f:
        content = f.read()
    # 10 outer + 2 magnet holes + 4 pin holes + 2 ring inners = 18 circles
    circle_count = content.count('<circle')
    check(f"Holder both 5-layer: 18 circles (found {circle_count})", circle_count == 18)

    # Sheet too small -> ValueError
    err_raised = False
    try:
        generate_holder_svg("both", os.path.join(tmpdir, "fail.svg"), settings,
                           layer_count=6, sheet_width_mm=200, sheet_height_mm=200)
    except ValueError:
        err_raised = True
    check("Both 6-layer on 200x200 sheet raises ValueError", err_raised)

    # Sheet just barely too small for a single 6-layer = 6 pieces (need ~3 cols)
    err_raised = False
    try:
        generate_holder_svg("large", os.path.join(tmpdir, "fail2.svg"), settings,
                           layer_count=6, sheet_width_mm=100, sheet_height_mm=100)
    except ValueError:
        err_raised = True
    check("Large 6-layer on 100x100 sheet raises ValueError", err_raised)

    # bad layer_count rejected
    err_raised = False
    try:
        generate_holder_svg("large", os.path.join(tmpdir, "fail3.svg"), settings,
                           layer_count=4)
    except ValueError:
        err_raised = True
    check("layer_count=4 raises ValueError", err_raised)

# ============================================================
print("\n=== Die Holder G-code Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    gcode_path = os.path.join(tmpdir, "holder_large.gcode")
    generate_holder_gcode("large", gcode_path, settings)
    check("Holder G-code (large, 6-layer default) created", os.path.exists(gcode_path))

    gcode_path = os.path.join(tmpdir, "holder_both.gcode")
    generate_holder_gcode("both", gcode_path, settings, layer_count=6,
                         sheet_width_mm=400, sheet_height_mm=400)
    check("Holder G-code (both, 6-layer, 400x400) created", os.path.exists(gcode_path))
    with open(gcode_path, 'r') as f:
        content = f.read()
    check("Holder G-code has header", "G90" in content)
    check("Holder G-code has moves", "G1" in content)

    gcode_path = os.path.join(tmpdir, "holder_small5.gcode")
    generate_holder_gcode("small", gcode_path, settings, layer_count=5)
    check("Holder G-code (small, 5-layer) created", os.path.exists(gcode_path))

    # Sheet too small -> ValueError
    err_raised = False
    try:
        generate_holder_gcode("both", os.path.join(tmpdir, "g_fail.gcode"), settings,
                             layer_count=6, sheet_width_mm=200, sheet_height_mm=200)
    except ValueError:
        err_raised = True
    check("G-code Both 6-layer on 200x200 raises ValueError", err_raised)

# ============================================================
print("\n=== Holder Constants ===")
# ============================================================

check("Holder outer radius is 42.5mm", HOLDER_OUTER_R == 42.5)
check("Holder magnet hole radius is 3.25mm (6.5mm dia)", HOLDER_MAGNET_HOLE_R == 3.25)
check("Holder pin hole radius is 1.75mm (3.5mm dia)", HOLDER_PIN_HOLE_R == 1.75)
check("Large holder inner radius is 35.0mm", HOLDER_LARGE_INNER_R == 35.0)
check("Small holder inner radius is 25.0mm", HOLDER_SMALL_INNER_R == 25.0)

# ============================================================
print("\n=== _holder_pieces_for ===")
# ============================================================

from svg_engine import _holder_pieces_for, _pack_holder_grid

check("Large 6-layer = 6 pieces", len(_holder_pieces_for("large", 6)) == 6)
check("Large 5-layer = 5 pieces", len(_holder_pieces_for("large", 5)) == 5)
check("Small 6-layer = 6 pieces", len(_holder_pieces_for("small", 6)) == 6)
check("Small 5-layer = 5 pieces", len(_holder_pieces_for("small", 5)) == 5)
check("Both 6-layer = 12 pieces", len(_holder_pieces_for("both", 6)) == 12)
check("Both 5-layer = 10 pieces", len(_holder_pieces_for("both", 5)) == 10)

# Composition: 1 solid + 1 magnet + (N-3) pin + 1 ring per holder
types_l6 = [p[0] for p in _holder_pieces_for("large", 6)]
check("Large 6-layer has 1 solid", types_l6.count('solid') == 1)
check("Large 6-layer has 1 magnet", types_l6.count('magnet') == 1)
check("Large 6-layer has 3 pin", types_l6.count('pin') == 3)
check("Large 6-layer has 1 ring", types_l6.count('ring') == 1)

types_l5 = [p[0] for p in _holder_pieces_for("large", 5)]
check("Large 5-layer has 2 pin (one fewer)", types_l5.count('pin') == 2)

# Both: two complete independent holders, both ring sizes present
both6 = _holder_pieces_for("both", 6)
both_types = [p[0] for p in both6]
check("Both 6-layer has 2 solid", both_types.count('solid') == 2)
check("Both 6-layer has 2 magnet", both_types.count('magnet') == 2)
check("Both 6-layer has 6 pin", both_types.count('pin') == 6)
check("Both 6-layer has 2 ring", both_types.count('ring') == 2)
ring_radii = sorted(p[1] for p in both6 if p[0] == 'ring')
check("Both 6-layer rings = small + large",
      ring_radii == sorted([HOLDER_SMALL_INNER_R, HOLDER_LARGE_INNER_R]))

# Validation
err = False
try: _holder_pieces_for("large", 7)
except ValueError: err = True
check("layer_count=7 rejected", err)

err = False
try: _holder_pieces_for("nonsense", 6)
except ValueError: err = True
check("invalid variant rejected", err)

# ============================================================
print("\n=== _pack_holder_grid ===")
# ============================================================

# 6 pieces (single 6-layer holder) on a roomy sheet
result = _pack_holder_grid(6, 400, 400)
check("6 pieces on 400x400 fits", result is not None)
cols, rows = result
check("6 pieces on 400x400: layout >= 6 slots", cols * rows >= 6)
check("6 pieces on 400x400: near-square", abs(cols - rows) <= 1)

# 12 pieces (Both 6-layer) on 400x400 — fits as 4x3 or similar
result = _pack_holder_grid(12, 400, 400)
check("12 pieces on 400x400 fits", result is not None)

# 12 pieces on 200x200 — too small
check("12 pieces on 200x200 doesn't fit",
      _pack_holder_grid(12, 200, 200) is None)

# Boundary: 1 piece needs 85 + 2*5 = 95mm in each dim
check("1 piece on 95x95 fits exactly",
      _pack_holder_grid(1, 95, 95) is not None)
check("1 piece on 94x94 doesn't fit",
      _pack_holder_grid(1, 94, 94) is None)

# 6 pieces on huge sheet should still pick a near-square layout, not 6x1
result = _pack_holder_grid(6, 2000, 2000)
cols, rows = result
check("6 pieces on huge sheet picks near-square (not 6x1)",
      abs(cols - rows) <= 1 and max(cols, rows) <= 3)

# ============================================================
print("\n=== Phil Noy Ring Engraving (regression check) ===")
# ============================================================
# Phil Noy gave away the pad-making method for free. The "DESIGNED BY
# PHIL NOY" engraving on every retaining ring was dropped once before
# and Phil was upset. This test guards against regressions.

with tempfile.TemporaryDirectory() as tmpdir:
    eng_color = settings["layer_colors"].get('die_engraving', '#00E000')

    svg_path = os.path.join(tmpdir, "credit_large.svg")
    generate_holder_svg("large", svg_path, settings)
    with open(svg_path, 'r') as f:
        large_count = f.read().count(eng_color)
    check(f"Large holder SVG has ring engraving polylines (found {large_count})",
          large_count > 0)

    svg_path = os.path.join(tmpdir, "credit_both.svg")
    generate_holder_svg("both", svg_path, settings,
                       sheet_width_mm=400, sheet_height_mm=400)
    with open(svg_path, 'r') as f:
        both_count = f.read().count(eng_color)
    # Both has 2 rings; engraving polyline count should be roughly 2x
    check(f"Both holder engraves on every ring "
          f"(large={large_count}, both={both_count}, expect both ~2x large)",
          both_count >= 1.5 * large_count)

# ============================================================
print("\n=== Engraving Toggle ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    # With engravings off
    no_eng_settings = DEFAULT_SETTINGS.copy()
    no_eng_settings["tooling_settings"] = {
        "engrave_ring": False, "engrave_cutout": False,
        "engraving_mode": "filled", "step_size": "0.5",
        "sheet_width": "12", "sheet_height": "12",
    }
    svg_path = os.path.join(tmpdir, "dies_no_eng.svg")
    pads = [{'size': 20.0, 'qty': 1}]
    generate_die_svg(pads, 300, 300, svg_path, no_eng_settings)
    with open(svg_path, 'r') as f:
        content = f.read()
    text_count = content.count('<text')
    check(f"No engraving: 0 text elements (found {text_count})", text_count == 0)

    # With ring engraving only
    ring_eng_settings = DEFAULT_SETTINGS.copy()
    ring_eng_settings["tooling_settings"] = {
        "engrave_ring": True, "engrave_cutout": False,
        "engraving_mode": "filled", "step_size": "0.5",
        "sheet_width": "12", "sheet_height": "12",
    }
    svg_path = os.path.join(tmpdir, "dies_ring_eng.svg")
    generate_die_svg(pads, 300, 300, svg_path, ring_eng_settings)
    with open(svg_path, 'r') as f:
        content = f.read()
    text_count = content.count('<text')
    check(f"Ring engraving only: 1 text element (found {text_count})", text_count == 1)

# ============================================================
print("\n=== Size Parsing (unit test of the logic) ===")
# ============================================================

# Test the parsing logic directly (simulating what ToolingTabMixin._parse_die_sizes does)
def parse_die_sizes(text, step=0.5):
    """Standalone version of the parsing logic for testing."""
    sizes = set()
    for part in text.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            parts = part.split('-', 1)
            try:
                lo = float(parts[0].strip())
                hi = float(parts[1].strip())
            except ValueError:
                continue
            if lo > hi:
                lo, hi = hi, lo
            current = lo
            while current <= hi + 0.001:
                sizes.add(round(current, 2))
                current += step
        else:
            try:
                sizes.add(float(part))
            except ValueError:
                continue
    return sorted(sizes)

check("Parse individual sizes: '7, 8.5, 25'",
      parse_die_sizes("7, 8.5, 25") == [7.0, 8.5, 25.0])

check("Parse range: '10-12'",
      parse_die_sizes("10-12") == [10.0, 10.5, 11.0, 11.5, 12.0])

check("Parse range with step 0.25: '10-11'",
      parse_die_sizes("10-11", step=0.25) == [10.0, 10.25, 10.5, 10.75, 11.0])

check("Parse combined: '7, 10-12, 40'",
      parse_die_sizes("7, 10-12, 40") == [7.0, 10.0, 10.5, 11.0, 11.5, 12.0, 40.0])

check("Parse full small set: '7-39.5' has correct count",
      len(parse_die_sizes("7-39.5")) == 66)  # (39.5-7)/0.5 + 1 = 66

check("Parse full large set: '40-60' has correct count",
      len(parse_die_sizes("40-60")) == 41)  # (60-40)/0.5 + 1 = 41

check("Parse full set: '7-60' has correct count",
      len(parse_die_sizes("7-60")) == 107)  # 66 + 41 = 107

check("Parse empty string", parse_die_sizes("") == [])

check("Parse invalid input gracefully", parse_die_sizes("abc, 10, xyz") == [10.0])

check("Parse reversed range: '12-10'",
      parse_die_sizes("12-10") == [10.0, 10.5, 11.0, 11.5, 12.0])

# ============================================================
print("\n=== Kerf Test SVG Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    for material in ["Felt", "Card", "Leather", "Acrylic"]:
        svg_path = os.path.join(tmpdir, f"kerf_test_{material.lower()}.svg")
        generate_kerf_test_svg(material, svg_path, settings)
        check(f"Kerf test SVG ({material}) created", os.path.exists(svg_path))
        with open(svg_path, 'r') as f:
            content = f.read()
        circle_count = content.count('<circle')
        check(f"Kerf test SVG ({material}) has 3 circles (found {circle_count})", circle_count == 3)
        text_count = content.count('<text')
        check(f"Kerf test SVG ({material}) has text (found {text_count})", text_count > 0)
        check(f"Kerf test SVG ({material}) mentions material name", material in content)

# ============================================================
print("\n=== Kerf Test G-code Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    for material in ["Felt", "Acrylic"]:
        gcode_path = os.path.join(tmpdir, f"kerf_test_{material.lower()}.gcode")
        generate_kerf_test_gcode(material, gcode_path, settings,
                                  cut_speed=180, cut_power=100,
                                  eng_speed=1000, eng_power=15)
        check(f"Kerf test G-code ({material}) created", os.path.exists(gcode_path))
        with open(gcode_path, 'r') as f:
            content = f.read()
        check(f"Kerf test G-code ({material}) has header", "G90" in content)
        check(f"Kerf test G-code ({material}) has moves", "G1" in content)
        # Verify cut speed is correct (S180 for speed 180)
        check(f"Kerf test G-code ({material}) uses correct cut speed",
              "F180" in content)

# ============================================================
print("\n=== Die Organizer SVG Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    upper_path = os.path.join(tmpdir, "organizer_upper.svg")
    generate_die_organizer_svg("upper", upper_path, settings)
    check("Organizer SVG (upper) created", os.path.exists(upper_path))
    with open(upper_path, 'r') as f:
        upper_content = f.read()
    # Asset is byte-for-byte from Matt's CAD — should still parse as SVG and
    # contain the lots-of-slot rects plus the corner mounting holes.
    check("Organizer upper has SVG header", "<svg" in upper_content)
    check("Organizer upper has corner mounting holes (4 circles)",
          upper_content.count("<circle") == 4)
    check("Organizer upper has many slot rects (>100)",
          upper_content.count("<rect") > 100)

    lower_path = os.path.join(tmpdir, "organizer_lower.svg")
    generate_die_organizer_svg("lower", lower_path, settings)
    check("Organizer SVG (lower) created", os.path.exists(lower_path))
    with open(lower_path, 'r') as f:
        lower_content = f.read()
    check("Organizer lower has SVG header", "<svg" in lower_content)
    check("Organizer lower has corner mounting holes",
          lower_content.count("<circle") == 4)
    # Lower is just plate + 4 holes, so far fewer rects
    check("Organizer lower has just the plate rect (1 <rect>)",
          lower_content.count("<rect") == 1)

    # Output is byte-identical to the bundled asset (Option A: copy as-is).
    import filecmp
    upper_asset = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                               'tooling_assets', 'die_organizer_upper.svg')
    check("Organizer upper output matches bundled asset byte-for-byte",
          filecmp.cmp(upper_path, upper_asset, shallow=False))

    # Variant validation
    err_raised = False
    try:
        generate_die_organizer_svg("middle", os.path.join(tmpdir, "fail.svg"), settings)
    except ValueError:
        err_raised = True
    check("Invalid variant rejected", err_raised)

# ============================================================
print(f"\n{'='*50}")
print(f"RESULTS: {passed} passed, {failed} failed out of {passed + failed} tests")
if failed == 0:
    print("ALL TESTS PASSED!")
else:
    print(f"*** {failed} TESTS FAILED ***")
print(f"{'='*50}\n")

sys.exit(0 if failed == 0 else 1)
