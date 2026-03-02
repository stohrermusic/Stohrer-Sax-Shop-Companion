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
    try_nest_partial, compute_remaining_pads,
    generate_die_svg, generate_die_svg_from_placed,
    generate_holder_svg, generate_kerf_test_svg,
    generate_organizer_svg, _layout_organizer,
    HOLDER_OUTER_R, HOLDER_MAGNET_HOLE_R, HOLDER_PIN_HOLE_R,
    HOLDER_LARGE_INNER_R, HOLDER_SMALL_INNER_R,
)
from gcode_engine import (
    generate_die_gcode_from_placed, generate_holder_gcode, can_generate_gcode,
    generate_kerf_test_gcode, generate_organizer_gcode,
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
    # Large holder
    svg_path = os.path.join(tmpdir, "holder_large.svg")
    generate_holder_svg("large", svg_path, settings)
    check("Holder SVG (large) created", os.path.exists(svg_path))
    with open(svg_path, 'r') as f:
        content = f.read()
    # 4 pieces: solid, magnet, pin, ring. Each has outer circle = 4 outer circles.
    # magnet has 1 hole, pin has 2 holes, ring has 1 inner circle = 4 additional circles
    circle_count = content.count('<circle')
    check(f"Holder large: >= 8 circles (found {circle_count})", circle_count >= 8)

    # Small holder
    svg_path = os.path.join(tmpdir, "holder_small.svg")
    generate_holder_svg("small", svg_path, settings)
    check("Holder SVG (small) created", os.path.exists(svg_path))

    # Both holders
    svg_path = os.path.join(tmpdir, "holder_both.svg")
    generate_holder_svg("both", svg_path, settings)
    check("Holder SVG (both) created", os.path.exists(svg_path))
    with open(svg_path, 'r') as f:
        content = f.read()
    # "Both" = 3 shared pieces + 2 retaining rings = 5 pieces
    # 5 outer circles + magnet hole + 2 pin holes + 2 inner rings = 10 circles
    circle_count = content.count('<circle')
    check(f"Holder both: >= 10 circles (found {circle_count})", circle_count >= 10)

# ============================================================
print("\n=== Die Holder G-code Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    gcode_path = os.path.join(tmpdir, "holder_large.gcode")
    generate_holder_gcode("large", gcode_path, settings)
    check("Holder G-code (large) created", os.path.exists(gcode_path))

    gcode_path = os.path.join(tmpdir, "holder_both.gcode")
    generate_holder_gcode("both", gcode_path, settings)
    check("Holder G-code (both) created", os.path.exists(gcode_path))
    with open(gcode_path, 'r') as f:
        content = f.read()
    check("Holder G-code has header", "G90" in content)
    check("Holder G-code has moves", "G1" in content)

# ============================================================
print("\n=== Holder Constants ===")
# ============================================================

check("Holder outer radius is 42.5mm", HOLDER_OUTER_R == 42.5)
check("Holder magnet hole radius is 3.25mm (6.5mm dia)", HOLDER_MAGNET_HOLE_R == 3.25)
check("Holder pin hole radius is 1.75mm (3.5mm dia)", HOLDER_PIN_HOLE_R == 1.75)
check("Large holder inner radius is 35.0mm", HOLDER_LARGE_INNER_R == 35.0)
check("Small holder inner radius is 25.0mm", HOLDER_SMALL_INNER_R == 25.0)

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
print("\n=== Die Organizer Layout ===")
# ============================================================

# Full set on 400x400mm
full_small = [s / 2 for s in range(14, 80)]  # 7.0 to 39.5
full_large = [s / 2 for s in range(80, 121)]  # 40.0 to 60.0
full_set = sorted(full_small + full_large)
check("Full set is 107 sizes", len(full_set) == 107)

sheets = _layout_organizer(full_set, True, 400, 400, 5.0)
check("Full set + holders fits on 400x400 at 5mm spacing",
      len(sheets) == 1)
if sheets:
    rows, width, height = sheets[0]
    total_slots = sum(len(r[3]) for r in rows)
    check(f"Full set: {total_slots} slots (107 dies + 2 holders = 109)",
          total_slots == 109)
    check(f"Full set sheet: {width:.0f}x{height:.0f} fits 400x400",
          width <= 400 and height <= 400)

# Small custom set
small_set = [10.0, 15.0, 20.0, 25.0]
sheets = _layout_organizer(small_set, False, 300, 300, 5.0)
check("4-die set fits on one sheet", len(sheets) == 1)
if sheets:
    rows, width, height = sheets[0]
    total_slots = sum(len(r[3]) for r in rows)
    check("4-die set: 4 slots", total_slots == 4)

# Force multi-sheet by using tiny sheet
sheets = _layout_organizer(full_set, False, 200, 100, 5.0)
check("Full set on 200x100: needs multiple sheets", len(sheets) > 1)

# ============================================================
print("\n=== Die Organizer SVG Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    base = os.path.join(tmpdir, "organizer")
    sizes = [10.0, 15.0, 20.0, 40.0, 45.0]
    generated = generate_organizer_svg(sizes, True, 400, 400, 5.0, base, settings, "slotted")
    check("Organizer SVG slotted created", len(generated) > 0)
    fname, w, h = generated[0]
    check("Organizer SVG file exists", os.path.exists(fname))
    with open(fname, 'r') as f:
        content = f.read()
    rect_count = content.count('<rect')
    # 7 slots (5 dies + 2 holders) + 1 outer rect = 8 rects
    check(f"Organizer SVG has rects (found {rect_count})", rect_count >= 7)
    text_count = content.count('<text')
    check(f"Organizer SVG has labels (found {text_count})", text_count >= 5)

    # Base layer
    base_gen = generate_organizer_svg(sizes, True, 400, 400, 5.0, base, settings, "base")
    check("Organizer SVG base created", len(base_gen) > 0)
    fname_b, _, _ = base_gen[0]
    with open(fname_b, 'r') as f:
        content_b = f.read()
    rect_count_b = content_b.count('<rect')
    check(f"Base layer: just 1 rect (found {rect_count_b})", rect_count_b == 1)

# ============================================================
print("\n=== Die Organizer G-code Generation ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    base = os.path.join(tmpdir, "organizer")
    sizes = [10.0, 20.0, 40.0]
    generated = generate_organizer_gcode(sizes, False, 400, 400, 5.0, base, settings, "slotted")
    check("Organizer G-code created", len(generated) > 0)
    fname, _, _ = generated[0]
    check("Organizer G-code file exists", os.path.exists(fname))
    with open(fname, 'r') as f:
        content = f.read()
    check("Organizer G-code has header", "G90" in content)
    check("Organizer G-code has moves", "G1" in content)

# ============================================================
print(f"\n{'='*50}")
print(f"RESULTS: {passed} passed, {failed} failed out of {passed + failed} tests")
if failed == 0:
    print("ALL TESTS PASSED!")
else:
    print(f"*** {failed} TESTS FAILED ***")
print(f"{'='*50}\n")
