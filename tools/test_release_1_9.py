#!/usr/bin/env python3
"""
Pre-release validation for Stohrer Sax Shop Companion v1.9.
Non-interactive (no GUI). Tests engine/logic functions directly.
Run: python tools/test_release_1_9.py
"""

import sys
import os
import json
import math
import tempfile

# Add parent dir so we can import project modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Prevent tkinter messagebox import from crashing headless
import unittest.mock
sys.modules['tkinter'] = unittest.mock.MagicMock()
sys.modules['tkinter.messagebox'] = unittest.mock.MagicMock()

from config import DEFAULT_SETTINGS, load_settings, SETTINGS_FILE
from svg_engine import (
    get_disc_diameter, _nest_discs, can_all_pads_fit,
    _point_in_polygon, _circle_fits_in_polygon, _nest_discs_polygon,
    _distance_point_to_segment, _distance_to_nearest_edge,
    try_nest_partial, compute_remaining_pads,
    leather_back_wrap, get_felt_thickness_mm,
)

# ---------------------------------------------------------------------------
# Test harness
# ---------------------------------------------------------------------------

results = []


def test(name, condition, detail=""):
    status = "PASS" if condition else "FAIL"
    results.append((name, condition))
    msg = f"  [{status}] {name}"
    if detail and not condition:
        msg += f"  -- {detail}"
    print(msg)


def make_settings(**overrides):
    """Return a copy of DEFAULT_SETTINGS with overrides applied."""
    s = json.loads(json.dumps(DEFAULT_SETTINGS))  # deep copy
    for k, v in overrides.items():
        if isinstance(v, dict) and isinstance(s.get(k), dict):
            s[k].update(v)
        else:
            s[k] = v
    return s


def make_pads(*specs):
    """Helper: make_pads((18.0, 3), (25.0, 2)) -> pad list."""
    pads = []
    for size, qty in specs:
        pads.append({'size': size, 'qty': qty})
    return pads


# ===========================================================================
print("=" * 60)
print("Stohrer Sax Shop Companion v1.9 — Pre-Release Validation")
print("=" * 60)

# ===========================================================================
# 1. SVG ENGINE: SIZING CALCULATIONS
# ===========================================================================
print("\n--- 1. Disc Sizing (all materials) ---")

settings = make_settings()

# Felt: pad_size - felt_offset
felt_d = get_disc_diameter(20.0, 'felt', settings)
test("Felt sizing: 20mm pad", abs(felt_d - (20.0 - 0.75)) < 0.001,
     f"got {felt_d}, expected {20.0 - 0.75}")

# Card: pad_size - (felt_offset + card_to_felt_offset)
card_d = get_disc_diameter(20.0, 'card', settings)
expected_card = 20.0 - (0.75 + 2.0)
test("Card sizing: 20mm pad", abs(card_d - expected_card) < 0.001,
     f"got {card_d}, expected {expected_card}")

# Exact: pad_size unchanged
exact_d = get_disc_diameter(25.5, 'exact_size', settings)
test("Exact sizing: 25.5mm pad", abs(exact_d - 25.5) < 0.001,
     f"got {exact_d}")

# Leather: pad_size + 2*(felt_thickness + wrap), with dart bonus for small pads
leather_big = get_disc_diameter(30.0, 'leather', settings)
test("Leather sizing: 30mm pad is larger than pad", leather_big > 30.0,
     f"got {leather_big}")

leather_small = get_disc_diameter(12.0, 'leather', settings)
test("Leather sizing: 12mm pad is larger than pad", leather_small > 12.0,
     f"got {leather_small}")

# Small leather with darts should be bigger than without (dart_wrap_bonus)
s_no_darts = make_settings(darts_enabled=False)
leather_no_darts = get_disc_diameter(12.0, 'leather', s_no_darts)
test("Leather dart bonus adds size for small pads",
     leather_small > leather_no_darts,
     f"with darts={leather_small}, without={leather_no_darts}")

# Die ring sizing
die_small = get_disc_diameter(35.0, 'die_ring', settings)
die_large = get_disc_diameter(45.0, 'die_ring', settings)
test("Die ring: small <=39.5 -> 50mm OD", abs(die_small - 50.0) < 0.001)
test("Die ring: large >=40 -> 70mm OD", abs(die_large - 70.0) < 0.001)

# ===========================================================================
# 2. SVG ENGINE: NESTING — EDGE BIAS DIRECTIONS
# ===========================================================================
print("\n--- 2. Nesting with Edge Bias ---")

# Use a generous sheet with a few pads so placement is deterministic
sheet_w, sheet_h = 200.0, 150.0
pads_3x20 = make_pads((20.0, 3))

def avg_position(placed):
    """Return average (cx, cy) of placed discs."""
    if not placed:
        return (0, 0)
    xs = [cx for _, cx, cy, r in placed]
    ys = [cy for _, cx, cy, r in placed]
    return (sum(xs) / len(xs), sum(ys) / len(ys))


# Center bias (default)
s_center = make_settings(edge_bias="center")
placed_c, fp, ft = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_center)
test("Center bias: all 3 pads placed", len(placed_c) == 3)
avg_cx, avg_cy = avg_position(placed_c)
# Center bias should place near top-left (default scan order is top-left to bottom-right)
# but the key thing is they all fit
test("Center bias: pads within sheet bounds",
     all(cx - r >= 0 and cx + r <= sheet_w and cy - r >= 0 and cy + r <= sheet_h
         for _, cx, cy, r in placed_c))

# North bias: pads should cluster toward top (low Y)
s_north = make_settings(edge_bias="n")
placed_n, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_north)
test("North bias: all 3 pads placed", len(placed_n) == 3)
avg_n_y = avg_position(placed_n)[1]

# South bias: pads toward bottom (high Y)
s_south = make_settings(edge_bias="s")
placed_s, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_south)
test("South bias: all 3 pads placed", len(placed_s) == 3)
avg_s_y = avg_position(placed_s)[1]
test("South bias avg Y > North bias avg Y", avg_s_y > avg_n_y,
     f"south={avg_s_y:.1f}, north={avg_n_y:.1f}")

# East bias: pads toward right (high X)
s_east = make_settings(edge_bias="e")
placed_e, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_east)
test("East bias: all 3 pads placed", len(placed_e) == 3)
avg_e_x = avg_position(placed_e)[0]

# West bias: pads toward left (low X)
s_west = make_settings(edge_bias="w")
placed_w, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_west)
test("West bias: all 3 pads placed", len(placed_w) == 3)
avg_w_x = avg_position(placed_w)[0]
test("East bias avg X > West bias avg X", avg_e_x > avg_w_x,
     f"east={avg_e_x:.1f}, west={avg_w_x:.1f}")

# Corner biases: verify smallest-first ordering and corner proximity
# NW corner: should pack near (0, 0)
s_nw = make_settings(edge_bias="nw")
placed_nw, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_nw)
test("NW corner bias: all 3 pads placed", len(placed_nw) == 3)
# Check closest pad is near the NW corner
dists_nw = [math.sqrt(cx**2 + cy**2) for _, cx, cy, r in placed_nw]
test("NW corner bias: closest pad near origin",
     min(dists_nw) < sheet_w / 3,
     f"min dist to (0,0)={min(dists_nw):.1f}")

# SE corner: should pack near (sheet_w, sheet_h)
s_se = make_settings(edge_bias="se")
placed_se, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_se)
test("SE corner bias: all 3 pads placed", len(placed_se) == 3)
dists_se = [math.sqrt((cx - sheet_w)**2 + (cy - sheet_h)**2) for _, cx, cy, r in placed_se]
test("SE corner bias: closest pad near SE corner",
     min(dists_se) < sheet_w / 3,
     f"min dist to SE={min(dists_se):.1f}")

# NE corner
s_ne = make_settings(edge_bias="ne")
placed_ne, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_ne)
test("NE corner bias: all 3 pads placed", len(placed_ne) == 3)
dists_ne = [math.sqrt((cx - sheet_w)**2 + cy**2) for _, cx, cy, r in placed_ne]
test("NE corner bias: closest pad near NE corner",
     min(dists_ne) < sheet_w / 3,
     f"min dist to NE={min(dists_ne):.1f}")

# SW corner
s_sw = make_settings(edge_bias="sw")
placed_sw, _, _ = _nest_discs(pads_3x20, 'felt', sheet_w, sheet_h, s_sw)
test("SW corner bias: all 3 pads placed", len(placed_sw) == 3)
dists_sw = [math.sqrt(cx**2 + (cy - sheet_h)**2) for _, cx, cy, r in placed_sw]
test("SW corner bias: closest pad near SW corner",
     min(dists_sw) < sheet_w / 3,
     f"min dist to SW={min(dists_sw):.1f}")

# Corner bias: smallest discs closest to corner (sort check)
# Use mixed sizes to verify sorting
mixed_pads = make_pads((10.0, 2), (25.0, 2))
s_nw2 = make_settings(edge_bias="nw")
placed_mixed_nw, _, _ = _nest_discs(mixed_pads, 'felt', sheet_w, sheet_h, s_nw2)
test("NW corner mixed sizes: all 4 placed", len(placed_mixed_nw) == 4)
# The smallest discs (10mm) should be closer to corner than the 25mm ones
small_dists = [math.sqrt(cx**2 + cy**2) for ps, cx, cy, r in placed_mixed_nw if ps == 10.0]
big_dists = [math.sqrt(cx**2 + cy**2) for ps, cx, cy, r in placed_mixed_nw if ps == 25.0]
if small_dists and big_dists:
    test("NW corner: small pads closer to corner than big pads",
         min(small_dists) < min(big_dists),
         f"small min={min(small_dists):.1f}, big min={min(big_dists):.1f}")
else:
    test("NW corner: small pads closer to corner than big pads", False, "missing placements")


# ===========================================================================
# 3. SVG ENGINE: can_all_pads_fit()
# ===========================================================================
print("\n--- 3. can_all_pads_fit() ---")

settings = make_settings()

# Easy case: small pads on big sheet
test("Easy fit: 3x 20mm on 200x150mm sheet",
     can_all_pads_fit(make_pads((20.0, 3)), 'felt', 200, 150, settings))

# Tight fit: should fail when sheet is too small
test("No fit: 10x 50mm on 50x50mm sheet",
     not can_all_pads_fit(make_pads((50.0, 10)), 'felt', 50, 50, settings))

# Single pad on adequate sheet
test("Single pad fits",
     can_all_pads_fit(make_pads((20.0, 1)), 'felt', 50, 50, settings))

# Very large pad that exceeds sheet
test("Oversized pad does not fit",
     not can_all_pads_fit(make_pads((100.0, 1)), 'felt', 50, 50, settings))

# Multiple materials
test("Card pads fit on large sheet",
     can_all_pads_fit(make_pads((18.0, 5)), 'card', 200, 150, settings))

test("Leather pads fit on large sheet",
     can_all_pads_fit(make_pads((20.0, 3)), 'leather', 200, 150, settings))


# ===========================================================================
# 4. POLYGON NESTING
# ===========================================================================
print("\n--- 4. Polygon Nesting ---")

# Simple square polygon (100mm x 100mm)
square_poly = [(0, 0), (100, 0), (100, 100), (0, 100)]

# Point-in-polygon tests
test("Point in square: center", _point_in_polygon(50, 50, square_poly))
test("Point in square: near corner", _point_in_polygon(5, 5, square_poly))
test("Point outside square", not _point_in_polygon(150, 50, square_poly))
test("Point outside square (negative)", not _point_in_polygon(-5, 50, square_poly))

# Circle fits in polygon
test("Circle fits: small circle in center",
     _circle_fits_in_polygon(50, 50, 10, square_poly, spacing_mm=1.0))
test("Circle too big: radius exceeds polygon",
     not _circle_fits_in_polygon(50, 50, 60, square_poly, spacing_mm=1.0))
test("Circle near edge: just barely fits",
     _circle_fits_in_polygon(50, 50, 48, square_poly, spacing_mm=1.0))
test("Circle near edge: too close (spacing)",
     not _circle_fits_in_polygon(50, 50, 49.5, square_poly, spacing_mm=1.0))

# Distance to segment
dist = _distance_point_to_segment(5, 5, 0, 0, 10, 0)
test("Distance point to segment", abs(dist - 5.0) < 0.001, f"got {dist}")

# Nest into square polygon
pads_for_poly = make_pads((15.0, 4))
placed_poly, fp_poly, ft_poly = _nest_discs(
    pads_for_poly, 'felt', 100, 100, settings, polygon=square_poly)
test("Polygon nesting: all 4 pads placed in 100mm square",
     fp_poly == ft_poly and ft_poly == 4,
     f"placed {fp_poly}/{ft_poly}")

# Verify all placed discs are inside polygon
all_inside = True
for ps, cx, cy, r in placed_poly:
    if not _circle_fits_in_polygon(cx, cy, r, square_poly, spacing_mm=0.5):
        all_inside = False
        break
test("Polygon nesting: all discs inside polygon", all_inside)

# Triangle polygon (right triangle)
triangle_poly = [(0, 0), (100, 0), (0, 100)]
pads_tri = make_pads((10.0, 3))
placed_tri, fp_tri, ft_tri = _nest_discs(
    pads_tri, 'felt', 100, 100, settings, polygon=triangle_poly)
test("Triangle polygon: 3 small pads placed",
     fp_tri == 3, f"placed {fp_tri}/3")

# Can't fit large pad in small polygon
tiny_poly = [(0, 0), (10, 0), (10, 10), (0, 10)]
placed_tiny, fp_tiny, ft_tiny = _nest_discs(
    make_pads((20.0, 1)), 'felt', 10, 10, settings, polygon=tiny_poly)
test("Tiny polygon: large pad does not fit", fp_tiny == 0)


# ===========================================================================
# 5. SCRAP MODE HELPERS
# ===========================================================================
print("\n--- 5. Scrap Mode ---")

settings = make_settings()

# try_nest_partial on adequate sheet
original_pads = make_pads((18.0, 5), (25.0, 3))
placed_s, remaining_s, any_placed_s = try_nest_partial(
    original_pads, 'felt', 200, 150, settings)
test("Scrap: all 8 pads placed on large sheet",
     len(placed_s) == 8 and len(remaining_s) == 0 and any_placed_s)

# try_nest_partial on small sheet — some placed, some remaining
placed_sm, remaining_sm, any_placed_sm = try_nest_partial(
    make_pads((20.0, 10)), 'felt', 60, 40, settings)
test("Scrap: partial placement on small sheet",
     any_placed_sm and len(remaining_sm) > 0,
     f"placed={len(placed_sm)}, remaining sizes={remaining_sm}")

# compute_remaining_pads
original = make_pads((18.0, 5), (25.0, 3))
# Simulate placing 3x 18mm and 1x 25mm
fake_placed = [
    (18.0, 10, 10, 8.625), (18.0, 30, 10, 8.625), (18.0, 50, 10, 8.625),
    (25.0, 80, 30, 12.125),
]
remaining = compute_remaining_pads(original, fake_placed)
# Should have 2x 18mm and 2x 25mm left
remaining_dict = {r['size']: r['qty'] for r in remaining}
test("compute_remaining: 2x 18mm left",
     remaining_dict.get(18.0) == 2,
     f"got {remaining_dict}")
test("compute_remaining: 2x 25mm left",
     remaining_dict.get(25.0) == 2,
     f"got {remaining_dict}")

# Max pads should be excluded from remaining
original_with_max = [{'size': 18.0, 'qty': 3}, {'size': 10.0, 'qty': 'max'}]
fake_placed_max = [
    (18.0, 10, 10, 8.625), (18.0, 30, 10, 8.625),
    (10.0, 50, 10, 4.625), (10.0, 70, 10, 4.625),
]
remaining_max = compute_remaining_pads(original_with_max, fake_placed_max)
remaining_max_dict = {r['size']: r['qty'] for r in remaining_max}
test("compute_remaining: max pads excluded from remaining",
     10.0 not in remaining_max_dict,
     f"got {remaining_max_dict}")
test("compute_remaining: 1x 18mm left after placing 2",
     remaining_max_dict.get(18.0) == 1,
     f"got {remaining_max_dict}")

# try_nest_partial on zero-size sheet — nothing placed
placed_z, remaining_z, any_z = try_nest_partial(
    make_pads((20.0, 3)), 'felt', 5, 5, settings)
remaining_z_total = sum(r['qty'] for r in remaining_z)
test("Scrap: nothing fits on tiny sheet",
     not any_z and remaining_z_total == 3,
     f"placed={len(placed_z)}, remaining_total={remaining_z_total}")


# ===========================================================================
# 6. CONFIG: SETTINGS MERGE & DEFAULTS
# ===========================================================================
print("\n--- 6. Config: Settings Merge & Defaults ---")

# Test that DEFAULT_SETTINGS has all expected top-level keys
test("DEFAULT_SETTINGS has edge_bias", "edge_bias" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has visible_tabs", "visible_tabs" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has toner_settings", "toner_settings" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has tuner_settings", "tuner_settings" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has gcode_settings", "gcode_settings" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has tooling_settings", "tooling_settings" in DEFAULT_SETTINGS)
test("DEFAULT_SETTINGS has filled_overscan_enabled", "filled_overscan_enabled" in DEFAULT_SETTINGS)

# Simulate loading with a partial/old settings file
with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as tmp:
    old_settings = {
        "units": "mm",
        "felt_offset": 1.0,
        "gcode_settings": {
            "felt": {
                "cut_speed": 999,
                "cut_power": 88,
            }
        }
    }
    json.dump(old_settings, tmp)
    tmp_path = tmp.name

# Temporarily override SETTINGS_FILE to test merge
import config
original_settings_file = config.SETTINGS_FILE
config.SETTINGS_FILE = tmp_path
try:
    merged = load_settings()

    test("Merge: preserves existing 'units' value",
         merged["units"] == "mm")
    test("Merge: preserves existing 'felt_offset'",
         merged["felt_offset"] == 1.0)
    test("Merge: fills in missing 'edge_bias' from defaults",
         merged.get("edge_bias") == DEFAULT_SETTINGS["edge_bias"])
    test("Merge: fills in missing 'visible_tabs' from defaults",
         merged.get("visible_tabs") == DEFAULT_SETTINGS["visible_tabs"])
    test("Merge: fills in missing 'toner_settings' from defaults",
         merged.get("toner_settings") == DEFAULT_SETTINGS["toner_settings"])
    test("Merge: gcode_settings.felt preserves user cut_speed",
         merged["gcode_settings"]["felt"]["cut_speed"] == 999)
    test("Merge: gcode_settings.felt fills in missing engraving_speed",
         merged["gcode_settings"]["felt"]["engraving_speed"] == DEFAULT_SETTINGS["gcode_settings"]["felt"]["engraving_speed"])
    test("Merge: gcode_settings.card filled from defaults entirely",
         merged["gcode_settings"]["card"] == DEFAULT_SETTINGS["gcode_settings"]["card"])
finally:
    config.SETTINGS_FILE = original_settings_file
    os.unlink(tmp_path)


# ===========================================================================
# 7. VISIBLE TABS DEFAULTS
# ===========================================================================
print("\n--- 7. Visible Tabs Defaults ---")

vt = DEFAULT_SETTINGS["visible_tabs"]
test("Tuner defaults to False", vt.get("Tuner") is False)
test("Toner defaults to False", vt.get("Toner") is False)
test("Key Height Library defaults to True", vt.get("Key Height Library") is True)
test("Serial Lookup defaults to True", vt.get("Serial Lookup") is True)
test("Screw Specs defaults to True", vt.get("Screw Specs") is True)
test("Tooling defaults to True", vt.get("Tooling") is True)


# ===========================================================================
# 8. EDGE CASES
# ===========================================================================
print("\n--- 8. Edge Cases ---")

settings = make_settings()

# Empty pad list
placed_empty, fp_e, ft_e = _nest_discs([], 'felt', 200, 150, settings)
test("Empty pad list: no placements", len(placed_empty) == 0 and fp_e == 0 and ft_e == 0)

# Single pad
placed_one, fp_o, ft_o = _nest_discs(make_pads((15.0, 1)), 'felt', 200, 150, settings)
test("Single pad: placed successfully", len(placed_one) == 1 and fp_o == 1)

# Max fill: fill a small sheet with small pads
max_pads = [{'size': 10.0, 'qty': 'max'}]
placed_max, fp_m, ft_m = _nest_discs(max_pads, 'felt', 100, 100, settings)
test("Max fill: multiple pads placed", len(placed_max) > 5,
     f"placed {len(placed_max)} pads")

# Very large pad on small sheet
placed_big, fp_b, ft_b = _nest_discs(make_pads((200.0, 1)), 'felt', 50, 50, settings)
test("Oversized pad: not placed", len(placed_big) == 0 and fp_b == 0)

# can_all_pads_fit with empty list
test("can_all_pads_fit: empty list returns True",
     can_all_pads_fit([], 'felt', 200, 150, settings))

# Leather sizing: tiny pad (< 6mm)
tiny_leather = get_disc_diameter(4.0, 'leather', settings)
test("Leather tiny pad: still produces valid diameter", tiny_leather > 4.0,
     f"got {tiny_leather}")

# Felt thickness unit conversion
s_inch = make_settings(felt_thickness=0.125, felt_thickness_unit="in")
ft_mm = get_felt_thickness_mm(s_inch)
test("Felt thickness: inch to mm conversion",
     abs(ft_mm - 3.175) < 0.001, f"got {ft_mm}")

s_mm = make_settings(felt_thickness=3.175, felt_thickness_unit="mm")
ft_mm2 = get_felt_thickness_mm(s_mm)
test("Felt thickness: mm stays mm",
     abs(ft_mm2 - 3.175) < 0.001, f"got {ft_mm2}")

# No overlap check: verify placed discs don't overlap
pads_many = make_pads((15.0, 5), (20.0, 5), (25.0, 3))
placed_many, _, _ = _nest_discs(pads_many, 'felt', 300, 200, settings)
overlap_found = False
for i in range(len(placed_many)):
    for j in range(i + 1, len(placed_many)):
        _, x1, y1, r1 = placed_many[i]
        _, x2, y2, r2 = placed_many[j]
        dist = math.sqrt((x1 - x2)**2 + (y1 - y2)**2)
        if dist < r1 + r2 + 0.5:  # 0.5mm tolerance (spacing is 1.0)
            overlap_found = True
            break
    if overlap_found:
        break
test("No overlapping discs in 13-pad layout",
     not overlap_found, f"{len(placed_many)} discs placed")

# Polygon nesting with can_all_pads_fit
test("can_all_pads_fit with polygon: fits",
     can_all_pads_fit(make_pads((10.0, 2)), 'felt', 100, 100, settings,
                      polygon=square_poly))
test("can_all_pads_fit with polygon: doesn't fit",
     not can_all_pads_fit(make_pads((60.0, 2)), 'felt', 100, 100, settings,
                          polygon=square_poly))


# ===========================================================================
# 9. MODULE IMPORTS
# ===========================================================================
print("\n--- 9. Module Imports ---")

try:
    import gcode_engine
    test("gcode_engine imports successfully", True)
    test("gcode_engine has generate_gcode_from_placed",
         hasattr(gcode_engine, 'generate_gcode_from_placed'))
    test("gcode_engine has generate_gcode",
         hasattr(gcode_engine, 'generate_gcode'))
except Exception as e:
    test("gcode_engine imports successfully", False, str(e))

try:
    import toner_engine
    test("toner_engine imports successfully", True)
    test("toner_engine has AUDIO_AVAILABLE flag",
         hasattr(toner_engine, 'AUDIO_AVAILABLE'))
except Exception as e:
    test("toner_engine imports successfully", False, str(e))

try:
    import tuner_engine
    test("tuner_engine imports successfully", True)
    test("tuner_engine has TunerEngine class",
         hasattr(tuner_engine, 'TunerEngine'))
except Exception as e:
    test("tuner_engine imports successfully", False, str(e))


# ===========================================================================
# SUMMARY
# ===========================================================================
print("\n" + "=" * 60)
passed = sum(1 for _, ok in results if ok)
failed = sum(1 for _, ok in results if not ok)
total = len(results)
print(f"RESULTS: {passed} passed, {failed} failed, {total} total")

if failed > 0:
    print("\nFailed tests:")
    for name, ok in results:
        if not ok:
            print(f"  - {name}")

print("=" * 60)
sys.exit(0 if failed == 0 else 1)
