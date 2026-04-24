"""
Test suite for dart range mode feature.

Tests get_dart_settings_for_size() helper, and verifies SVG/G-code engines
produce correct output in both universal and range modes.
"""

import sys
import os
import copy
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from config import DEFAULT_SETTINGS, get_dart_settings_for_size
from svg_engine import get_disc_diameter, _render_svg_discs
from gcode_engine import generate_gcode_from_placed

passed = 0
failed = 0


def check(label, condition):
    global passed, failed
    if condition:
        print(f"  PASS  {label}")
        passed += 1
    else:
        print(f"  FAIL  {label}")
        failed += 1


def make_settings(**overrides):
    s = copy.deepcopy(DEFAULT_SETTINGS)
    s.update(overrides)
    return s


# =============================================================================
print("--- get_dart_settings_for_size: Universal Mode ---")

s = make_settings()
check("Universal: pad below threshold returns settings", get_dart_settings_for_size(10.0, s) is not None)
check("Universal: pad at threshold returns None", get_dart_settings_for_size(18.0, s) is None)
check("Universal: pad above threshold returns None", get_dart_settings_for_size(25.0, s) is None)
check("Universal: returns correct overwrap", get_dart_settings_for_size(10.0, s)["overwrap"] == 0.5)
check("Universal: returns correct wrap_bonus", get_dart_settings_for_size(10.0, s)["wrap_bonus"] == 0.75)
check("Universal: returns correct frequency_multiplier", get_dart_settings_for_size(10.0, s)["frequency_multiplier"] == 1.0)
check("Universal: returns correct shape_factor", get_dart_settings_for_size(10.0, s)["shape_factor"] == 0.0)
check("Universal: returns correct engraving_on", get_dart_settings_for_size(10.0, s)["engraving_on"] is True)

# Custom universal settings
s2 = make_settings(dart_threshold=12.0, dart_overwrap=0.3, dart_wrap_bonus=1.0)
check("Custom threshold: 11.9 returns settings", get_dart_settings_for_size(11.9, s2) is not None)
check("Custom threshold: 12.0 returns None", get_dart_settings_for_size(12.0, s2) is None)
check("Custom overwrap: returns 0.3", get_dart_settings_for_size(10.0, s2)["overwrap"] == 0.3)
check("Custom wrap_bonus: returns 1.0", get_dart_settings_for_size(10.0, s2)["wrap_bonus"] == 1.0)

# Darts disabled
s3 = make_settings(darts_enabled=False)
check("Darts disabled: returns None for any size", get_dart_settings_for_size(10.0, s3) is None)
check("Darts disabled: returns None for small pad", get_dart_settings_for_size(5.0, s3) is None)

# =============================================================================
print("\n--- get_dart_settings_for_size: Range Mode ---")

ranges = [
    {"min_size": 7.0, "max_size": 11.5, "overwrap": 0.3, "wrap_bonus": 0.5,
     "frequency_multiplier": 1.2, "shape_factor": 0.1, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.0}},
    {"min_size": 12.0, "max_size": 14.0, "overwrap": 0.5, "wrap_bonus": 0.75,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
    {"min_size": 14.5, "max_size": 20.0, "overwrap": 0.7, "wrap_bonus": 1.0,
     "frequency_multiplier": 0.8, "shape_factor": 0.3, "engraving_on": False,
     "engraving_loc": {"mode": "centered", "value": 0}},
]

sr = make_settings(dart_range_mode="range", dart_ranges=ranges)

check("Range: 8mm matches first range", get_dart_settings_for_size(8.0, sr) is not None)
check("Range: 8mm overwrap=0.3", get_dart_settings_for_size(8.0, sr)["overwrap"] == 0.3)
check("Range: 8mm freq_mult=1.2", get_dart_settings_for_size(8.0, sr)["frequency_multiplier"] == 1.2)

check("Range: 13mm matches second range", get_dart_settings_for_size(13.0, sr) is not None)
check("Range: 13mm overwrap=0.5", get_dart_settings_for_size(13.0, sr)["overwrap"] == 0.5)

check("Range: 15mm matches third range", get_dart_settings_for_size(15.0, sr) is not None)
check("Range: 15mm overwrap=0.7", get_dart_settings_for_size(15.0, sr)["overwrap"] == 0.7)
check("Range: 15mm engraving_on=False", get_dart_settings_for_size(15.0, sr)["engraving_on"] is False)

check("Range: 6mm no match returns None", get_dart_settings_for_size(6.0, sr) is None)
check("Range: 11.8mm gap returns None", get_dart_settings_for_size(11.8, sr) is None)
check("Range: 25mm no match returns None", get_dart_settings_for_size(25.0, sr) is None)

# Boundary tests
check("Range: 7.0mm (min boundary) matches", get_dart_settings_for_size(7.0, sr) is not None)
check("Range: 11.5mm (max boundary) matches", get_dart_settings_for_size(11.5, sr) is not None)
check("Range: 12.0mm (second range min) matches", get_dart_settings_for_size(12.0, sr) is not None)
check("Range: 20.0mm (third range max) matches", get_dart_settings_for_size(20.0, sr) is not None)

# Darts disabled overrides range mode
sr_off = make_settings(dart_range_mode="range", dart_ranges=ranges, darts_enabled=False)
check("Range + disabled: returns None", get_dart_settings_for_size(8.0, sr_off) is None)

# Empty ranges
sr_empty = make_settings(dart_range_mode="range", dart_ranges=[])
check("Range empty: returns None", get_dart_settings_for_size(10.0, sr_empty) is None)

# =============================================================================
print("\n--- get_disc_diameter: Range Mode ---")

# In universal mode, leather below threshold gets dart bonus (larger diameter)
su = make_settings()
d_universal_star = get_disc_diameter(10.0, 'leather', su)
d_universal_plain = get_disc_diameter(25.0, 'leather', su)

# Star pads should be larger (they have the wrap bonus)
check("Universal: star pad has larger diameter than plain", d_universal_star / 10.0 > d_universal_plain / 25.0)

# In range mode, pad in range gets dart bonus
sr2 = make_settings(dart_range_mode="range", dart_ranges=[
    {"min_size": 7.0, "max_size": 15.0, "overwrap": 0.5, "wrap_bonus": 0.75,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
])
d_range_star = get_disc_diameter(10.0, 'leather', sr2)
d_range_no_match = get_disc_diameter(25.0, 'leather', sr2)

check("Range: pad in range gets dart bonus", d_range_star == d_universal_star)
check("Range: pad outside range gets plain diameter", d_range_no_match == d_universal_plain)

# Different wrap_bonus per range
sr3 = make_settings(dart_range_mode="range", dart_ranges=[
    {"min_size": 7.0, "max_size": 11.5, "overwrap": 0.5, "wrap_bonus": 0.5,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
    {"min_size": 12.0, "max_size": 18.0, "overwrap": 0.5, "wrap_bonus": 1.5,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
])
d_small_bonus = get_disc_diameter(10.0, 'leather', sr3)
d_large_bonus = get_disc_diameter(15.0, 'leather', sr3)

# Larger wrap bonus on the 12-18 range should produce a bigger diameter for same pad
# But 15mm pad is bigger than 10mm, so compare ratios instead
check("Range: different wrap_bonus produces different diameters",
      d_small_bonus != d_large_bonus)

# Non-leather materials unaffected by range mode
check("Range: felt unaffected", get_disc_diameter(10.0, 'felt', sr2) == get_disc_diameter(10.0, 'felt', su))
check("Range: card unaffected", get_disc_diameter(10.0, 'card', sr2) == get_disc_diameter(10.0, 'card', su))

# =============================================================================
print("\n--- SVG Rendering: Range Mode ---")

import svgwrite

def render_svg_test(pad_size, material, settings):
    """Render a single pad and return the SVG content."""
    placed = [(pad_size, 20.0, 20.0, get_disc_diameter(pad_size, material, settings) / 2)]
    dwg = svgwrite.Drawing(size=("100mm", "100mm"))
    _render_svg_discs(dwg, placed, material, 0, settings, False, 0.1)
    return dwg.tostring()

# Universal: leather below threshold produces star path
svg_star = render_svg_test(10.0, 'leather', su)
check("Universal SVG: star pad has path element", '<path' in svg_star)

# Universal: leather above threshold produces circle
svg_circle = render_svg_test(25.0, 'leather', su)
check("Universal SVG: plain pad has circle element", '<circle' in svg_circle)

# Range mode: pad in range produces star
svg_range_star = render_svg_test(10.0, 'leather', sr2)
check("Range SVG: pad in range has path element", '<path' in svg_range_star)

# Range mode: pad not in range produces circle
svg_range_circle = render_svg_test(25.0, 'leather', sr2)
check("Range SVG: pad outside range has circle element", '<circle' in svg_range_circle)

# Range mode: pad in gap between ranges produces circle
sr_gap = make_settings(dart_range_mode="range", dart_ranges=[
    {"min_size": 7.0, "max_size": 10.0, "overwrap": 0.5, "wrap_bonus": 0.75,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
    {"min_size": 15.0, "max_size": 20.0, "overwrap": 0.5, "wrap_bonus": 0.75,
     "frequency_multiplier": 1.0, "shape_factor": 0.0, "engraving_on": True,
     "engraving_loc": {"mode": "from_outside", "value": 2.5}},
])
svg_gap = render_svg_test(12.0, 'leather', sr_gap)
check("Range SVG: pad in gap between ranges is circle", '<circle' in svg_gap)

# =============================================================================
print("\n--- G-code: Range Mode ---")

import tempfile

def gcode_test(pad_size, material, settings):
    """Generate G-code for a single pad and return the content."""
    r = get_disc_diameter(pad_size, material, settings) / 2
    placed = [(pad_size, 20.0, 20.0, r)]
    with tempfile.NamedTemporaryFile(mode='w', suffix='.gcode', delete=False) as f:
        path = f.name
    try:
        generate_gcode_from_placed(placed, material, 100, 100, path, 0, settings)
        with open(path, 'r') as f:
            return f.read()
    finally:
        os.unlink(path)

# Universal: leather star pad generates star pattern (not simple circle)
gc_star = gcode_test(10.0, 'leather', su)
# Star patterns have many more points than a simple circle
lines_star = [l for l in gc_star.split('\n') if l.startswith('G1') and 'S' not in l.split('G1')[1][:5]]

gc_circle = gcode_test(25.0, 'leather', su)
check("Universal G-code: star pad generates output", len(gc_star) > 0)
check("Universal G-code: plain pad generates output", len(gc_circle) > 0)

# Range mode
gc_range_star = gcode_test(10.0, 'leather', sr2)
gc_range_circle = gcode_test(25.0, 'leather', sr2)
check("Range G-code: pad in range generates output", len(gc_range_star) > 0)
check("Range G-code: pad outside range generates output", len(gc_range_circle) > 0)

# =============================================================================
print("\n--- Backward Compatibility ---")

# Settings without range keys should work (universal mode by default)
s_legacy = copy.deepcopy(DEFAULT_SETTINGS)
# Simulate old config without range keys
s_legacy.pop("dart_range_mode", None)
s_legacy.pop("dart_ranges", None)
check("Legacy: missing range keys defaults to universal",
      get_dart_settings_for_size(10.0, s_legacy) is not None)
check("Legacy: missing range keys returns correct overwrap",
      get_dart_settings_for_size(10.0, s_legacy)["overwrap"] == 0.5)

# =============================================================================
print(f"\n{'='*60}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed} tests")
print(f"{'='*60}")
if failed:
    print("SOME TESTS FAILED")
    sys.exit(1)
else:
    print("ALL TESTS PASSED")
