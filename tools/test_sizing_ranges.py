"""
Test suite for sizing range mode, engraving settings range mode,
and engraving placement range mode features.

Tests get_sizing_for_size(), get_engraving_settings_for_size(),
get_engraving_placement_for_size() helpers, and verifies that
get_disc_diameter() and should_have_center_hole() respect range settings.
"""

import sys
import os
import copy
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from config import (DEFAULT_SETTINGS, get_sizing_for_size,
                    get_engraving_settings_for_size,
                    get_engraving_placement_for_size)
from svg_engine import get_disc_diameter, should_have_center_hole

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
print("--- get_sizing_for_size: Universal Mode ---")

s = make_settings()
result = get_sizing_for_size(10.0, s)
check("Universal: returns a dict", isinstance(result, dict))
check("Universal: has felt_offset", "felt_offset" in result)
check("Universal: felt_offset matches default", result["felt_offset"] == 0.75)
check("Universal: card_to_felt_offset matches default", result["card_to_felt_offset"] == 2.0)
check("Universal: leather_wrap_multiplier matches default", result["leather_wrap_multiplier"] == 1.0)
check("Universal: min_hole_size matches default", result["min_hole_size"] == 16.5)
check("Universal: felt_thickness matches default", result["felt_thickness"] == 3.175)
check("Universal: felt_thickness_unit matches default", result["felt_thickness_unit"] == "mm")

# Custom universal values
s2 = make_settings(felt_offset=1.0, card_to_felt_offset=1.5, min_hole_size=20.0)
r2 = get_sizing_for_size(10.0, s2)
check("Custom universal: felt_offset=1.0", r2["felt_offset"] == 1.0)
check("Custom universal: card_to_felt_offset=1.5", r2["card_to_felt_offset"] == 1.5)
check("Custom universal: min_hole_size=20.0", r2["min_hole_size"] == 20.0)

# Universal mode returns same result for any pad size
check("Universal: same result for 5mm",
      get_sizing_for_size(5.0, s)["felt_offset"] == 0.75)
check("Universal: same result for 50mm",
      get_sizing_for_size(50.0, s)["felt_offset"] == 0.75)

# =============================================================================
print("\n--- get_sizing_for_size: Range Mode ---")

sizing_ranges = [
    {"min_size": 7.0, "max_size": 12.0,
     "felt_offset": 0.5, "card_to_felt_offset": 1.0,
     "leather_wrap_multiplier": 0.8, "min_hole_size": 10.0,
     "felt_thickness": 2.5, "felt_thickness_unit": "mm"},
    {"min_size": 15.0, "max_size": 25.0,
     "felt_offset": 1.0, "card_to_felt_offset": 2.5,
     "leather_wrap_multiplier": 1.2, "min_hole_size": 20.0,
     "felt_thickness": 4.0, "felt_thickness_unit": "mm"},
]

sr = make_settings(sizing_range_mode="range", sizing_ranges=sizing_ranges)

# Match first range
r_first = get_sizing_for_size(10.0, sr)
check("Range: 10mm matches first range", r_first is not None)
check("Range: 10mm felt_offset=0.5", r_first["felt_offset"] == 0.5)
check("Range: 10mm card_to_felt_offset=1.0", r_first["card_to_felt_offset"] == 1.0)
check("Range: 10mm min_hole_size=10.0", r_first["min_hole_size"] == 10.0)
check("Range: 10mm felt_thickness=2.5", r_first["felt_thickness"] == 2.5)

# Match second range
r_second = get_sizing_for_size(20.0, sr)
check("Range: 20mm matches second range", r_second is not None)
check("Range: 20mm felt_offset=1.0", r_second["felt_offset"] == 1.0)
check("Range: 20mm min_hole_size=20.0", r_second["min_hole_size"] == 20.0)

# Boundary tests
check("Range: 7.0mm (min boundary) matches first", get_sizing_for_size(7.0, sr)["felt_offset"] == 0.5)
check("Range: 12.0mm (max boundary) matches first", get_sizing_for_size(12.0, sr)["felt_offset"] == 0.5)
check("Range: 15.0mm (min boundary) matches second", get_sizing_for_size(15.0, sr)["felt_offset"] == 1.0)
check("Range: 25.0mm (max boundary) matches second", get_sizing_for_size(25.0, sr)["felt_offset"] == 1.0)

# No match falls back to universal (NOT None like darts!)
r_gap = get_sizing_for_size(13.0, sr)
check("Range: 13mm (gap) falls back to universal, not None", r_gap is not None)
check("Range: 13mm (gap) gets default felt_offset", r_gap["felt_offset"] == 0.75)

r_below = get_sizing_for_size(5.0, sr)
check("Range: 5mm (below all ranges) falls back to universal", r_below is not None)
check("Range: 5mm gets default felt_offset", r_below["felt_offset"] == 0.75)

r_above = get_sizing_for_size(30.0, sr)
check("Range: 30mm (above all ranges) falls back to universal", r_above is not None)
check("Range: 30mm gets default felt_offset", r_above["felt_offset"] == 0.75)

# Empty ranges fall back to universal
sr_empty = make_settings(sizing_range_mode="range", sizing_ranges=[])
r_empty = get_sizing_for_size(10.0, sr_empty)
check("Range empty: falls back to universal", r_empty is not None)
check("Range empty: gets default felt_offset", r_empty["felt_offset"] == 0.75)

# =============================================================================
print("\n--- get_engraving_settings_for_size: Universal Mode ---")

s = make_settings()
eng = get_engraving_settings_for_size(10.0, s)
check("Eng universal: returns a dict", isinstance(eng, dict))
check("Eng universal: engraving_on is True", eng["engraving_on"] is True)
check("Eng universal: has engraving_font_size", "engraving_font_size" in eng)
check("Eng universal: felt font size is 3.0", eng["engraving_font_size"]["felt"] == 3.0)
check("Eng universal: leather font size is 3.0", eng["engraving_font_size"]["leather"] == 3.0)

# Custom universal
s3 = make_settings(engraving_on=False)
eng3 = get_engraving_settings_for_size(10.0, s3)
check("Eng custom universal: engraving_on=False", eng3["engraving_on"] is False)

# =============================================================================
print("\n--- get_engraving_settings_for_size: Range Mode ---")

eng_ranges = [
    {"min_size": 7.0, "max_size": 12.0,
     "engraving_on": True,
     "engraving_font_size": {"felt": 2.0, "card": 2.0, "leather": 2.0, "exact_size": 2.0}},
    {"min_size": 15.0, "max_size": 25.0,
     "engraving_on": False,
     "engraving_font_size": {"felt": 4.0, "card": 4.0, "leather": 4.0, "exact_size": 4.0}},
]

se = make_settings(engraving_settings_range_mode="range", engraving_settings_ranges=eng_ranges)

eng_first = get_engraving_settings_for_size(10.0, se)
check("Eng range: 10mm matches first range", eng_first["engraving_on"] is True)
check("Eng range: 10mm felt font=2.0", eng_first["engraving_font_size"]["felt"] == 2.0)

eng_second = get_engraving_settings_for_size(20.0, se)
check("Eng range: 20mm matches second range", eng_second["engraving_on"] is False)
check("Eng range: 20mm felt font=4.0", eng_second["engraving_font_size"]["felt"] == 4.0)

# No match falls back to universal
eng_gap = get_engraving_settings_for_size(13.0, se)
check("Eng range: 13mm (gap) falls back to universal", eng_gap is not None)
check("Eng range: 13mm gets default engraving_on=True", eng_gap["engraving_on"] is True)
check("Eng range: 13mm gets default felt font=3.0", eng_gap["engraving_font_size"]["felt"] == 3.0)

# Empty ranges
se_empty = make_settings(engraving_settings_range_mode="range", engraving_settings_ranges=[])
eng_empty = get_engraving_settings_for_size(10.0, se_empty)
check("Eng range empty: falls back to universal", eng_empty["engraving_on"] is True)

# =============================================================================
print("\n--- get_engraving_placement_for_size: Universal Mode ---")

s = make_settings()
plc = get_engraving_placement_for_size(10.0, s)
check("Placement universal: returns a dict", isinstance(plc, dict))
check("Placement universal: has engraving_location", "engraving_location" in plc)
check("Placement universal: felt mode is from_inside",
      plc["engraving_location"]["felt"]["mode"] == "from_inside")
check("Placement universal: felt value is 4.0",
      plc["engraving_location"]["felt"]["value"] == 4.0)
check("Placement universal: leather mode is from_outside",
      plc["engraving_location"]["leather"]["mode"] == "from_outside")

# =============================================================================
print("\n--- get_engraving_placement_for_size: Range Mode ---")

plc_ranges = [
    {"min_size": 7.0, "max_size": 12.0,
     "engraving_location": {
         "felt": {"mode": "centered", "value": 0},
         "card": {"mode": "centered", "value": 0},
         "leather": {"mode": "centered", "value": 0},
         "exact_size": {"mode": "centered", "value": 0},
     }},
    {"min_size": 15.0, "max_size": 25.0,
     "engraving_location": {
         "felt": {"mode": "from_outside", "value": 3.0},
         "card": {"mode": "from_outside", "value": 3.0},
         "leather": {"mode": "from_outside", "value": 3.0},
         "exact_size": {"mode": "from_outside", "value": 3.0},
     }},
]

sp = make_settings(engraving_placement_range_mode="range", engraving_placement_ranges=plc_ranges)

plc_first = get_engraving_placement_for_size(10.0, sp)
check("Placement range: 10mm matches first range",
      plc_first["engraving_location"]["felt"]["mode"] == "centered")

plc_second = get_engraving_placement_for_size(20.0, sp)
check("Placement range: 20mm matches second range",
      plc_second["engraving_location"]["felt"]["mode"] == "from_outside")
check("Placement range: 20mm felt value=3.0",
      plc_second["engraving_location"]["felt"]["value"] == 3.0)

# No match falls back to universal
plc_gap = get_engraving_placement_for_size(13.0, sp)
check("Placement range: 13mm (gap) falls back to universal", plc_gap is not None)
check("Placement range: 13mm gets default felt mode=from_inside",
      plc_gap["engraving_location"]["felt"]["mode"] == "from_inside")

# Empty ranges
sp_empty = make_settings(engraving_placement_range_mode="range", engraving_placement_ranges=[])
plc_empty = get_engraving_placement_for_size(10.0, sp_empty)
check("Placement range empty: falls back to universal",
      plc_empty["engraving_location"]["felt"]["mode"] == "from_inside")

# =============================================================================
print("\n--- get_disc_diameter: Sizing Ranges ---")

# Universal baseline
su = make_settings()
d_felt_univ = get_disc_diameter(10.0, 'felt', su)
d_card_univ = get_disc_diameter(10.0, 'card', su)
d_leather_univ = get_disc_diameter(10.0, 'leather', su)

check("Universal felt: 10mm - 0.75 = 9.25", d_felt_univ == 9.25)
check("Universal card: 10mm - (0.75 + 2.0) = 7.25", d_card_univ == 7.25)

# Range mode with different offsets for small pads
sr_sizing = make_settings(
    sizing_range_mode="range",
    sizing_ranges=[
        {"min_size": 7.0, "max_size": 15.0,
         "felt_offset": 0.5, "card_to_felt_offset": 1.0,
         "leather_wrap_multiplier": 1.0, "min_hole_size": 16.5,
         "felt_thickness": 3.175, "felt_thickness_unit": "mm"},
    ]
)

d_felt_range = get_disc_diameter(10.0, 'felt', sr_sizing)
d_card_range = get_disc_diameter(10.0, 'card', sr_sizing)
d_leather_range = get_disc_diameter(10.0, 'leather', sr_sizing)

check("Range felt: 10mm - 0.5 = 9.5", d_felt_range == 9.5)
check("Range card: 10mm - (0.5 + 1.0) = 8.5", d_card_range == 8.5)
check("Range felt differs from universal", d_felt_range != d_felt_univ)
check("Range card differs from universal", d_card_range != d_card_univ)

# Pad outside range falls back to universal sizing
d_felt_outside = get_disc_diameter(25.0, 'felt', sr_sizing)
d_felt_outside_univ = get_disc_diameter(25.0, 'felt', su)
check("Pad outside range: felt matches universal", d_felt_outside == d_felt_outside_univ)

# Different felt_thickness affects leather diameter
sr_thick = make_settings(
    sizing_range_mode="range",
    darts_enabled=False,  # disable darts so leather is plain circle
    sizing_ranges=[
        {"min_size": 7.0, "max_size": 15.0,
         "felt_offset": 0.75, "card_to_felt_offset": 2.0,
         "leather_wrap_multiplier": 1.0, "min_hole_size": 16.5,
         "felt_thickness": 5.0, "felt_thickness_unit": "mm"},
    ]
)
su_no_darts = make_settings(darts_enabled=False)
d_leather_thick = get_disc_diameter(10.0, 'leather', sr_thick)
d_leather_default_thick = get_disc_diameter(10.0, 'leather', su_no_darts)
check("Range: thicker felt produces larger leather diameter",
      d_leather_thick > d_leather_default_thick)

# exact_size unaffected by sizing ranges
check("Range: exact_size unaffected",
      get_disc_diameter(10.0, 'exact_size', sr_sizing) == 10.0)

# =============================================================================
print("\n--- should_have_center_hole: Sizing Ranges ---")

# Universal: default min_hole_size is 16.5
su = make_settings()
check("Universal: 20mm pad gets hole (>= 16.5)", should_have_center_hole(20.0, 3.5, su) is True)
check("Universal: 10mm pad no hole (< 16.5)", should_have_center_hole(10.0, 3.5, su) is False)
check("Universal: 16.5mm pad gets hole (== 16.5)", should_have_center_hole(16.5, 3.5, su) is True)
check("Universal: hole_dia=0 means no hole", should_have_center_hole(20.0, 0, su) is False)

# Range mode with lower min_hole_size for small pads
sr_hole = make_settings(
    sizing_range_mode="range",
    sizing_ranges=[
        {"min_size": 7.0, "max_size": 15.0,
         "felt_offset": 0.75, "card_to_felt_offset": 2.0,
         "leather_wrap_multiplier": 1.0, "min_hole_size": 8.0,
         "felt_thickness": 3.175, "felt_thickness_unit": "mm"},
    ]
)

check("Range: 10mm pad gets hole (min_hole_size=8.0)",
      should_have_center_hole(10.0, 3.5, sr_hole) is True)
check("Range: 7mm pad no hole (7 < 8.0)",
      should_have_center_hole(7.0, 3.5, sr_hole) is False)
check("Range: 8mm pad gets hole (== 8.0)",
      should_have_center_hole(8.0, 3.5, sr_hole) is True)

# Pad outside range falls back to universal min_hole_size (16.5)
check("Range: 20mm pad outside range uses universal min_hole_size",
      should_have_center_hole(20.0, 3.5, sr_hole) is True)
check("Range: 16mm pad outside range (< 16.5) no hole",
      should_have_center_hole(16.0, 3.5, sr_hole) is False)

# =============================================================================
print("\n--- Backward Compatibility: Missing Range Keys ---")

s_legacy = copy.deepcopy(DEFAULT_SETTINGS)
# Simulate old config without any range keys
s_legacy.pop("sizing_range_mode", None)
s_legacy.pop("sizing_ranges", None)
s_legacy.pop("engraving_settings_range_mode", None)
s_legacy.pop("engraving_settings_ranges", None)
s_legacy.pop("engraving_placement_range_mode", None)
s_legacy.pop("engraving_placement_ranges", None)

# get_sizing_for_size falls back to universal
r_legacy = get_sizing_for_size(10.0, s_legacy)
check("Legacy sizing: returns a dict (not None)", r_legacy is not None)
check("Legacy sizing: felt_offset is default", r_legacy["felt_offset"] == 0.75)
check("Legacy sizing: min_hole_size is default", r_legacy["min_hole_size"] == 16.5)

# get_engraving_settings_for_size falls back to universal
eng_legacy = get_engraving_settings_for_size(10.0, s_legacy)
check("Legacy engraving: returns a dict", eng_legacy is not None)
check("Legacy engraving: engraving_on is True", eng_legacy["engraving_on"] is True)

# get_engraving_placement_for_size falls back to universal
plc_legacy = get_engraving_placement_for_size(10.0, s_legacy)
check("Legacy placement: returns a dict", plc_legacy is not None)
check("Legacy placement: has engraving_location", "engraving_location" in plc_legacy)

# get_disc_diameter works with legacy settings
d_legacy_felt = get_disc_diameter(10.0, 'felt', s_legacy)
d_default_felt = get_disc_diameter(10.0, 'felt', make_settings())
check("Legacy: get_disc_diameter matches default", d_legacy_felt == d_default_felt)

# should_have_center_hole works with legacy settings
check("Legacy: should_have_center_hole works",
      should_have_center_hole(20.0, 3.5, s_legacy) is True)

# =============================================================================
print(f"\n{'='*60}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed} tests")
print(f"{'='*60}")
if failed:
    print("SOME TESTS FAILED")
    sys.exit(1)
else:
    print("ALL TESTS PASSED")
