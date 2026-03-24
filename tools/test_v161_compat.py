"""
Test that a v1.61 config file generates SVGs and G-code without crashing.
Simulates upgrading from v1.61 to current version.
"""
import json, os, sys, copy, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from config import load_settings, SETTINGS_FILE, DEFAULT_SETTINGS

# Exact v1.61 config structure (from git show c88e282:config.py)
OLD_CONFIG = {
    "units": "in",
    "felt_offset": 0.75,
    "card_to_felt_offset": 2.0,
    "leather_wrap_multiplier": 1.0,
    "sheet_width": "13.5",
    "sheet_height": "10",
    "hole_option": "3.5mm",
    "custom_hole_size": "4.0",
    "min_hole_size": 16.5,
    "felt_thickness": 3.175,
    "felt_thickness_unit": "mm",
    "engraving_on": True,
    "show_engraving_warning": True,
    "last_output_dir": "",
    "resonance_clicks": 0,
    "compatibility_mode": False,
    "darts_enabled": True,
    "dart_threshold": 18.0,
    "dart_overwrap": 0.5,
    "dart_wrap_bonus": 0.75,
    "dart_frequency_multiplier": 1.0,
    "dart_shape_factor": 0.0,
    "dart_engraving_on": True,
    "dart_engraving_loc": {"mode": "from_outside", "value": 2.5},
    "max_fill_style": "center_out",
    "card_use_paper_size": False,
    "card_paper_size": "letter",
    "gcode_output_enabled": False,
    "gcode_settings": {
        "felt": {
            "engraving_mode": "line",
            "engraving_speed": 1200, "engraving_power": 8,
            "filled_engraving_speed": 1200, "filled_engraving_power": 8,
            "filled_line_spacing": 0.15,
            "hole_speed": 300, "hole_power": 35,
            "cut_speed": 600, "cut_power": 60,
            "kerf_width": 0.5,
            "air_assist_engraving": True, "air_assist_filled_engraving": True,
            "air_assist_hole": True, "air_assist_cut": True,
        },
        "card": {
            "engraving_mode": "line",
            "engraving_speed": 1500, "engraving_power": 10,
            "filled_engraving_speed": 1200, "filled_engraving_power": 15,
            "filled_line_spacing": 0.15,
            "hole_speed": 400, "hole_power": 22.5,
            "cut_speed": 1500, "cut_power": 50,
            "kerf_width": 0.2,
            "air_assist_engraving": True, "air_assist_filled_engraving": True,
            "air_assist_hole": True, "air_assist_cut": True,
        },
        "leather": {
            "engraving_mode": "line",
            "engraving_speed": 1500, "engraving_power": 10,
            "filled_engraving_speed": 1200, "filled_engraving_power": 15,
            "filled_line_spacing": 0.15,
            "hole_speed": 400, "hole_power": 22.5,
            "cut_speed": 900, "cut_power": 35,
            "kerf_width": 0.2,
            "air_assist_engraving": True, "air_assist_filled_engraving": True,
            "air_assist_hole": True, "air_assist_cut": True,
        },
    },
    "gcode_return_speed": 1000,
    "gcode_cut_grouping": "layer",
    "engraving_font_size": {"felt": 3.0, "card": 3.0, "leather": 3.0, "exact_size": 3.0},
    "engraving_location": {
        "felt": {"mode": "from_inside", "value": 4.0},
        "card": {"mode": "from_inside", "value": 4.0},
        "leather": {"mode": "from_outside", "value": 1.0},
        "exact_size": {"mode": "from_inside", "value": 4.0},
    },
    "layer_colors": {
        "felt_outline": "C10", "felt_center_hole": "C09", "felt_engraving": "C00",
        "card_outline": "C15", "card_center_hole": "C14", "card_engraving": "C01",
        "leather_outline": "C05", "leather_center_hole": "C03", "leather_engraving": "C02",
        "exact_size_outline": "C25", "exact_size_center_hole": "C24", "exact_size_engraving": "C06",
    },
}

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

# Write v1.61 config to disk and load through merge
backup = None
if os.path.exists(SETTINGS_FILE):
    with open(SETTINGS_FILE, 'r') as f:
        backup = f.read()

with open(SETTINGS_FILE, 'w') as f:
    json.dump(OLD_CONFIG, f, indent=2)

try:
    settings = load_settings()

    print("--- Settings Migration from v1.61 ---")
    check("Settings loaded without crash", settings is not None)
    check("units preserved", settings["units"] == "in")
    check("felt_offset preserved", settings["felt_offset"] == 0.75)
    check("engraving_location preserved", settings["engraving_location"]["felt"]["mode"] == "from_inside")
    check("dart_range_mode filled", settings.get("dart_range_mode") == "universal")
    check("sizing_range_mode filled", settings.get("sizing_range_mode") == "universal")
    check("edge_bias filled", settings.get("edge_bias") == "center")
    check("engraving_settings_range_mode filled", settings.get("engraving_settings_range_mode") == "universal")
    check("engraving_placement_range_mode filled", settings.get("engraving_placement_range_mode") == "universal")

    from svg_engine import (generate_svg, get_disc_diameter,
                            check_for_oversized_engravings, nest_pads)
    from gcode_engine import generate_gcode_from_placed

    pads = [
        {"size": 8.0, "qty": 2},
        {"size": 10.5, "qty": 3},
        {"size": 14.0, "qty": 2},
        {"size": 18.0, "qty": 1},
        {"size": 25.0, "qty": 2},
        {"size": 36.0, "qty": 1},
    ]

    width_mm = 13.5 * 25.4
    height_mm = 10.0 * 25.4
    hole_dia = 3.5

    print("\n--- Disc Diameter Calculations ---")
    for material in ["felt", "card", "leather"]:
        all_ok = True
        for pad in pads:
            d = get_disc_diameter(pad["size"], material, settings)
            if d <= 0:
                all_ok = False
        check(f"get_disc_diameter works for {material}", all_ok)

    print("\n--- Nesting ---")
    for material in ["felt", "card", "leather"]:
        placed = nest_pads(pads, material, width_mm, height_mm, settings)
        check(f"nesting places pads for {material} ({len(placed)} placed)", len(placed) > 0)

    print("\n--- SVG Generation ---")
    for material in ["felt", "card", "leather"]:
        outfile = os.path.join(tempfile.gettempdir(), f"test_v161_{material}.svg")
        try:
            generate_svg(pads, material, width_mm, height_mm, outfile, hole_dia, settings)
            exists = os.path.exists(outfile)
            size = os.path.getsize(outfile) if exists else 0
            check(f"SVG generated for {material} ({size} bytes)", exists and size > 100)
        except Exception as e:
            check(f"SVG generated for {material}", False)
            print(f"    ERROR: {e}")
        finally:
            if os.path.exists(outfile):
                os.remove(outfile)

    print("\n--- G-code Generation ---")
    for material in ["felt", "card", "leather"]:
        placed = nest_pads(pads, material, width_mm, height_mm, settings)
        outfile = os.path.join(tempfile.gettempdir(), f"test_v161_{material}.gcode")
        try:
            generate_gcode_from_placed(placed, material, width_mm, height_mm,
                                       outfile, hole_dia, settings)
            exists = os.path.exists(outfile)
            size = os.path.getsize(outfile) if exists else 0
            check(f"G-code generated for {material} ({size} bytes)", exists and size > 100)
        except Exception as e:
            check(f"G-code generated for {material}", False)
            print(f"    ERROR: {e}")
        finally:
            if os.path.exists(outfile):
                os.remove(outfile)

    print("\n--- Edge Cases ---")
    # Test with null values (corrupted config)
    with open(SETTINGS_FILE, 'w') as f:
        json.dump({"units": None, "felt_offset": None, "engraving_location": None}, f)
    s2 = load_settings()
    check("Null values rejected, units has default", s2["units"] == "in")
    check("Null values rejected, felt_offset has default", s2["felt_offset"] == 0.75)
    check("Null engraving_location uses default dict", isinstance(s2["engraving_location"], dict))

    # Test SVG generation with recovered-from-null settings
    outfile = os.path.join(tempfile.gettempdir(), "test_null_recovery.svg")
    try:
        generate_svg(pads, "felt", width_mm, height_mm, outfile, hole_dia, s2)
        check("SVG generates after null recovery", os.path.exists(outfile))
    except Exception as e:
        check("SVG generates after null recovery", False)
        print(f"    ERROR: {e}")
    finally:
        if os.path.exists(outfile):
            os.remove(outfile)

    # Test with completely empty config
    with open(SETTINGS_FILE, 'w') as f:
        json.dump({}, f)
    s3 = load_settings()
    check("Empty config loads to defaults", s3["units"] == DEFAULT_SETTINGS["units"])
    outfile = os.path.join(tempfile.gettempdir(), "test_empty_config.svg")
    try:
        generate_svg(pads, "felt", width_mm, height_mm, outfile, hole_dia, s3)
        check("SVG generates from empty config", os.path.exists(outfile))
    except Exception as e:
        check("SVG generates from empty config", False)
        print(f"    ERROR: {e}")
    finally:
        if os.path.exists(outfile):
            os.remove(outfile)

finally:
    if backup:
        with open(SETTINGS_FILE, 'w') as f:
            f.write(backup)
    elif os.path.exists(SETTINGS_FILE):
        os.remove(SETTINGS_FILE)

print(f"\n{'='*60}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed} tests")
print(f"{'='*60}")
if failed:
    print("SOME TESTS FAILED")
    sys.exit(1)
else:
    print("ALL TESTS PASSED")
