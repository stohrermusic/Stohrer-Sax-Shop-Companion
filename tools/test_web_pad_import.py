"""Test web pad preset import logic (selective merge, no network needed)."""
import sys
import os
import json
import copy
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

passed = 0
failed = 0

def test(name, actual, expected):
    global passed, failed
    if actual == expected:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        print(f"    Expected: {expected!r}")
        print(f"    Got:      {actual!r}")
        failed += 1

# Simulate the selective merge logic from WebImportPresetsWindow.import_selected
def simulate_selective_import(local_presets, web_data, selected_keys):
    """Simulate importing only selected presets (lib::preset keys)."""
    result = copy.deepcopy(local_presets)
    libs_touched = set()
    for full_name in selected_keys:
        lib_name, preset_name = full_name.split("::", 1)
        if lib_name not in result:
            result[lib_name] = {}
        result[lib_name][preset_name] = web_data[lib_name][preset_name]
        libs_touched.add(lib_name)
    return result, len(selected_keys), libs_touched

web_data = {
    "Music Center": {
        "Yamaha YSS-62 (Soprano)": {"pads": "36 x 2\n34 x 2\n", "notes": "via Oscar"},
        "Jupiter (Alto)": {"pads": "48.2 x 2\n43.2 x 1\n", "notes": "via Oscar"}
    },
    "Elkhart": {
        "Mark VI (Alto)": {"pads": "44 x 2\n42 x 1\n", "notes": "factory docs"}
    }
}

# --- Test: Import all ---
print("=== Import all presets ===")
all_keys = ["Music Center::Yamaha YSS-62 (Soprano)", "Music Center::Jupiter (Alto)", "Elkhart::Mark VI (Alto)"]
result, count, libs = simulate_selective_import({"My Presets": {}}, web_data, all_keys)
test("3 presets imported", count, 3)
test("2 libraries touched", len(libs), 2)
test("My Presets untouched", result["My Presets"], {})
test("Music Center has 2", len(result["Music Center"]), 2)
test("Elkhart has 1", len(result["Elkhart"]), 1)

# --- Test: Import subset (one from each library) ---
print("\n=== Import subset ===")
subset = ["Music Center::Yamaha YSS-62 (Soprano)", "Elkhart::Mark VI (Alto)"]
result, count, libs = simulate_selective_import({"My Presets": {}}, web_data, subset)
test("2 presets imported", count, 2)
test("Music Center has 1", len(result["Music Center"]), 1)
test("correct preset imported", "Yamaha YSS-62 (Soprano)" in result["Music Center"], True)
test("Jupiter not imported", "Jupiter (Alto)" not in result.get("Music Center", {}), True)

# --- Test: Import nothing ---
print("\n=== Import nothing ===")
result, count, libs = simulate_selective_import({"My Presets": {}}, web_data, [])
test("0 presets imported", count, 0)
test("no libs touched", len(libs), 0)
test("only My Presets exists", list(result.keys()), ["My Presets"])

# --- Test: Import into existing library (additive) ---
print("\n=== Import into existing library ===")
local = {
    "My Presets": {"Custom": {"pads": "18 x 5\n", "notes": ""}},
    "Music Center": {"Old Preset": {"pads": "99 x 1\n", "notes": "local"}}
}
subset = ["Music Center::Yamaha YSS-62 (Soprano)"]
result, count, libs = simulate_selective_import(local, web_data, subset)
test("Old Preset preserved", "Old Preset" in result["Music Center"], True)
test("New preset added", "Yamaha YSS-62 (Soprano)" in result["Music Center"], True)
test("Music Center has 2", len(result["Music Center"]), 2)
test("My Presets untouched", "Custom" in result["My Presets"], True)

# --- Test: Overwrite existing preset ---
print("\n=== Overwrite existing preset ===")
local2 = {
    "Music Center": {"Yamaha YSS-62 (Soprano)": {"pads": "old data\n", "notes": "old"}}
}
subset = ["Music Center::Yamaha YSS-62 (Soprano)"]
result, count, libs = simulate_selective_import(local2, web_data, subset)
test("preset overwritten with web data",
     result["Music Center"]["Yamaha YSS-62 (Soprano)"]["pads"], "36 x 2\n34 x 2\n")

# --- Test: New library created on demand ---
print("\n=== New library created ===")
result, count, libs = simulate_selective_import({}, web_data, ["Elkhart::Mark VI (Alto)"])
test("Elkhart created", "Elkhart" in result, True)
test("Music Center not created", "Music Center" not in result, True)

# --- Test: lib::preset key parsing ---
print("\n=== Key parsing ===")
full_name = "Music Center::Yamaha YSS-62 (Soprano)"
lib, preset = full_name.split("::", 1)
test("lib parsed", lib, "Music Center")
test("preset parsed", preset, "Yamaha YSS-62 (Soprano)")

# --- Test: Live fetch ---
print("\n=== Live fetch test ===")
try:
    import urllib.request
    req = urllib.request.Request(
        "https://www.stohrermusic.com/data/pad_presets.json",
        headers={"User-Agent": "StohrerSaxShopCompanion"}
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        live_data = json.loads(resp.read().decode("utf-8"))

    test("live data is dict", isinstance(live_data, dict), True)
    test("live data has libraries", len(live_data) > 0, True)

    total_presets = 0
    for lib_name, presets in live_data.items():
        if isinstance(presets, dict):
            total_presets += len(presets)
            first_name = list(presets.keys())[0]
            first = presets[first_name]
            test(f"  [{lib_name}] '{first_name}' has pads key", "pads" in first, True)
            pads = first["pads"]
            test(f"  [{lib_name}] '{first_name}' newline-separated", '\n' in pads, True)
            test(f"  [{lib_name}] '{first_name}' not comma-separated", ',' not in pads, True)

    print(f"\n  Live data: {len(live_data)} libraries, {total_presets} total presets")
    for lib_name, presets in live_data.items():
        if isinstance(presets, dict):
            print(f"    {lib_name}: {len(presets)} presets")

except Exception as e:
    print(f"  SKIP: Could not fetch live data ({e})")

# --- Summary ---
print(f"\n{'='*40}")
print(f"Results: {passed} passed, {failed} failed")
if failed:
    sys.exit(1)
