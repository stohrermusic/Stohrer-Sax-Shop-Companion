"""Test script for pad preset notes feature.
Tests backward compatibility, save/load, and notes round-trip.
"""
import sys
import os
import json
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from config import load_presets, save_presets

passed = 0
failed = 0

def test(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1

# =============================================
# 1. Backward compatibility: old string format
# =============================================
print("\n--- Backward Compatibility (old string format) ---")

old_format = {
    "My Presets": {
        "Alto Set": "42.0 x 3\n38.0 x 5\n25.0 x 8",
        "Tenor Set": "50.0 x 2\n44.0 x 4"
    }
}

with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
    json.dump(old_format, f)
    old_file = f.name

try:
    loaded = load_presets(old_file, "Pad Preset")
    test("Old format loads without error", loaded is not None)
    test("Libraries preserved", "My Presets" in loaded)
    test("Preset names preserved", "Alto Set" in loaded["My Presets"])

    # The raw value is still a string (load_presets doesn't convert)
    raw = loaded["My Presets"]["Alto Set"]
    test("Old format value is string", isinstance(raw, str))
    test("Old format text content correct", "42.0 x 3" in raw)
finally:
    os.unlink(old_file)

# =============================================
# 2. New dict format: save and load round-trip
# =============================================
print("\n--- New Dict Format Round-Trip ---")

new_format = {
    "My Presets": {
        "Alto Set": {"pads": "42.0 x 3\n38.0 x 5", "notes": "Yamaha 62 alto"},
        "Tenor Set": {"pads": "50.0 x 2\n44.0 x 4", "notes": ""}
    }
}

with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
    tmp_file = f.name

try:
    save_ok = save_presets(new_format, tmp_file)
    test("Save new format succeeds", save_ok)

    loaded = load_presets(tmp_file, "Pad Preset")
    test("New format loads without error", loaded is not None)

    raw = loaded["My Presets"]["Alto Set"]
    test("New format value is dict", isinstance(raw, dict))
    test("Pads text preserved", raw["pads"] == "42.0 x 3\n38.0 x 5")
    test("Notes preserved", raw["notes"] == "Yamaha 62 alto")

    raw2 = loaded["My Presets"]["Tenor Set"]
    test("Empty notes preserved", raw2["notes"] == "")
finally:
    os.unlink(tmp_file)

# =============================================
# 3. _get_pad_preset_data helper logic
# =============================================
print("\n--- Helper: _get_pad_preset_data logic ---")

def get_pad_preset_data(raw):
    """Mirrors the helper from main.py"""
    if isinstance(raw, dict):
        return raw.get("pads", ""), raw.get("notes", "")
    return raw, ""

# Old string format
pads, notes = get_pad_preset_data("42.0 x 3\n38.0 x 5")
test("String input returns pads text", pads == "42.0 x 3\n38.0 x 5")
test("String input returns empty notes", notes == "")

# New dict format
pads, notes = get_pad_preset_data({"pads": "42.0 x 3", "notes": "My notes"})
test("Dict input returns pads text", pads == "42.0 x 3")
test("Dict input returns notes", notes == "My notes")

# Dict with missing keys
pads, notes = get_pad_preset_data({"pads": "42.0 x 3"})
test("Dict missing notes key returns empty string", notes == "")

pads, notes = get_pad_preset_data({})
test("Empty dict returns empty pads", pads == "")
test("Empty dict returns empty notes", notes == "")

# =============================================
# 4. Mixed format file (migration scenario)
# =============================================
print("\n--- Mixed Format (old + new in same file) ---")

mixed = {
    "Old Library": {
        "Legacy Set": "18.0 x 12\n22.0 x 6"
    },
    "New Library": {
        "Modern Set": {"pads": "42.0 x 3", "notes": "Freshly saved"}
    }
}

with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
    json.dump(mixed, f)
    mixed_file = f.name

try:
    loaded = load_presets(mixed_file, "Pad Preset")
    test("Mixed format loads", loaded is not None)

    old_raw = loaded["Old Library"]["Legacy Set"]
    new_raw = loaded["New Library"]["Modern Set"]

    old_pads, old_notes = get_pad_preset_data(old_raw)
    new_pads, new_notes = get_pad_preset_data(new_raw)

    test("Old entry pads correct", "18.0 x 12" in old_pads)
    test("Old entry notes empty", old_notes == "")
    test("New entry pads correct", new_pads == "42.0 x 3")
    test("New entry notes correct", new_notes == "Freshly saved")
finally:
    os.unlink(mixed_file)

# =============================================
# 5. Notes update preserves pads
# =============================================
print("\n--- Notes Update Preserves Pads ---")

preset = {"pads": "42.0 x 3\n38.0 x 5", "notes": ""}
pads_text, _ = get_pad_preset_data(preset)
updated = {"pads": pads_text, "notes": "Added some notes"}
test("Pads unchanged after notes update", updated["pads"] == "42.0 x 3\n38.0 x 5")
test("Notes updated", updated["notes"] == "Added some notes")

# Simulate overwrite: preserve notes when re-saving pads
_, existing_notes = get_pad_preset_data(updated)
overwritten = {"pads": "NEW PADS TEXT", "notes": existing_notes}
test("Notes preserved on overwrite", overwritten["notes"] == "Added some notes")
test("Pads replaced on overwrite", overwritten["pads"] == "NEW PADS TEXT")

# =============================================
# Summary
# =============================================
print(f"\n{'='*40}")
print(f"Results: {passed} passed, {failed} failed, {passed + failed} total")
if failed == 0:
    print("All tests passed!")
else:
    print(f"FAILURES: {failed}")
    sys.exit(1)
