"""
Test suite for config.py logic.

Tests settings merge, save/load round-trip, tone profiles, migration helpers,
and audio device filtering.

Usage:
    python tools/test_config.py

No GUI windows are opened. All output goes to stdout.
"""

import copy
import json
import os
import sys
import tempfile
import shutil
import types

# Add parent directory to path so we can import project modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Mock tkinter before importing config
_fake_tk = types.ModuleType("tkinter")
_fake_mb = types.ModuleType("tkinter.messagebox")
_fake_mb.showerror = lambda *a, **kw: None
_fake_mb.showinfo = lambda *a, **kw: None
_fake_mb.showwarning = lambda *a, **kw: None
_fake_tk.messagebox = _fake_mb
sys.modules.setdefault("tkinter", _fake_tk)
sys.modules.setdefault("tkinter.messagebox", _fake_mb)

import config
from config import (
    DEFAULT_SETTINGS,
    load_settings,
    save_settings,
    find_config_files_in_directory,
    load_presets,
    save_presets,
    get_input_devices,
)

# Tone profile functions live in toner_engine
try:
    from toner_engine import (
        load_tone_presets,
        save_tone_presets,
        DEFAULT_LIBRARY,
    )
    HAS_TONER = True
except ImportError:
    HAS_TONER = False

# ============================================================================
# TEST HARNESS
# ============================================================================

_results = []

def run_test(name, fn):
    """Run a test function and record PASS/FAIL."""
    try:
        fn()
        _results.append(("PASS", name, None))
        print(f"  PASS  {name}")
    except AssertionError as e:
        _results.append(("FAIL", name, str(e)))
        print(f"  FAIL  {name} -- {e}")
    except Exception as e:
        _results.append(("ERROR", name, str(e)))
        print(f"  ERROR {name} -- {type(e).__name__}: {e}")


def assert_true(condition, msg=""):
    if not condition:
        raise AssertionError(msg or "Expected True but got False")

def assert_false(condition, msg=""):
    if condition:
        raise AssertionError(msg or "Expected False but got True")

def assert_equal(a, b, msg=""):
    if a != b:
        raise AssertionError(msg or f"Expected {a!r} == {b!r}")

def assert_in(item, collection, msg=""):
    if item not in collection:
        raise AssertionError(msg or f"Expected {item!r} in {collection!r}")


# ============================================================================
# HELPERS
# ============================================================================

def _save_json(path, data):
    """Write JSON to a file."""
    with open(path, 'w') as f:
        json.dump(data, f, indent=2)


def _load_with_file(filepath):
    """Call load_settings() with SETTINGS_FILE temporarily pointed at filepath."""
    orig = config.SETTINGS_FILE
    config.SETTINGS_FILE = filepath
    try:
        return load_settings()
    finally:
        config.SETTINGS_FILE = orig


def _save_with_file(settings, filepath):
    """Call save_settings() with SETTINGS_FILE temporarily pointed at filepath."""
    orig = config.SETTINGS_FILE
    config.SETTINGS_FILE = filepath
    try:
        save_settings(settings)
    finally:
        config.SETTINGS_FILE = orig


# ============================================================================
# TESTS: Settings merge (load_settings)
# ============================================================================

def test_new_top_level_keys_filled():
    """New top-level keys in DEFAULT_SETTINGS are filled in when loading old config."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        # Save a minimal config missing many keys
        _save_json(path, {"units": "mm", "felt_offset": 1.0})
        result = _load_with_file(path)
        # User values preserved
        assert_equal(result["units"], "mm")
        assert_equal(result["felt_offset"], 1.0)
        # New keys filled from defaults
        assert_equal(result["engraving_on"], DEFAULT_SETTINGS["engraving_on"])
        assert_equal(result["edge_bias"], DEFAULT_SETTINGS["edge_bias"])
        assert_in("gcode_settings", result)
        assert_in("tuner_settings", result)
    finally:
        shutil.rmtree(tmpdir)


def test_nested_dict_keys_merged():
    """Nested dict keys (like gcode_settings.felt) get merged correctly."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        # Save config with partial gcode_settings.felt (missing some keys)
        _save_json(path, {
            "gcode_settings": {
                "felt": {
                    "cut_speed": 999,
                    "cut_power": 88,
                    # Missing: engraving_speed, hole_speed, etc.
                }
            }
        })
        result = _load_with_file(path)
        felt = result["gcode_settings"]["felt"]
        # User values preserved
        assert_equal(felt["cut_speed"], 999)
        assert_equal(felt["cut_power"], 88)
        # Missing keys filled from defaults
        assert_equal(felt["engraving_speed"], DEFAULT_SETTINGS["gcode_settings"]["felt"]["engraving_speed"])
        assert_equal(felt["hole_speed"], DEFAULT_SETTINGS["gcode_settings"]["felt"]["hole_speed"])
        assert_equal(felt["kerf_width"], DEFAULT_SETTINGS["gcode_settings"]["felt"]["kerf_width"])
    finally:
        shutil.rmtree(tmpdir)


def test_user_values_preserved():
    """User values are not overwritten by defaults."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        custom = copy.deepcopy(DEFAULT_SETTINGS)
        custom["units"] = "mm"
        custom["felt_offset"] = 2.5
        custom["edge_bias"] = "ne"
        custom["gcode_settings"]["felt"]["cut_speed"] = 1234
        custom["tuner_settings"]["reference_pitch"] = 442.0
        custom["key_layout"]["show_B"] = False
        _save_json(path, custom)
        result = _load_with_file(path)
        assert_equal(result["units"], "mm")
        assert_equal(result["felt_offset"], 2.5)
        assert_equal(result["edge_bias"], "ne")
        assert_equal(result["gcode_settings"]["felt"]["cut_speed"], 1234)
        assert_equal(result["tuner_settings"]["reference_pitch"], 442.0)
        assert_equal(result["key_layout"]["show_B"], False)
    finally:
        shutil.rmtree(tmpdir)


def test_deepcopy_isolation():
    """Mutating a loaded settings dict does NOT corrupt DEFAULT_SETTINGS."""
    # Take a snapshot of defaults before
    orig_units = DEFAULT_SETTINGS["units"]
    orig_felt_speed = DEFAULT_SETTINGS["gcode_settings"]["felt"]["cut_speed"]

    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        _save_json(path, {"units": "in"})
        result = _load_with_file(path)

        # Mutate the loaded result
        result["units"] = "CORRUPTED"
        result["gcode_settings"]["felt"]["cut_speed"] = -999
        result["tuner_settings"]["reference_pitch"] = -1

        # DEFAULT_SETTINGS must be unchanged
        assert_equal(DEFAULT_SETTINGS["units"], orig_units)
        assert_equal(DEFAULT_SETTINGS["gcode_settings"]["felt"]["cut_speed"], orig_felt_speed)
    finally:
        shutil.rmtree(tmpdir)


def test_new_nested_subkey_filled():
    """A new sub-key added to a nested dict default gets filled in."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        # Save config with tuner_settings missing some keys
        _save_json(path, {
            "tuner_settings": {
                "stripe_color": "#FF0000",
                "reference_pitch": 441.0,
                # Missing: fps, ring_brightness, etc.
            }
        })
        result = _load_with_file(path)
        ts = result["tuner_settings"]
        assert_equal(ts["stripe_color"], "#FF0000")
        assert_equal(ts["reference_pitch"], 441.0)
        # Missing keys filled
        assert_equal(ts["fps"], DEFAULT_SETTINGS["tuner_settings"]["fps"])
        assert_equal(ts["ring_brightness"], DEFAULT_SETTINGS["tuner_settings"]["ring_brightness"])
    finally:
        shutil.rmtree(tmpdir)


def test_visible_tabs_merge():
    """visible_tabs (nested dict of booleans) merges correctly."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        # Old config with only some tabs
        _save_json(path, {
            "visible_tabs": {
                "Key Height Library": False,
                "Serial Lookup": True,
            }
        })
        result = _load_with_file(path)
        vt = result["visible_tabs"]
        assert_equal(vt["Key Height Library"], False)
        assert_equal(vt["Serial Lookup"], True)
        # New keys filled from defaults
        assert_equal(vt["Screw Specs"], DEFAULT_SETTINGS["visible_tabs"]["Screw Specs"])
        assert_equal(vt["Tuner"], DEFAULT_SETTINGS["visible_tabs"]["Tuner"])
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# TESTS: Settings save/load round-trip
# ============================================================================

def test_save_load_roundtrip():
    """Save settings to a temp file, load them back, verify match."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        original = copy.deepcopy(DEFAULT_SETTINGS)
        original["units"] = "mm"
        original["felt_offset"] = 1.23
        original["gcode_settings"]["felt"]["cut_speed"] = 777
        original["tuner_settings"]["reference_pitch"] = 443.0

        _save_with_file(original, path)
        loaded = _load_with_file(path)

        assert_equal(loaded["units"], "mm")
        assert_equal(loaded["felt_offset"], 1.23)
        assert_equal(loaded["gcode_settings"]["felt"]["cut_speed"], 777)
        assert_equal(loaded["tuner_settings"]["reference_pitch"], 443.0)
    finally:
        shutil.rmtree(tmpdir)


def test_load_missing_file_returns_defaults():
    """Loading from a non-existent file returns defaults."""
    result = _load_with_file("/nonexistent/path/settings.json")
    assert_equal(result["units"], DEFAULT_SETTINGS["units"])
    assert_equal(result["felt_offset"], DEFAULT_SETTINGS["felt_offset"])


def test_load_corrupt_json_returns_defaults():
    """Loading corrupt JSON returns defaults gracefully."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        with open(path, 'w') as f:
            f.write("{this is not valid json!!!")
        result = _load_with_file(path)
        assert_equal(result["units"], DEFAULT_SETTINGS["units"])
    finally:
        shutil.rmtree(tmpdir)


def test_load_empty_file_returns_defaults():
    """Loading an empty file returns defaults."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        with open(path, 'w') as f:
            f.write("")
        result = _load_with_file(path)
        assert_equal(result["units"], DEFAULT_SETTINGS["units"])
    finally:
        shutil.rmtree(tmpdir)


def test_roundtrip_all_defaults():
    """Full DEFAULT_SETTINGS survives a save/load round trip."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        original = copy.deepcopy(DEFAULT_SETTINGS)
        _save_with_file(original, path)
        loaded = _load_with_file(path)

        # Check a sampling of keys across all levels
        assert_equal(loaded["units"], original["units"])
        assert_equal(loaded["gcode_settings"]["felt"]["cut_speed"],
                     original["gcode_settings"]["felt"]["cut_speed"])
        assert_equal(loaded["gcode_settings"]["acrylic"]["engraving_mode"],
                     original["gcode_settings"]["acrylic"]["engraving_mode"])
        assert_equal(loaded["tuner_settings"]["stripe_color"],
                     original["tuner_settings"]["stripe_color"])
        assert_equal(loaded["toner_settings"]["sax_type"],
                     original["toner_settings"]["sax_type"])
        assert_equal(loaded["key_layout"]["show_B"], original["key_layout"]["show_B"])
        assert_equal(loaded["layer_colors"]["felt_outline"],
                     original["layer_colors"]["felt_outline"])
        assert_equal(loaded["visible_tabs"]["Tuner"], original["visible_tabs"]["Tuner"])
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# TESTS: Tone profiles load/save (toner_engine)
# ============================================================================

def test_tone_profiles_roundtrip():
    """Tone profiles save/load round-trip with nested library format."""
    if not HAS_TONER:
        raise AssertionError("SKIP: toner_engine not available")
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "tone_profiles.json")
        profiles = {
            "My Library": {
                "Test Horn": {
                    "sessions": [{"date": "2026-01-01", "captures": []}],
                    "setup": {"horn": "Selmer", "mpc": "S80"}
                }
            },
            "Another Library": {
                "Other Horn": {
                    "sessions": [],
                    "setup": {"horn": "Yamaha"}
                }
            }
        }
        ok = save_tone_presets(profiles, path)
        assert_true(ok, "save_tone_presets should return True")
        loaded = load_tone_presets(path)
        assert_in("My Library", loaded)
        assert_in("Another Library", loaded)
        assert_in("Test Horn", loaded["My Library"])
        assert_equal(loaded["My Library"]["Test Horn"]["setup"]["horn"], "Selmer")
    finally:
        shutil.rmtree(tmpdir)


def test_tone_profiles_missing_file():
    """Loading tone profiles from non-existent file returns default library."""
    if not HAS_TONER:
        raise AssertionError("SKIP: toner_engine not available")
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "nonexistent.json")
        result = load_tone_presets(path)
        assert_in(DEFAULT_LIBRARY, result)
        assert_equal(result[DEFAULT_LIBRARY], {})
    finally:
        shutil.rmtree(tmpdir)


def test_tone_profiles_corrupt_json():
    """Loading corrupt JSON tone profiles returns default library."""
    if not HAS_TONER:
        raise AssertionError("SKIP: toner_engine not available")
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "corrupt.json")
        with open(path, 'w') as f:
            f.write("not json {{{")
        result = load_tone_presets(path)
        assert_in(DEFAULT_LIBRARY, result)
    finally:
        shutil.rmtree(tmpdir)


def test_tone_profiles_flat_migration():
    """Legacy flat profile format (profiles at top level) gets migrated."""
    if not HAS_TONER:
        raise AssertionError("SKIP: toner_engine not available")
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "flat_profiles.json")
        # Flat format: profile names at top with 'sessions' key
        flat = {
            "My Selmer": {
                "sessions": [{"date": "2026-01-01", "captures": []}],
                "setup": {"horn": "Selmer SBA"}
            },
            "My Yamaha": {
                "sessions": [],
                "setup": {"horn": "Yamaha 62"}
            }
        }
        _save_json(path, flat)
        loaded = load_tone_presets(path)
        # Should be wrapped in DEFAULT_LIBRARY
        assert_in(DEFAULT_LIBRARY, loaded)
        assert_in("My Selmer", loaded[DEFAULT_LIBRARY])
        assert_in("My Yamaha", loaded[DEFAULT_LIBRARY])
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# TESTS: Migration helpers
# ============================================================================

def test_find_config_files_includes_tone_profiles():
    """find_config_files_in_directory() includes tone_profiles.json."""
    tmpdir = tempfile.mkdtemp()
    try:
        # Create all expected config files
        for fn in ["app_settings.json", "pad_presets.json", "key_height_library.json",
                    "screw_specs.json", "tone_profiles.json"]:
            with open(os.path.join(tmpdir, fn), 'w') as f:
                f.write("{}")
        found = find_config_files_in_directory(tmpdir)
        assert_in("tone_profiles.json", found)
        assert_in("app_settings.json", found)
        assert_in("screw_specs.json", found)
        assert_equal(len(found), 5)
    finally:
        shutil.rmtree(tmpdir)


def test_find_config_files_empty_dir():
    """find_config_files_in_directory() returns empty list for empty dir."""
    tmpdir = tempfile.mkdtemp()
    try:
        found = find_config_files_in_directory(tmpdir)
        assert_equal(found, [])
    finally:
        shutil.rmtree(tmpdir)


def test_find_config_files_partial():
    """find_config_files_in_directory() returns only files that exist."""
    tmpdir = tempfile.mkdtemp()
    try:
        with open(os.path.join(tmpdir, "app_settings.json"), 'w') as f:
            f.write("{}")
        with open(os.path.join(tmpdir, "tone_profiles.json"), 'w') as f:
            f.write("{}")
        found = find_config_files_in_directory(tmpdir)
        assert_equal(len(found), 2)
        assert_in("app_settings.json", found)
        assert_in("tone_profiles.json", found)
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# TESTS: Presets load/save
# ============================================================================

def test_presets_roundtrip():
    """Preset save/load round-trip with nested library format."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "presets.json")
        data = {
            "My Presets": {
                "Alto set": "16.5 17 18 20 22 24 25",
                "Tenor set": "20 22 24 25 26 28 30"
            }
        }
        ok = save_presets(data, path)
        assert_true(ok, "save_presets should return True")
        loaded = load_presets(path)
        assert_in("My Presets", loaded)
        assert_equal(loaded["My Presets"]["Alto set"], "16.5 17 18 20 22 24 25")
    finally:
        shutil.rmtree(tmpdir)


def test_presets_missing_file():
    """Loading presets from non-existent file returns empty dict."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "nonexistent.json")
        loaded = load_presets(path)
        assert_equal(loaded, {})
    finally:
        shutil.rmtree(tmpdir)


def test_presets_corrupt_json():
    """Loading corrupt preset JSON returns empty dict."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "corrupt.json")
        with open(path, 'w') as f:
            f.write("{{{{broken")
        loaded = load_presets(path)
        assert_equal(loaded, {})
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# TESTS: get_input_devices filtering
# ============================================================================

def test_input_devices_filters_bluetooth():
    """get_input_devices filters out Bluetooth and sound mapper devices."""
    # We mock sounddevice to control exactly what devices are "available"
    fake_sd = types.ModuleType("sounddevice")
    fake_devices = [
        {"name": "USB Microphone", "max_input_channels": 2, "default_samplerate": 44100},
        {"name": "Bluetooth Hands-Free Audio", "max_input_channels": 1, "default_samplerate": 16000},
        {"name": "BthHFEnum Headset", "max_input_channels": 1, "default_samplerate": 16000},
        {"name": "Primary Sound Capture Driver", "max_input_channels": 2, "default_samplerate": 44100},
        {"name": "Sound Mapper - Input", "max_input_channels": 2, "default_samplerate": 44100},
        {"name": "AT2020 USB Condenser", "max_input_channels": 2, "default_samplerate": 48000},
        {"name": "Output Only Device", "max_input_channels": 0, "default_samplerate": 44100},
        {"name": "Cheap Mic Low Rate", "max_input_channels": 1, "default_samplerate": 22050},
    ]
    fake_sd.query_devices = lambda: fake_devices

    # Temporarily inject our fake module
    orig_sd = sys.modules.get("sounddevice")
    sys.modules["sounddevice"] = fake_sd
    try:
        devices = get_input_devices()
        names = [name for _, name in devices]
        # Should include these
        assert_in("USB Microphone", names, "USB Microphone should be included")
        assert_in("AT2020 USB Condenser", names, "AT2020 should be included")
        # Should filter these out
        for bad_name in ["Bluetooth Hands-Free Audio", "BthHFEnum Headset",
                         "Primary Sound Capture Driver", "Sound Mapper - Input",
                         "Output Only Device", "Cheap Mic Low Rate"]:
            assert_true(bad_name not in names, f"{bad_name} should be filtered out")
    finally:
        if orig_sd is not None:
            sys.modules["sounddevice"] = orig_sd
        else:
            del sys.modules["sounddevice"]


def test_input_devices_no_sounddevice():
    """get_input_devices returns empty list when sounddevice is unavailable."""
    # Temporarily remove sounddevice
    orig_sd = sys.modules.get("sounddevice")
    sys.modules["sounddevice"] = None  # Forces ImportError on `import sounddevice`
    try:
        devices = get_input_devices()
        assert_equal(devices, [])
    finally:
        if orig_sd is not None:
            sys.modules["sounddevice"] = orig_sd
        else:
            del sys.modules["sounddevice"]


def test_input_devices_deduplication():
    """get_input_devices deduplicates devices with similar names."""
    fake_sd = types.ModuleType("sounddevice")
    # First 30 chars must match for dedup: "Realtek High Definition Audio " (30 chars)
    fake_devices = [
        {"name": "Realtek High Definition Audio (WASAPI)", "max_input_channels": 2, "default_samplerate": 44100},
        {"name": "Realtek High Definition Audio (WDM)", "max_input_channels": 2, "default_samplerate": 44100},
    ]
    fake_sd.query_devices = lambda: fake_devices

    orig_sd = sys.modules.get("sounddevice")
    sys.modules["sounddevice"] = fake_sd
    try:
        devices = get_input_devices()
        # Should deduplicate — only one USB Microphone
        assert_equal(len(devices), 1)
    finally:
        if orig_sd is not None:
            sys.modules["sounddevice"] = orig_sd
        else:
            del sys.modules["sounddevice"]


# ============================================================================
# TESTS: Edge cases for settings merge
# ============================================================================

def test_deeply_nested_engraving_location_merge():
    """engraving_location (dict of dicts) merges correctly."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        _save_json(path, {
            "engraving_location": {
                "felt": {"mode": "from_outside", "value": 2.0},
                # card, leather, exact_size missing
            }
        })
        result = _load_with_file(path)
        el = result["engraving_location"]
        # User value preserved
        assert_equal(el["felt"]["mode"], "from_outside")
        assert_equal(el["felt"]["value"], 2.0)
        # Defaults filled for missing materials
        assert_equal(el["card"]["mode"], DEFAULT_SETTINGS["engraving_location"]["card"]["mode"])
        assert_equal(el["leather"]["value"], DEFAULT_SETTINGS["engraving_location"]["leather"]["value"])
    finally:
        shutil.rmtree(tmpdir)


def test_gauge_bias_merge():
    """toner_settings.gauge_bias (nested dict) merges correctly."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        _save_json(path, {
            "toner_settings": {
                "gauge_bias": {"richness": -3},
                # Missing: warmth — should be filled from defaults
            }
        })
        result = _load_with_file(path)
        gb = result["toner_settings"]["gauge_bias"]
        assert_equal(gb["richness"], -3)
        # Missing keys filled from defaults
        assert_equal(gb["warmth"], DEFAULT_SETTINGS["toner_settings"]["gauge_bias"]["warmth"])
    finally:
        shutil.rmtree(tmpdir)


def test_load_settings_returns_copy_not_reference():
    """Two consecutive loads return independent objects."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        _save_json(path, {"units": "in"})
        result1 = _load_with_file(path)
        result2 = _load_with_file(path)
        result1["units"] = "CHANGED"
        assert_equal(result2["units"], "in", "Second load should be independent of first")
    finally:
        shutil.rmtree(tmpdir)


def test_extra_keys_in_saved_file_ignored():
    """Keys in the saved file that are NOT in DEFAULT_SETTINGS are dropped."""
    tmpdir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmpdir, "settings.json")
        _save_json(path, {
            "units": "mm",
            "some_future_key": "should be dropped",
            "another_unknown": 42
        })
        result = _load_with_file(path)
        assert_equal(result["units"], "mm")
        assert_true("some_future_key" not in result, "Unknown keys should be dropped")
        assert_true("another_unknown" not in result, "Unknown keys should be dropped")
    finally:
        shutil.rmtree(tmpdir)


# ============================================================================
# MAIN
# ============================================================================

def main():
    global _results
    _results = []

    sections = [
        ("Settings Merge (load_settings)", [
            ("New top-level keys filled from defaults", test_new_top_level_keys_filled),
            ("Nested dict keys merged correctly", test_nested_dict_keys_merged),
            ("User values preserved (not overwritten)", test_user_values_preserved),
            ("Deepcopy: mutations don't corrupt DEFAULT_SETTINGS", test_deepcopy_isolation),
            ("New nested sub-keys filled", test_new_nested_subkey_filled),
            ("visible_tabs merge", test_visible_tabs_merge),
            ("Deeply nested engraving_location merge", test_deeply_nested_engraving_location_merge),
            ("gauge_bias nested merge", test_gauge_bias_merge),
            ("Extra keys in file are dropped", test_extra_keys_in_saved_file_ignored),
        ]),
        ("Settings Save/Load Round-Trip", [
            ("Save and load round-trip", test_save_load_roundtrip),
            ("Missing file returns defaults", test_load_missing_file_returns_defaults),
            ("Corrupt JSON returns defaults", test_load_corrupt_json_returns_defaults),
            ("Empty file returns defaults", test_load_empty_file_returns_defaults),
            ("Full defaults round-trip", test_roundtrip_all_defaults),
            ("Two loads return independent objects", test_load_settings_returns_copy_not_reference),
        ]),
        ("Tone Profiles (toner_engine)", [
            ("Tone profiles round-trip", test_tone_profiles_roundtrip),
            ("Missing file returns default library", test_tone_profiles_missing_file),
            ("Corrupt JSON returns default library", test_tone_profiles_corrupt_json),
            ("Flat format migration", test_tone_profiles_flat_migration),
        ]),
        ("Migration Helpers", [
            ("find_config_files includes tone_profiles.json", test_find_config_files_includes_tone_profiles),
            ("find_config_files on empty dir", test_find_config_files_empty_dir),
            ("find_config_files partial", test_find_config_files_partial),
        ]),
        ("Presets Load/Save", [
            ("Presets round-trip", test_presets_roundtrip),
            ("Presets missing file", test_presets_missing_file),
            ("Presets corrupt JSON", test_presets_corrupt_json),
        ]),
        ("get_input_devices", [
            ("Filters Bluetooth and sound mapper", test_input_devices_filters_bluetooth),
            ("No sounddevice returns empty list", test_input_devices_no_sounddevice),
            ("Deduplication of similar device names", test_input_devices_deduplication),
        ]),
    ]

    for section_name, tests in sections:
        print(f"\n--- {section_name} ---")
        for test_name, test_fn in tests:
            run_test(test_name, test_fn)

    # Summary
    print("\n" + "=" * 70)
    passed = sum(1 for r in _results if r[0] == "PASS")
    failed = sum(1 for r in _results if r[0] == "FAIL")
    errors = sum(1 for r in _results if r[0] == "ERROR")
    total = len(_results)

    print(f"Results: {passed} passed, {failed} failed, {errors} errors out of {total} tests")

    if failed > 0 or errors > 0:
        print("\nFailed/Error tests:")
        for status, name, msg in _results:
            if status in ("FAIL", "ERROR"):
                print(f"  {status}: {name}")
                if msg:
                    print(f"         {msg}")

    print("=" * 70)
    return 1 if (failed + errors) > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
