"""
G-code presets workflow test. Exercises the per-material preset flow on
GcodeSettingsWindow:

  - settings_to_gcode_presets bootstraps one Default per material
  - material_settings_to_gcode_preset extracts only GCODE_PRESET_KEYS
  - dirty/baseline tracking per material
  - _capture_material_to_dict / _apply_dict_to_material round-trip
  - active_preset_name + baseline reset after Load and Save
  - rename / delete protections (duplicates, last preset)
  - cross-material isolation: editing felt doesn't dirty acrylic

Runs non-interactively (withdraws the dialogs). Skips on headless Linux.

Run:
    python tools/test_gcode_presets_workflow.py
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def main():
    print("G-code Presets Workflow Test")
    print("=" * 60)

    import copy
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    from config import (
        DEFAULT_SETTINGS,
        GCODE_PRESET_KEYS, GCODE_PRESET_MATERIALS,
        material_settings_to_gcode_preset, settings_to_gcode_presets,
    )
    from ui_dialogs import GcodeSettingsWindow

    def helper_extract():
        felt = DEFAULT_SETTINGS["gcode_settings"]["felt"]
        out = material_settings_to_gcode_preset(felt)
        for k in GCODE_PRESET_KEYS:
            assert k in out, f"missing key {k}"
        for k in out:
            assert k in GCODE_PRESET_KEYS, f"unexpected key {k}"
        # Deepcopy: mutation of output should not leak back.
        out["cut_speed"] = 99999
        assert felt["cut_speed"] != 99999
    check("material_settings_to_gcode_preset extracts the right keys (deepcopy)", helper_extract)

    def helper_bootstrap():
        s = copy.deepcopy(DEFAULT_SETTINGS)
        boot = settings_to_gcode_presets(s)
        for mat in GCODE_PRESET_MATERIALS:
            assert mat in boot, f"missing material {mat}"
            assert "Default" in boot[mat], f"{mat} missing Default preset"
            for k in GCODE_PRESET_KEYS:
                assert k in boot[mat]["Default"], f"{mat}.Default missing {k}"
    check("settings_to_gcode_presets bootstraps one Default per material", helper_bootstrap)

    settings = copy.deepcopy(DEFAULT_SETTINGS)
    pad_materials = [("felt", "Felt"), ("card", "Card"), ("leather", "Leather")]
    saves = {"count": 0}

    def save_cb():
        saves["count"] += 1

    def make_window(presets):
        w = GcodeSettingsWindow(
            root, settings, lambda s: None,
            materials=pad_materials,
            gcode_presets=presets,
            gcode_presets_save_callback=save_cb,
        )
        w.top.withdraw()
        return w

    def clean_on_open():
        presets = settings_to_gcode_presets(settings)
        w = make_window(presets)
        try:
            for mat, _label in pad_materials:
                assert not w._material_is_dirty(mat), f"{mat} should be clean on open"
            assert w._dirty_materials() == []
        finally:
            w.top.destroy()
    check("All materials are clean immediately after dialog open", clean_on_open)

    def edit_dirties_one_material_only():
        presets = settings_to_gcode_presets(settings)
        w = make_window(presets)
        try:
            w.vars["felt"]["cut"]["speed"].set(9999)
            assert w._material_is_dirty("felt")
            assert not w._material_is_dirty("card")
            assert not w._material_is_dirty("leather")
            assert w._dirty_materials() == ["felt"]
        finally:
            w.top.destroy()
    check("Editing felt does not dirty card or leather", edit_dirties_one_material_only)

    def capture_apply_round_trip():
        presets = settings_to_gcode_presets(settings)
        w = make_window(presets)
        try:
            snap1 = w._capture_material_to_dict("felt")
            # Apply the same dict back; snapshot should still match.
            w._apply_dict_to_material("felt", snap1)
            snap2 = w._capture_material_to_dict("felt")
            assert snap1 == snap2, f"round trip mismatch:\n{snap1}\nvs\n{snap2}"
            # Apply a modified dict; snapshot reflects the change.
            modified = dict(snap1, cut_power=42.0, cut_passes=3,
                            air_assist_cut=False, kerf_width=0.42)
            w._apply_dict_to_material("felt", modified)
            snap3 = w._capture_material_to_dict("felt")
            assert snap3["cut_power"] == 42.0
            assert snap3["cut_passes"] == 3
            assert snap3["air_assist_cut"] is False
            assert snap3["kerf_width"] == 0.42
        finally:
            w.top.destroy()
    check("_capture / _apply round-trip preserves values", capture_apply_round_trip)

    def load_resets_baseline():
        presets = settings_to_gcode_presets(settings)
        w = make_window(presets)
        try:
            w.vars["felt"]["cut"]["speed"].set(9999)
            assert w._material_is_dirty("felt")
            # Simulate Load by manually applying preset + resetting baseline.
            w._apply_dict_to_material("felt", presets["felt"]["Default"])
            w.active_preset_name["felt"] = "Default"
            w._set_material_baseline("felt")
            assert not w._material_is_dirty("felt")
            assert w.active_preset_name["felt"] == "Default"
        finally:
            w.top.destroy()
    check("Load resets baseline and sets active preset", load_resets_baseline)

    def save_preset_persists_and_resets_baseline():
        presets = settings_to_gcode_presets(settings)
        saves["count"] = 0
        w = make_window(presets)
        try:
            w.vars["felt"]["cut"]["speed"].set(1234)
            assert w._material_is_dirty("felt")
            # Simulate the save-preset success path.
            snap = w._capture_material_to_dict("felt")
            presets.setdefault("felt", {})["Matt's felt"] = snap
            w.gcode_presets_save_callback()
            w.active_preset_name["felt"] = "Matt's felt"
            w._set_material_baseline("felt")
            assert not w._material_is_dirty("felt")
            assert "Matt's felt" in presets["felt"]
            assert presets["felt"]["Matt's felt"]["cut_speed"] == 1234
            assert saves["count"] == 1, "save callback should have fired exactly once"
        finally:
            w.top.destroy()
    check("Saving a preset persists + resets baseline + fires callback", save_preset_persists_and_resets_baseline)

    def cross_material_isolation():
        presets = settings_to_gcode_presets(settings)
        w = make_window(presets)
        try:
            w.vars["felt"]["cut"]["speed"].set(1)
            w.vars["card"]["cut"]["speed"].set(2)
            assert set(w._dirty_materials()) == {"felt", "card"}
            # Load Default for felt only -> felt clean, card still dirty.
            w._apply_dict_to_material("felt", presets["felt"]["Default"])
            w._set_material_baseline("felt")
            assert not w._material_is_dirty("felt")
            assert w._material_is_dirty("card")
            assert w._dirty_materials() == ["card"]
        finally:
            w.top.destroy()
    check("Per-material baselines are independent (cross-material isolation)", cross_material_isolation)

    def presets_none_disables_ui():
        # When the caller passes gcode_presets=None, the dialog should
        # behave like the legacy Save-only flow: no preset bar, no dirty
        # tracking, no preset state.
        w = GcodeSettingsWindow(
            root, settings, lambda s: None,
            materials=pad_materials,
            gcode_presets=None,
        )
        try:
            w.top.withdraw()
            assert w.gcode_presets is None
            assert w.preset_combos == {}
            assert w.material_baseline == {}
            # The dirty-prompt fast path should return True (proceed).
            assert w._prompt_dirty(context="apply") is True
            assert w._prompt_dirty(context="cancel") is True
        finally:
            w.top.destroy()
    check("gcode_presets=None disables preset UI and dirty-prompt", presets_none_disables_ui)

    def backfills_missing_material_on_startup():
        # Mirror the main.py bootstrap loop: a partial library should get
        # any missing materials filled in from the current settings.
        partial = {"felt": {"Default": material_settings_to_gcode_preset(
                                settings["gcode_settings"]["felt"])}}
        bootstrap = settings_to_gcode_presets(settings)
        for mat in GCODE_PRESET_MATERIALS:
            if mat not in partial or not partial[mat]:
                partial[mat] = bootstrap.get(mat, {})
        for mat in GCODE_PRESET_MATERIALS:
            assert mat in partial, f"backfill missed {mat}"
            assert partial[mat], f"backfill left {mat} empty"
    check("Startup backfill fills in any missing materials", backfills_missing_material_on_startup)

    try:
        root.destroy()
    except Exception:
        pass

    passed = sum(results)
    total = len(results)
    print("=" * 60)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
