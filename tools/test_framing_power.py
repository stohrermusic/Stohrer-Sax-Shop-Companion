"""
Tests for per-material framing power (Options > Machine > Framing Power).

Framing traces the outline at very low power so you can see where the cut
will land without marking the material. How low "visible but harmless" is
depends entirely on the material — 1% reads fine on card and is invisible
on dark leather, which is what drove making this per-material.

The non-negotiable: framing must ALWAYS use M4 dynamic power, whatever the
S value. M4 scales the beam with feed rate so it drops to zero when the
head decelerates or sits at a vertex while the user jogs or thinks. M3
(constant power) would keep firing the full S value during any stall and
scorch the material — and raising the power makes that worse, so the guard
matters more now than it did at 1%.

Run:
    python tools/test_framing_power.py
"""
import os
import re
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import config  # noqa: E402
from config import get_framing_power_s, GCODE_PRESET_MATERIALS, FRAMING_POWER_S_MAX  # noqa: E402
import gcode_engine as ge  # noqa: E402

# ui_dialogs resolves _() at import time, so the catalog has to be
# installed before any test imports it. main.py does this itself, but
# tests don't run in a fixed order relative to that import.
from i18n import init_translation  # noqa: E402
init_translation('en')

results = []

SQUARE = [(0, 0), (60, 0), (60, 40), (0, 40)]


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


# --------------------------------------------------------------------------
# M4 dynamic power — the guard Matt asked for explicitly
# --------------------------------------------------------------------------

def test_polygon_framing_always_uses_m4():
    """At every power level, including 0 and the ceiling."""
    for power in (0, 1, 10, 25, 60, FRAMING_POWER_S_MAX):
        lines = ge.generate_polygon_framing_gcode(SQUARE, power_s=power)
        body = '\n'.join(lines)
        assert re.search(rf'^M4 S{power}$', body, re.M), \
            f"power {power}: expected 'M4 S{power}', got:\n{body[:200]}"
        assert 'M3' not in body, (
            f"power {power}: M3 constant power would keep firing while the "
            "head sits at a vertex")


def test_bbox_framing_always_uses_m4():
    for power in (0, 10, 25, FRAMING_POWER_S_MAX):
        lines = ge.generate_framing_gcode(0, 0, 50, 50, power_s=power)
        body = '\n'.join(lines)
        assert 'M4' in body, f"power {power}: bbox framing lost M4"
        assert 'M3' not in body, f"power {power}: bbox framing used M3"


def test_power_value_reaches_the_gcode():
    """The setting has to actually change the emitted S value, or the
    dialog is decorative."""
    for power in (10, 25, 40):
        body = '\n'.join(ge.generate_polygon_framing_gcode(SQUARE, power_s=power))
        assert f'S{power}' in body, f"S{power} missing from framing G-code"


# --------------------------------------------------------------------------
# Per-material lookup
# --------------------------------------------------------------------------

def test_per_material_lookup():
    s = dict(config.DEFAULT_SETTINGS)
    s['laser_framing_power_by_material'] = {
        'felt': 10, 'card': 12, 'leather': 30, 'acrylic': 10, 'basswood': 15,
    }
    assert get_framing_power_s('leather', s) == 30
    assert get_framing_power_s('card', s) == 12
    assert get_framing_power_s('basswood', s) == 15


def test_unknown_material_falls_back_to_global():
    """'exact_size' is a Pad Maker material the per-material dict doesn't
    cover, so it must not raise or return None."""
    s = dict(config.DEFAULT_SETTINGS)
    s['laser_framing_power_s'] = 18
    s['laser_framing_power_by_material'] = {'leather': 30}
    assert get_framing_power_s('exact_size', s) == 18
    assert get_framing_power_s('felt', s) == 18
    assert get_framing_power_s('leather', s) == 30


def test_legacy_config_without_the_dict():
    """Configs written before this setting existed only have the scalar."""
    s = dict(config.DEFAULT_SETTINGS)
    s['laser_framing_power_s'] = 22
    del s['laser_framing_power_by_material']
    for mat in GCODE_PRESET_MATERIALS:
        assert get_framing_power_s(mat, s) == 22


def test_corrupt_dict_falls_back_rather_than_raising():
    s = dict(config.DEFAULT_SETTINGS)
    s['laser_framing_power_s'] = 11
    for junk in (None, [], "nope", 5):
        s['laser_framing_power_by_material'] = junk
        assert get_framing_power_s('leather', s) == 11, f"junk {junk!r} not handled"
    s['laser_framing_power_by_material'] = {'leather': 'abc'}
    assert get_framing_power_s('leather', s) == 11


def test_defaults_preserve_existing_behavior():
    """Shipping a new default power would silently change every user's
    framing brightness. Out of the box every material must still be 10."""
    s = config.DEFAULT_SETTINGS
    for mat in GCODE_PRESET_MATERIALS:
        assert get_framing_power_s(mat, s) == 10, \
            f"{mat} default drifted from the historical S10"
    assert s['laser_framing_power_s'] == 10


def test_all_preset_materials_have_an_entry():
    d = config.DEFAULT_SETTINGS['laser_framing_power_by_material']
    for mat in GCODE_PRESET_MATERIALS:
        assert mat in d, f"{mat} missing from the default framing-power map"


def test_ceiling_is_well_below_engraving_power():
    """Framing is a preview, not a cut. The ceiling has to stop someone
    typing 900 and burning the piece they're lining up."""
    assert FRAMING_POWER_S_MAX <= 200, \
        f"framing ceiling {FRAMING_POWER_S_MAX} is too close to cut power"


# --------------------------------------------------------------------------
# Wiring
# --------------------------------------------------------------------------

def test_frame_and_cut_uses_the_per_material_value():
    import inspect
    import main
    src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
    assert 'get_framing_power_s(material' in src, \
        "Frame & Cut no longer looks up framing power per material"


def test_calibration_frames_at_basswood_power():
    import inspect
    import ui_dialogs
    src = inspect.getsource(ui_dialogs.CameraCalibrationDialog)
    assert "get_framing_power_s('basswood'" in src, \
        "calibration framing should use basswood's power (the card material)"


def test_dialog_talks_in_percent_but_stores_s():
    """S values are meaningless to a user — every other power field in the
    app is a percentage. The dialog reads and writes percent; the stored
    unit stays Grbl S, because reinterpreting the existing key's 10 as
    '10%' would raise everyone's framing power tenfold."""
    import inspect
    import main
    src = inspect.getsource(main.PadSVGGeneratorApp._on_machine_framing_power)
    assert 'pct * 10' in src.replace('  ', ' '), \
        "dialog should convert entered percent back to a Grbl S value"
    assert '/ 10.0' in src, "dialog should display stored S as a percent"
    assert 'FRAMING_POWER_S_MAX / 10' in src, \
        "the entry range should be expressed in percent too"


def test_percent_round_trip_is_exact_at_the_defaults():
    """1% must survive display->edit->store as exactly S10, not S9 or S11."""
    for s_value in (0, 10, 25, 30, 100, FRAMING_POWER_S_MAX):
        pct = s_value / 10.0
        assert int(round(pct * 10)) == s_value, \
            f"S{s_value} does not round-trip through percent"


def test_menu_item_is_wired_and_gated():
    import inspect
    import main
    src = inspect.getsource(main.PadSVGGeneratorApp)
    assert '_on_machine_framing_power' in src, "no Framing Power menu command"
    assert "'framing_power'" in src, "menu index not recorded"
    # must grey out with the rest of the calibration-dependent items
    gate = re.search(r"for key in \(([^)]*)\)", src)
    assert gate and 'framing_power' in gate.group(1), \
        "Framing Power is not gated with the other Machine items"


if __name__ == '__main__':
    print("Per-material framing power")
    print("=" * 60)
    for name, fn in sorted(list(globals().items())):
        if name.startswith('test_') and callable(fn):
            check(name[5:].replace('_', ' '), fn)
    print("=" * 60)
    passed, total = sum(results), len(results)
    print(f"{passed}/{total} passed")
    sys.exit(0 if passed == total else 1)
