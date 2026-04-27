"""
Test script for tuner tab updates:
  - Per-ring brightness slider (blending uniform vs per-ring)
  - Overall brightness slider
  - Faceplate color
  - "Experimental" branding (not "Stohrer")
  - Tab name "Tuner" (not "Strobe Tuner")
  - Simple mode removed
  - VU meter in strobe mode
  - Settings save/restore with new keys
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

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


# ============================================================
# 1. Config defaults
# ============================================================
print("\n--- Config Defaults ---")
from config import DEFAULT_SETTINGS

ts = DEFAULT_SETTINGS["tuner_settings"]
test("ring_brightness default is 100", ts.get("ring_brightness") == 100)
test("overall_brightness default is 80", ts.get("overall_brightness") == 80)
test("faceplate_color default is #1A1A1A", ts.get("faceplate_color") == "#1A1A1A")
test("no tuner_mode in defaults (simple mode removed)", "tuner_mode" not in ts)
test("stripe_color default is #00FF00", ts.get("stripe_color") == "#00FF00")
test("fps default is '60'", ts.get("fps") == "60")

# ============================================================
# 2. Import checks
# ============================================================
print("\n--- Import Checks ---")
from tuner_tab import (
    StrobeWheel, TunerTabMixin, _scale_color, _build_ref_notes,
    TRANSPOSITION_KEYS, DEFAULT_FACEPLATE,
    NUM_RINGS, RING_SEGMENTS,
)

test("DEFAULT_FACEPLATE is #1A1A1A", DEFAULT_FACEPLATE == "#1A1A1A")
test("TRANSPOSITION_KEYS has 4 entries", len(TRANSPOSITION_KEYS) == 4)
test("NUM_RINGS is 7", NUM_RINGS == 7)
test("RING_SEGMENTS length matches NUM_RINGS", len(RING_SEGMENTS) == NUM_RINGS)

# ============================================================
# 3. _scale_color helper
# ============================================================
print("\n--- _scale_color ---")
test("full brightness preserves color", _scale_color("#FF0000", 1.0) == "#ff0000")
test("zero brightness is black", _scale_color("#FF8800", 0.0) == "#000000")
test("50% scales correctly", _scale_color("#FF0000", 0.5) == "#7f0000")
test("handles lowercase", _scale_color("#00ff00", 0.5) == "#007f00")

# ============================================================
# 4. _build_ref_notes
# ============================================================
print("\n--- _build_ref_notes ---")
notes = _build_ref_notes(440.0)
test("ref notes list is 48 entries (4 octaves x 12)", len(notes) == 48)
a4 = [f for n, f in notes if n == "A4"]
test("A4 is 440.0 Hz", len(a4) == 1 and abs(a4[0] - 440.0) < 0.01)
a3 = [f for n, f in notes if n == "A3"]
test("A3 is 220.0 Hz", len(a3) == 1 and abs(a3[0] - 220.0) < 0.01)

# ============================================================
# 5. StrobeWheel.update with ring brightness parameters
# ============================================================
print("\n--- StrobeWheel.update signature ---")
import inspect
sig = inspect.signature(StrobeWheel.update)
params = list(sig.parameters.keys())
test("update has ring_brightness_pct param", "ring_brightness_pct" in params)
test("update has overall_brightness_pct param", "overall_brightness_pct" in params)
test("ring_brightness_pct default is 100",
     sig.parameters["ring_brightness_pct"].default == 100)
test("overall_brightness_pct default is 80",
     sig.parameters["overall_brightness_pct"].default == 80)

# ============================================================
# 6. Simple mode fully removed
# ============================================================
print("\n--- Simple Mode Removed ---")
test("no _tuner_on_mode_changed", not hasattr(TunerTabMixin, '_tuner_on_mode_changed'))
test("no _tuner_build_simple", not hasattr(TunerTabMixin, '_tuner_build_simple'))
test("no _simple_update", not hasattr(TunerTabMixin, '_simple_update'))
test("no _simple_text_color", not hasattr(TunerTabMixin, '_simple_text_color'))
test("no _tuner_rebuild_display", not hasattr(TunerTabMixin, '_tuner_rebuild_display'))

with open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tuner_tab.py"), encoding="utf-8") as f:
    tab_source = f.read()

test("no 'simple' mode references in source",
     "_simple_" not in tab_source and "tuner_mode" not in tab_source)

# ============================================================
# 7. VU meter method exists
# ============================================================
print("\n--- VU Meter ---")
test("has _vu_update method", hasattr(TunerTabMixin, '_vu_update'))

# ============================================================
# 8. Branding check
# ============================================================
print("\n--- Branding ---")
test("'Stohrer' not used as faceplate branding",
     tab_source.count("Stohrer") == 1 and 'text="Stohrer"' not in tab_source)
test("'Experimental' IS in tuner_tab.py", "Experimental" in tab_source)

# ============================================================
# 9. Tab name check
# ============================================================
print("\n--- Tab Name ---")
with open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "main.py"), encoding="utf-8") as f:
    main_source = f.read()

test("Tab text is 'Tuner' not 'Strobe Tuner'",
     "text='Tuner'" in main_source and "text='Strobe Tuner'" not in main_source)

# ============================================================
# 10. Settings save method includes new keys (no tuner_mode)
# ============================================================
print("\n--- Settings Save ---")
save_source = inspect.getsource(TunerTabMixin._tuner_save_settings)
test("save includes ring_brightness", "ring_brightness" in save_source)
test("save includes overall_brightness", "overall_brightness" in save_source)
test("save includes faceplate_color", "faceplate_color" in save_source)
test("save does NOT include tuner_mode", "tuner_mode" not in save_source)

# ============================================================
# 11. User guide updated
# ============================================================
print("\n--- User Guide ---")
with open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "ui_dialogs.py"), encoding="utf-8") as f:
    guide_source = f.read()

test("User guide has 'Tuner' heading", 'self._h2("Tuner")' in guide_source)
test("User guide does NOT mention Simple Mode", "Simple Mode" not in guide_source)
test("User guide mentions VU meter", "VU meter" in guide_source or "analog" in guide_source.lower())
test("User guide mentions per-wheel BIAS NOTE control",
     "BIAS > NOTE" in guide_source)
test("User guide mentions Faceplate Color", "Faceplate Color" in guide_source)
test("User guide mentions DISP BRIGHT (master brightness)",
     "DISP > BRIGHT" in guide_source)

# ============================================================
# 12. Engine unchanged
# ============================================================
print("\n--- Engine Stability ---")
from tuner_engine import TunerEngine
import numpy as np

engine = TunerEngine()
test("Engine creates OK", engine is not None)
test("Reference pitch default is 440", engine.reference_pitch == 440.0)

result = engine.analyze_buffer(np.zeros(4096, dtype=np.float32))
test("Silent buffer gives zero magnitudes", all(m == 0.0 for m in result.magnitudes))
test("Silent buffer gives zero phase offsets", all(p == 0.0 for p in result.phase_offsets))

t = np.arange(4096, dtype=np.float64) / 44100
tone = (0.5 * np.sin(2 * np.pi * 440 * t)).astype(np.float32)
result = engine.analyze_buffer(tone)
test("A4 tone: pc 9 has highest magnitude",
     result.magnitudes[9] == max(result.magnitudes))
test("A4 tone: pc 9 is active", result.active[9])
test("A4 tone: ring_magnitudes has 7 rings per pc",
     len(result.ring_magnitudes[9]) == 7)

# ============================================================
# 13. StrobeWheel constructor
# ============================================================
print("\n--- StrobeWheel Constructor ---")
wheel_sig = inspect.signature(StrobeWheel.__init__)
wheel_params = list(wheel_sig.parameters.keys())
test("StrobeWheel.__init__ has faceplate_color param", "faceplate_color" in wheel_params)


# ============================================================
print(f"\n{'='*50}")
print(f"RESULTS: {passed} passed, {failed} failed out of {passed + failed}")
if failed:
    print("SOME TESTS FAILED!")
    sys.exit(1)
else:
    print("ALL TESTS PASSED!")
