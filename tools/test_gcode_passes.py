"""
Test script for the multi-pass G-code feature: per-layer passes settings
for pad materials (felt/card/leather) and acrylic in the tooling tab.
"""

import sys
import os
import copy
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
from gcode_engine import (
    generate_gcode_layer,
    generate_gcode_from_placed,
    generate_die_gcode_from_placed,
    generate_holder_gcode,
)


passed = 0
failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


# Sample stroke: a 4-point square outline. Two strokes so we exercise
# the multi-stroke path.
SAMPLE_STROKES = [
    [(0.0, 0.0), (10.0, 0.0), (10.0, 10.0), (0.0, 10.0), (0.0, 0.0)],
    [(20.0, 0.0), (30.0, 0.0), (30.0, 10.0), (20.0, 10.0), (20.0, 0.0)],
]


def count_g0_moves(lines):
    return sum(1 for ln in lines if ln.startswith("G0 "))


# ============================================================
print("\n=== generate_gcode_layer: passes=1 (default) ===")
# ============================================================

lines_1 = generate_gcode_layer(SAMPLE_STROKES, 600, 50, "C00")
g0_count_1 = count_g0_moves(lines_1)

check("Default passes=1 emits header without pass count",
      "x1 passes" not in "\n".join(lines_1) and "passes" not in lines_1[0].lower()
      or lines_1[0] == "; Cut @ 600 mm/min, 50% power")

check("Default passes=1 has one G0 per stroke",
      g0_count_1 == len(SAMPLE_STROKES))

check("Default passes=1 has no pass-marker comments",
      not any("Pass " in ln and " of " in ln for ln in lines_1))


# ============================================================
print("\n=== generate_gcode_layer: passes=3 ===")
# ============================================================

lines_3 = generate_gcode_layer(SAMPLE_STROKES, 600, 50, "C00", passes=3)
g0_count_3 = count_g0_moves(lines_3)

check("passes=3 header notes the multi-pass count",
      "x3 passes" in lines_3[0])

check("passes=3 has 3x the G0 moves (each stroke repeated)",
      g0_count_3 == len(SAMPLE_STROKES) * 3)

check("passes=3 emits pass marker comments",
      any("Pass 1 of 3" in ln for ln in lines_3)
      and any("Pass 2 of 3" in ln for ln in lines_3)
      and any("Pass 3 of 3" in ln for ln in lines_3))


# ============================================================
print("\n=== generate_gcode_layer: edge cases ===")
# ============================================================

# Zero/negative passes coerce to 1
lines_zero = generate_gcode_layer(SAMPLE_STROKES, 600, 50, "C00", passes=0)
check("passes=0 coerces to 1 (no infinite loop, single-pass output)",
      count_g0_moves(lines_zero) == len(SAMPLE_STROKES))

lines_neg = generate_gcode_layer(SAMPLE_STROKES, 600, 50, "C00", passes=-2)
check("passes=-2 coerces to 1",
      count_g0_moves(lines_neg) == len(SAMPLE_STROKES))

lines_str = generate_gcode_layer(SAMPLE_STROKES, 600, 50, "C00", passes="garbage")
check("passes='garbage' coerces to 1",
      count_g0_moves(lines_str) == len(SAMPLE_STROKES))

# Empty strokes -> empty output regardless of passes count
empty = generate_gcode_layer([], 600, 50, "C00", passes=5)
check("Empty strokes -> empty output even with passes=5",
      empty == [])


# ============================================================
print("\n=== Pad pipeline: cut_passes flows through to G-code ===")
# ============================================================

settings_default = copy.deepcopy(DEFAULT_SETTINGS)
settings_3pass = copy.deepcopy(DEFAULT_SETTINGS)
settings_3pass["gcode_settings"]["leather"]["cut_passes"] = 3

# Single placed pad: (size, cx, cy, radius)
placed = [(10.0, 30.0, 30.0, 5.0)]

with tempfile.TemporaryDirectory() as td:
    path1 = os.path.join(td, "default.gcode")
    path3 = os.path.join(td, "triple.gcode")
    generate_gcode_from_placed(placed, "leather", 100, 100, path1, 0, settings_default)
    generate_gcode_from_placed(placed, "leather", 100, 100, path3, 0, settings_3pass)

    with open(path1) as f:
        gc1 = f.read()
    with open(path3) as f:
        gc3 = f.read()

check("Default leather pad has no multi-pass marker",
      "x3 passes" not in gc1 and "Pass 1 of 3" not in gc1)

check("3-pass leather pad has 'x3 passes' marker on cut layer",
      "x3 passes" in gc3)

check("3-pass leather pad emits 'Pass N of 3' markers",
      "Pass 1 of 3" in gc3 and "Pass 3 of 3" in gc3)

# G-code grew by roughly 3x worth of cut moves
check("3-pass leather G-code is meaningfully larger than 1-pass",
      len(gc3) > len(gc1) * 1.5)


# ============================================================
print("\n=== Acrylic die-ring pipeline: cut_passes flows through ===")
# ============================================================

settings_acr_3 = copy.deepcopy(DEFAULT_SETTINGS)
settings_acr_3["gcode_settings"]["acrylic"]["cut_passes"] = 2

placed_die = [(10.0, 50.0, 50.0, 25.0)]  # 50mm die ring at center

with tempfile.TemporaryDirectory() as td:
    path_def = os.path.join(td, "die_default.gcode")
    path_2 = os.path.join(td, "die_2pass.gcode")
    generate_die_gcode_from_placed(placed_die, 200, 200, path_def, settings_default)
    generate_die_gcode_from_placed(placed_die, 200, 200, path_2, settings_acr_3)

    with open(path_def) as f:
        die_def = f.read()
    with open(path_2) as f:
        die_2 = f.read()

check("Default acrylic die has no multi-pass marker",
      "x2 passes" not in die_def)

check("2-pass acrylic die has 'x2 passes' marker (inner + outer cuts both repeat)",
      die_2.count("x2 passes") >= 2)


# ============================================================
print("\n=== Acrylic holder pipeline: cut_passes flows through ===")
# ============================================================

settings_holder = copy.deepcopy(DEFAULT_SETTINGS)
settings_holder["gcode_settings"]["acrylic"]["cut_passes"] = 2
settings_holder["gcode_settings"]["acrylic"]["hole_passes"] = 2

with tempfile.TemporaryDirectory() as td:
    path_def = os.path.join(td, "holder_default.gcode")
    path_2 = os.path.join(td, "holder_2pass.gcode")
    generate_holder_gcode("large", path_def, settings_default)
    generate_holder_gcode("large", path_2, settings_holder)

    with open(path_def) as f:
        h_def = f.read()
    with open(path_2) as f:
        h_2 = f.read()

check("Default acrylic holder has no multi-pass marker",
      "x2 passes" not in h_def)

check("2-pass acrylic holder repeats hole + cut layers",
      h_2.count("x2 passes") >= 2)


# ============================================================
print("\n=== DEFAULT_SETTINGS has passes=1 for every material ===")
# ============================================================

for mat in ("felt", "card", "leather", "acrylic"):
    mat_cfg = DEFAULT_SETTINGS["gcode_settings"][mat]
    for key in ("engraving_passes", "filled_engraving_passes", "hole_passes", "cut_passes"):
        check(f"{mat}.{key} default = 1",
              mat_cfg.get(key) == 1)


# ============================================================
print(f"\n=== Total: {passed} passed, {failed} failed ===")
# ============================================================

sys.exit(0 if failed == 0 else 1)
