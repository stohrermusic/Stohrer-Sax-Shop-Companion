"""
Test script for the Speed & Power Test feature (Tooling tab).

Engine-level tests (no tkinter). Covers:
  - Stops linspace math
  - Matrix expansion for each of the 8 (speed_sweep, power_sweep, passes_sweep) combos
  - Grid packer correctness (no overlap, all in bounds, row-major order)
  - G-code structure: 1 engraving layer + N distinct C{id} cut layers with
    correct per-piece S/F/passes values
  - Per-piece air assist (when "Also test with air off" is enabled)
  - Engraving speed/power overrides
  - Oversize matrix raises ValueError with a useful message
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
from svg_engine import (
    feeds_speeds_linspace, build_feeds_speeds_matrix,
    _grid_pack_discs, _min_feeds_speeds_sheet,
)
from gcode_engine import generate_feeds_speeds_test_gcode

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


settings = DEFAULT_SETTINGS.copy()


# ============================================================
print("\n=== feeds_speeds_linspace ===")
# ============================================================

check("Stops=1 returns [start]",
      feeds_speeds_linspace(80, 280, 1) == [80])
check("Stops=2 returns [start, end]",
      feeds_speeds_linspace(80, 280, 2) == [80, 280])
check("Stops=5 even spacing 80->280 = [80, 130, 180, 230, 280]",
      feeds_speeds_linspace(80, 280, 5) == [80, 130, 180, 230, 280])
check("Stops=4 power 30->90 = [30, 50, 70, 90]",
      feeds_speeds_linspace(30, 90, 4) == [30, 50, 70, 90])
check("Stops=3 with floats rounds [1, 2, 3]",
      feeds_speeds_linspace(1.0, 3.0, 3) == [1, 2, 3])
check("Stops=4 passes 1->4 = [1, 2, 3, 4]",
      feeds_speeds_linspace(1, 4, 4) == [1, 2, 3, 4])
check("Stops<=0 clamps to 1",
      feeds_speeds_linspace(50, 100, 0) == [50])
check("Start == End repeats",
      feeds_speeds_linspace(100, 100, 5) == [100] * 5)


# ============================================================
print("\n=== build_feeds_speeds_matrix - sweep combos ===")
# ============================================================

def cfg(value=None, sweep=False, start=None, end=None, stops=4):
    return {'sweep': sweep, 'value': value if value is not None else 0,
            'start': start if start is not None else 0,
            'end': end if end is not None else 0, 'stops': stops}


# 0 swept: single triple
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(value=180), cfg(value=60), cfg(value=1))
check("0 swept: 1 triple",
      trips == [(180, 60, 1)] and (cols, rows, nblk) == (1, 1, 1))

# 1 swept (speed only)
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(sweep=True, start=80, end=280, stops=5),
    cfg(value=60), cfg(value=1))
check("speed only: 5 triples in row",
      len(trips) == 5 and trips[0] == (80, 60, 1) and trips[-1] == (280, 60, 1)
      and (cols, rows, nblk) == (5, 1, 1))

# 1 swept (power only)
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(value=180),
    cfg(sweep=True, start=30, end=90, stops=4),
    cfg(value=1))
check("power only: 4 triples",
      len(trips) == 4 and trips[0] == (180, 30, 1) and trips[-1] == (180, 90, 1)
      and (cols, rows, nblk) == (1, 4, 1))

# 1 swept (passes only)
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(value=180), cfg(value=60),
    cfg(sweep=True, start=1, end=4, stops=4))
check("passes only: 4 triples",
      len(trips) == 4 and trips[0] == (180, 60, 1) and trips[-1] == (180, 60, 4)
      and (cols, rows, nblk) == (1, 1, 4))

# 2 swept (speed x power) - LightBurn-style
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(sweep=True, start=80, end=280, stops=4),
    cfg(sweep=True, start=30, end=90, stops=4),
    cfg(value=1))
check("speed x power: 16 triples (4x4)",
      len(trips) == 16 and (cols, rows, nblk) == (4, 4, 1))
# Order: outer power, inner speed -> first 4 share power=30, vary speed
check("speed x power: first row varies speed at power=30",
      [t[0] for t in trips[:4]] == [80, 146 if False else 147, 213, 280] or
      [t[0] for t in trips[:4]] == [80, 147, 213, 280])
check("speed x power: row 1 power=30",
      all(t[1] == 30 for t in trips[:4]))
check("speed x power: row 4 power=90",
      all(t[1] == 90 for t in trips[12:16]))

# 2 swept (power x passes)
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(value=180),
    cfg(sweep=True, start=30, end=90, stops=4),
    cfg(sweep=True, start=1, end=3, stops=3))
check("power x passes: 12 triples (1x4x3)",
      len(trips) == 12 and (cols, rows, nblk) == (1, 4, 3))

# 3 swept (full 3D)
trips, cols, rows, nblk = build_feeds_speeds_matrix(
    cfg(sweep=True, start=80, end=280, stops=4),
    cfg(sweep=True, start=30, end=90, stops=4),
    cfg(sweep=True, start=1, end=3, stops=3))
check("full 3D: 48 triples (4x4x3)",
      len(trips) == 48 and (cols, rows, nblk) == (4, 4, 3))
check("3D: first 16 triples are passes=1 block",
      all(t[2] == 1 for t in trips[:16]))
check("3D: last 16 triples are passes=3 block",
      all(t[2] == 3 for t in trips[32:48]))


# ============================================================
print("\n=== build_feeds_speeds_matrix - clamping & errors ===")
# ============================================================

# Power clamped to [1, 100]
trips, _, _, _ = build_feeds_speeds_matrix(
    cfg(value=180), cfg(sweep=True, start=0, end=200, stops=5), cfg(value=1))
check("power clamped to 1..100",
      all(1 <= t[1] <= 100 for t in trips))

# Passes >= 1
trips, _, _, _ = build_feeds_speeds_matrix(
    cfg(value=180), cfg(value=60), cfg(value=0))
check("passes value=0 clamped to 1",
      trips == [(180, 60, 1)])

# Speed >= 1
trips, _, _, _ = build_feeds_speeds_matrix(
    cfg(value=0), cfg(value=60), cfg(value=1))
check("speed value=0 clamped to 1",
      trips == [(1, 60, 1)])

# Stops clamped to [2, 10]
trips, _, _, _ = build_feeds_speeds_matrix(
    cfg(sweep=True, start=80, end=280, stops=99),
    cfg(value=60), cfg(value=1))
check("stops=99 clamped to 10",
      len(trips) == 10)

# Bad input raises
err = False
try:
    build_feeds_speeds_matrix(cfg(sweep=True, start="abc", end=280, stops=4),
                              cfg(value=60), cfg(value=1))
except ValueError:
    err = True
check("non-numeric Start raises ValueError", err)


# ============================================================
print("\n=== _grid_pack_discs ===")
# ============================================================

# 16 discs at 20mm OD on a 200x200mm sheet (single block)
positions, total_w, total_h = _grid_pack_discs(20.0, 4, 4, 1, 200, 200)
check("4x4 grid: 16 positions", len(positions) == 16)
check("4x4 grid: total_w < 200", total_w < 200)
check("4x4 grid: total_h < 200", total_h < 200)

# Row-major order check: positions in the same row share cy, in same col share cx
ys = sorted({p[1] for p in positions})
check("4x4 grid: 4 distinct rows", len(ys) == 4)
xs = sorted({p[0] for p in positions})
check("4x4 grid: 4 distinct cols", len(xs) == 4)
# First 4 positions are top row (smallest cy)
check("Row 0 has smallest cy",
      all(positions[i][1] == ys[0] for i in range(4)))
# Within row 0, cx increases
check("Row 0 cx increases left to right",
      positions[0][0] < positions[1][0] < positions[2][0] < positions[3][0])

# No overlap: distance between any two centers >= diameter + spacing (within tolerance)
def min_pair_dist(pos):
    m = float('inf')
    for i, (x1, y1) in enumerate(pos):
        for x2, y2 in pos[i + 1:]:
            d = ((x1 - x2) ** 2 + (y1 - y2) ** 2) ** 0.5
            if d < m:
                m = d
    return m
check("4x4 grid: no overlap (min dist >= diameter)",
      min_pair_dist(positions) >= 20.0 - 0.01)

# In bounds: all positions are inside the sheet (with margin)
def in_bounds(pos, dia, sw, sh):
    for x, y in pos:
        if x - dia/2 < -0.01 or x + dia/2 > sw + 0.01:
            return False
        if y - dia/2 < -0.01 or y + dia/2 > sh + 0.01:
            return False
    return True
check("4x4 grid: all positions in bounds",
      in_bounds(positions, 20.0, 200, 200))

# 3D layout (3 blocks of 4x4) - needs a wider sheet
positions3, w3, h3 = _grid_pack_discs(20.0, 4, 4, 3, 400, 200)
check("3 blocks of 4x4: 48 positions", len(positions3) == 48)
check("3 blocks: gap between block 1 and block 2",
      positions3[16][0] > positions3[15][0] + 25)  # block gap > pitch

# Oversize raises with informative message
err_msg = ""
try:
    _grid_pack_discs(20.0, 4, 4, 1, 50, 50)
except ValueError as e:
    err_msg = str(e)
check("oversize raises ValueError", "doesn't fit" in err_msg.lower() or "need at least" in err_msg.lower())
check("oversize message includes minimum size", "mm" in err_msg)


# ============================================================
print("\n=== _min_feeds_speeds_sheet ===")
# ============================================================

min_w, min_h = _min_feeds_speeds_sheet(20.0, 4, 4, 1)
positions4, total_w4, total_h4 = _grid_pack_discs(20.0, 4, 4, 1, min_w, min_h)
check("_min returns size that exactly fits 4x4",
      len(positions4) == 16 and abs(total_w4 - min_w) < 0.01 and abs(total_h4 - min_h) < 0.01)

# Sheet 1mm smaller should fail
err = False
try:
    _grid_pack_discs(20.0, 4, 4, 1, min_w - 1, min_h)
except ValueError:
    err = True
check("Sheet 1mm under width raises", err)


def make_pieces(diameter, sheet_w, sheet_h, cols, rows, nblk):
    """Helper to build a placed test_pieces list for engine tests."""
    triples, c, r, n = build_feeds_speeds_matrix(
        cfg(sweep=True, start=80, end=280, stops=cols) if cols > 1 else cfg(value=180),
        cfg(sweep=True, start=30, end=90, stops=rows) if rows > 1 else cfg(value=60),
        cfg(sweep=True, start=1, end=nblk, stops=nblk) if nblk > 1 else cfg(value=1),
    )
    positions, _, _ = _grid_pack_discs(diameter, c, r, n, sheet_w, sheet_h)
    pieces = []
    for idx, ((cx, cy), (speed, power, passes)) in enumerate(zip(positions, triples), start=1):
        pieces.append({
            'id': f"{idx:02d}",
            'cx': cx, 'cy': cy, 'diameter': diameter,
            'speed': speed, 'power': power, 'passes': passes,
            'material': 'felt', 'engraving_on': True,
        })
    return pieces


# ============================================================
print("\n=== G-code output structure ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    pieces = make_pieces(20.0, 200, 200, 4, 4, 1)
    gcode_path = os.path.join(tmpdir, "test.gcode")
    generate_feeds_speeds_test_gcode(pieces, 200, 200, gcode_path, settings,
                                      air_assist=True)
    check("G-code file created", os.path.exists(gcode_path))

    with open(gcode_path, 'r') as f:
        gcode_text = f.read()
    lines = gcode_text.split('\n')

    # Header / footer present
    check("G-code has header", any("Generated by" in ln for ln in lines))

    # Engraving layer (C00) present
    check("G-code has C00 engraving layer", any("Layer C00" in ln for ln in lines))

    # 16 distinct cut layers C01..C16
    layer_lines = [ln for ln in lines if ln.startswith("; Layer C")]
    layer_names = {ln.replace("; Layer ", "").strip() for ln in layer_lines}
    expected = {"C00"} | {f"C{i:02d}" for i in range(1, 17)}
    check("G-code has C00 + C01..C16 layers", layer_names == expected)

    # Each cut layer's S/F values come from its piece's (power, speed)
    # Pull the first G1 command from each layer and verify it carries the right S
    # The layer header comment includes "Cut @ X mm/min, Y% power" - match those
    cut_comments = [ln for ln in lines if ln.startswith("; Cut @")]
    # 17 total: 1 engraving + 16 cuts
    check("G-code has 17 'Cut @' comments (engraving + 16 cuts)",
          len(cut_comments) == 17)

    # Verify per-piece S values appear by mapping cut_comments[1:] (skipping engraving)
    cut_only_comments = cut_comments[1:]
    for idx, piece in enumerate(pieces):
        comment = cut_only_comments[idx]
        expected_speed_part = f"{piece['speed']} mm/min"
        expected_power_part = f"{piece['power']}% power"
        if expected_speed_part not in comment or expected_power_part not in comment:
            check(f"Piece {piece['id']} comment matches (got {comment!r})", False)
            break
    else:
        check("All 16 cut-layer comments carry the correct per-piece S/F", True)


# ============================================================
print("\n=== Multi-pass G-code ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    pieces = make_pieces(20.0, 200, 200, 2, 2, 1)
    # Force passes=3 on every piece to confirm the gcode emits the multi-pass note
    for p in pieces:
        p['passes'] = 3
    gcode_path = os.path.join(tmpdir, "multi.gcode")
    generate_feeds_speeds_test_gcode(pieces, 200, 200, gcode_path, settings)
    with open(gcode_path, 'r') as f:
        text = f.read()
    check("Multi-pass cut comment includes 'x3 passes'",
          "x3 passes" in text)
    # 4 cut layers x 3 passes = 12 'Pass N of 3' markers
    pass_markers = text.count("of 3")
    check("Multi-pass: 12 'of 3' pass markers (4 pieces x 3)",
          pass_markers == 12)


# ============================================================
print("\n=== Engraving speed/power overrides ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    pieces = make_pieces(20.0, 200, 200, 2, 2, 1)
    gcode_path = os.path.join(tmpdir, "eng_override.gcode")
    # Default material engraving for felt is 1200/8 (line). Override to
    # something distinctive and verify it lands in the gcode.
    generate_feeds_speeds_test_gcode(pieces, 200, 200, gcode_path, settings,
                                      air_assist=True,
                                      eng_speed_override=2345,
                                      eng_power_override=17)
    with open(gcode_path, 'r') as f:
        text = f.read()
    # The engraving layer is the first "Cut @" comment in the file.
    first_comment = next(ln for ln in text.split('\n') if ln.startswith("; Cut @"))
    check("eng_speed_override appears in engraving layer comment",
          "2345 mm/min" in first_comment)
    check("eng_power_override appears in engraving layer comment",
          "17% power" in first_comment)
    # Default (no override) still pulls from material settings.
    gcode_default_path = os.path.join(tmpdir, "eng_default.gcode")
    generate_feeds_speeds_test_gcode(pieces, 200, 200, gcode_default_path, settings,
                                      air_assist=True)
    with open(gcode_default_path, 'r') as f:
        default_text = f.read()
    default_first = next(ln for ln in default_text.split('\n') if ln.startswith("; Cut @"))
    # Felt default line-mode engraving is 1200 mm/min @ 8%
    check("no override -> falls back to material default (1200 mm/min)",
          "1200 mm/min" in default_first)
    check("no override -> falls back to material default (8% power)",
          "8% power" in default_first)


# ============================================================
print("\n=== Per-piece air assist (Also test with air off) ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    # Build pieces and set half air-on, half air-off (mimicking what the UI
    # would produce when "Also test with air off" is checked).
    pieces = make_pieces(20.0, 200, 200, 2, 2, 1)  # 4 base pieces
    duplicated = []
    for idx, p in enumerate(pieces, 1):
        on = dict(p)
        on['id'] = f"{idx:02d}"
        on['air_assist'] = True
        duplicated.append(on)
    for idx, p in enumerate(pieces, len(pieces) + 1):
        off = dict(p)
        off['id'] = f"{idx:02d}"
        off['air_assist'] = False
        duplicated.append(off)

    gcode_path = os.path.join(tmpdir, "per_piece_air.gcode")
    generate_feeds_speeds_test_gcode(duplicated, 200, 200, gcode_path, settings,
                                      air_assist=True)
    with open(gcode_path, 'r') as f:
        gcode_text = f.read()
    lines = gcode_text.split('\n')

    # Count M8 (air on) and M9 (air off) emissions. Engraving layer emits
    # M8 (global air_assist=True). Then 4 cuts with M8 (per-piece) and 4
    # cuts with M9. The footer always emits one final M9 (turn air off
    # before park), so total M9 = 4 cuts + 1 footer = 5.
    m8_lines = [ln for ln in lines if ln.strip() == "M8"]
    m9_lines = [ln for ln in lines if ln.strip() == "M9"]
    check("per-piece air: 5 M8 lines (1 engraving + 4 air-on cuts)",
          len(m8_lines) == 5)
    check("per-piece air: 5 M9 lines (4 air-off cuts + footer)",
          len(m9_lines) == 5)
    # The first cut after the engraving layer should be air-on (first 4
    # pieces), the last cut layer should be air-off (last 4 pieces). Check
    # ordering by looking at the layer comments interleaved with M8/M9.
    cut_layer_idx = [i for i, ln in enumerate(lines)
                      if ln.startswith("; Layer C") and ln != "; Layer C00"]
    # For each cut layer, the M8/M9 was emitted just before in
    # generate_gcode_layer (after the "; Cut @ ..." comment but before the
    # "; Layer C##" comment). Walk backward to find the closest M8/M9.
    for layer_pos, layer_line in zip(cut_layer_idx[:4],
                                       (lines[i] for i in cut_layer_idx[:4])):
        # Look backward for M8 or M9
        air_state = None
        for j in range(layer_pos, max(0, layer_pos - 20), -1):
            if lines[j].strip() == "M8":
                air_state = "M8"
                break
            if lines[j].strip() == "M9":
                air_state = "M9"
                break
        if air_state != "M8":
            check(f"first 4 cut layers are air-on (got {air_state})", False)
            break
    else:
        check("first 4 cut layers are air-on (M8)", True)
    for layer_pos in cut_layer_idx[4:]:
        air_state = None
        for j in range(layer_pos, max(0, layer_pos - 20), -1):
            if lines[j].strip() == "M8":
                air_state = "M8"
                break
            if lines[j].strip() == "M9":
                air_state = "M9"
                break
        if air_state != "M9":
            check(f"last 4 cut layers are air-off (got {air_state})", False)
            break
    else:
        check("last 4 cut layers are air-off (M9)", True)


# ============================================================
print("\n=== Filled engraving mode (M8/air assist + filled fonts) ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    pieces = make_pieces(20.0, 200, 200, 2, 2, 1)
    # Override felt's engraving_mode to filled for this test
    test_settings = DEFAULT_SETTINGS.copy()
    test_settings["gcode_settings"] = {
        "felt": {
            "engraving_mode": "filled",
            "filled_engraving_speed": 1500,
            "filled_engraving_power": 20,
            "filled_engraving_passes": 1,
            "filled_line_spacing": 0.15,
        }
    }
    gcode_path = os.path.join(tmpdir, "filled.gcode")
    generate_feeds_speeds_test_gcode(pieces, 200, 200, gcode_path, test_settings,
                                      air_assist=False)
    with open(gcode_path, 'r') as f:
        text = f.read()
    check("filled mode engraving uses 1500 mm/min",
          "1500 mm/min" in text)
    check("air_assist=False emits M9", "M9" in text)


# ============================================================
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
