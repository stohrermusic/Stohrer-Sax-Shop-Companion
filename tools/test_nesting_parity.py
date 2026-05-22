"""
Parity test: vectorized numpy radial scan vs Python reference.

The Python reference (`_scan_radial_python`) is preserved verbatim from the
pre-vectorization implementation. The numpy version (`_scan_radial_numpy`)
MUST return bit-identical results for the same inputs, otherwise packing
density and visual output would drift across builds.

Also covers end-to-end parity through `_nest_discs` and an informational
timing comparison for the card-paper scrap-mode scenario.
"""

import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
import svg_engine
from svg_engine import (
    _scan_radial_python, _scan_radial_numpy, _nest_discs, _HAS_NUMPY,
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


if not _HAS_NUMPY:
    print("\n!! numpy not available — skipping parity test")
    sys.exit(0)


# ============================================================
print("\n=== Direct parity: _scan_radial_python vs _scan_radial_numpy ===")
# ============================================================

def parity_case(label, dia, placed_list, target_x, target_y, w, h, spacing=1.0):
    r_val = dia / 2.0
    py = _scan_radial_python(dia, r_val, placed_list, target_x, target_y, w, h, spacing)
    nz = _scan_radial_numpy(dia, r_val, placed_list, target_x, target_y, w, h, spacing)
    if py is None and nz is None:
        check(f"{label}: both None", True)
        return
    if py is None or nz is None:
        check(f"{label}: one is None (py={py}, np={nz})", False)
        return
    same = abs(py[0] - nz[0]) < 1e-9 and abs(py[1] - nz[1]) < 1e-9
    check(f"{label}: positions match (py={py}, np={nz})", same)


# Empty sheet, all 5 radial targets ------------------------------------
sheet_w, sheet_h = 200.0, 200.0
parity_case("center-target empty 200x200", 20, [], 100, 100, sheet_w, sheet_h)
parity_case("nw-target empty 200x200", 20, [], 0, 0, sheet_w, sheet_h)
parity_case("ne-target empty 200x200", 20, [], sheet_w, 0, sheet_w, sheet_h)
parity_case("sw-target empty 200x200", 20, [], 0, sheet_h, sheet_w, sheet_h)
parity_case("se-target empty 200x200", 20, [], sheet_w, sheet_h, sheet_w, sheet_h)

# Card-paper sized sheet (A4) ------------------------------------------
parity_case("center-target empty A4 (210x297)", 20, [], 105, 148.5, 210, 297)
parity_case("nw-target empty A4", 20, [], 0, 0, 210, 297)

# Square sheets of various sizes ---------------------------------------
for w in (50, 80, 120, 200, 300):
    parity_case(f"center empty {w}x{w}", 15, [], w / 2, w / 2, w, w)

# Disc just barely fits ------------------------------------------------
parity_case("disc fits exactly 22x22", 20, [], 11, 11, 22, 22)
parity_case("disc fits exactly 21x21 (boundary)", 20, [], 10.5, 10.5, 21, 21)

# Sheet too small — both should return None
parity_case("too small 10x10", 20, [], 5, 5, 10, 10)
parity_case("too narrow 5x100", 20, [], 2.5, 50, 5, 100)

# Various spacings -----------------------------------------------------
parity_case("spacing=2 200x200", 20, [], 100, 100, 200, 200, spacing=2.0)
parity_case("spacing=1.5 100x100", 20, [], 50, 50, 100, 100, spacing=1.5)
parity_case("spacing=0.5 100x100", 20, [], 50, 50, 100, 100, spacing=0.5)


# ============================================================
print("\n=== Parity with obstacles in the placed list ===")
# ============================================================

# One obstacle near center
placed_1 = [(15.0, 100.0, 100.0, 7.5)]
parity_case("center-target, 1 obstacle at center", 20, placed_1, 100, 100, 200, 200)
parity_case("nw-target, 1 obstacle at center", 20, placed_1, 0, 0, 200, 200)

# A small cluster
placed_cluster = [
    (20.0, 50.0, 50.0, 10.0),
    (20.0, 90.0, 50.0, 10.0),
    (20.0, 130.0, 50.0, 10.0),
    (20.0, 50.0, 90.0, 10.0),
    (20.0, 90.0, 90.0, 10.0),
]
for tgt_name, tx, ty in [("center", 100, 100), ("nw", 0, 0),
                          ("ne", 200, 0), ("sw", 0, 200), ("se", 200, 200)]:
    parity_case(f"{tgt_name}-target, cluster of 5", 20, placed_cluster, tx, ty, 200, 200)

# Crowded — many obstacles
placed_crowded = []
for ry in range(20, 200, 25):
    for rx in range(20, 200, 25):
        placed_crowded.append((10.0, float(rx), float(ry), 5.0))
parity_case("center-target, 49 obstacles in grid", 12, placed_crowded, 100, 100, 200, 200)
parity_case("se-target, 49 obstacles in grid", 12, placed_crowded, 200, 200, 200, 200)

# Mixed-size obstacles
placed_mixed = [
    (30.0, 60.0, 60.0, 15.0),
    (20.0, 130.0, 60.0, 10.0),
    (15.0, 90.0, 120.0, 7.5),
]
parity_case("center-target, mixed obstacle sizes", 18, placed_mixed, 100, 100, 200, 200)


# ============================================================
print("\n=== End-to-end parity through _nest_discs ===")
# ============================================================

def nest_e2e_case(label, pads, material, w, h, settings):
    """Run _nest_discs with numpy enabled and disabled, assert identical placements."""
    saved = svg_engine._HAS_NUMPY
    try:
        # Force python path
        svg_engine._HAS_NUMPY = False
        placed_py, fp_py, ft_py = _nest_discs(pads, material, w, h, settings)
        # Numpy path
        svg_engine._HAS_NUMPY = True
        placed_np, fp_np, ft_np = _nest_discs(pads, material, w, h, settings)
    finally:
        svg_engine._HAS_NUMPY = saved

    if len(placed_py) != len(placed_np):
        check(f"{label}: placement counts match (py={len(placed_py)}, np={len(placed_np)})", False)
        return
    if (fp_py, ft_py) != (fp_np, ft_np):
        check(f"{label}: fixed_placed/fixed_total match", False)
        return
    ok = True
    for (a, b) in zip(placed_py, placed_np):
        if a[0] != b[0] or abs(a[1] - b[1]) > 1e-9 or abs(a[2] - b[2]) > 1e-9 or abs(a[3] - b[3]) > 1e-9:
            ok = False
            break
    check(f"{label}: {len(placed_py)} placements identical", ok)


settings = DEFAULT_SETTINGS.copy()

# Small fixed-set
settings["edge_bias"] = "center"
pads_small = [{'size': 20.0, 'qty': 4}]
nest_e2e_case("4 felt pads on 100x80, center", pads_small, "felt", 100.0, 80.0, settings)

# Mixed sizes
pads_mixed = [{'size': float(s), 'qty': 2} for s in (10, 15, 20, 25, 30, 35)]
nest_e2e_case("12 felt pads mixed sizes 200x200", pads_mixed, "felt", 200.0, 200.0, settings)

# Card on A4 — Matt's actual scenario
pads_card_typical = [{'size': float(s), 'qty': 1} for s in
                      (7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                       21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33)]
nest_e2e_case("27 card pads on A4 paper (210x297)", pads_card_typical, "card", 210.0, 297.0, settings)

# All radial biases
for bias in ("center", "nw", "ne", "sw", "se"):
    settings["edge_bias"] = bias
    nest_e2e_case(f"bias={bias}, 4 pads 100x80", pads_small, "felt", 100.0, 80.0, settings)

# Cardinal biases should be unaffected (they take a different code path)
for bias in ("n", "s", "e", "w"):
    settings["edge_bias"] = bias
    nest_e2e_case(f"bias={bias} (cardinal, unchanged path), 4 pads", pads_small, "felt", 100.0, 80.0, settings)

# Max-quantity fills
settings["edge_bias"] = "center"
pads_max = [{'size': 20.0, 'qty': 2}, {'size': 10.0, 'qty': 'max'}]
nest_e2e_case("2 fixed + max small pads, 150x150 center", pads_max, "felt", 150.0, 150.0, settings)


# ============================================================
print("\n=== Edge cases ===")
# ============================================================

# Sheet exactly fits one disc
parity_case("center 22x22 dia=20", 20, [], 11, 11, 22, 22)

# Off-target outside sheet — common when target is corner
parity_case("target outside sheet (north of)", 20, [], 100, -50, 200, 200)
parity_case("target outside sheet (south of)", 20, [], 100, 300, 200, 200)

# Small disc on tiny sheet
parity_case("dia=4 on 10x10", 4, [], 5, 5, 10, 10)


# ============================================================
print("\n=== Timing: card-paper scrap-mode scenario (informational) ===")
# ============================================================

# Card on A4 with all small sizes — close to the path Matt was hitting
settings["edge_bias"] = "center"
pads_realworld = [{'size': float(s), 'qty': 1} for s in range(7, 40)]  # 33 pads 7-39mm

svg_engine._HAS_NUMPY = False
t0 = time.perf_counter()
placed_py, _, _ = _nest_discs(pads_realworld, "card", 210.0, 297.0, settings)
t_py = time.perf_counter() - t0

svg_engine._HAS_NUMPY = True
t0 = time.perf_counter()
placed_np, _, _ = _nest_discs(pads_realworld, "card", 210.0, 297.0, settings)
t_np = time.perf_counter() - t0

print(f"  Python: {t_py:.2f}s ({len(placed_py)} pads placed)")
print(f"  Numpy:  {t_np:.2f}s ({len(placed_np)} pads placed)")
if t_np > 0:
    print(f"  Speedup: {t_py / t_np:.1f}x")

# Hard requirement: numpy must be at least as fast as Python on this scenario.
# (Loose threshold — the win is much bigger than 2x in practice, but we want
# the test to pass even on a slow CI box.)
check("numpy at least 2x faster than Python on card+A4 scenario",
      t_py / max(t_np, 0.001) >= 2.0)

# And the placements must match exactly.
check(f"timing-test placements identical ({len(placed_py)} each)",
      len(placed_py) == len(placed_np)
      and all(a[0] == b[0]
              and abs(a[1] - b[1]) < 1e-9
              and abs(a[2] - b[2]) < 1e-9
              and abs(a[3] - b[3]) < 1e-9
              for a, b in zip(placed_py, placed_np)))


# ============================================================
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
