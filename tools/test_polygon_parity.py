"""
Parity test: vectorized polygon nesting vs Python reference.

Confirms the numpy versions of polygon scan functions produce bit-identical
placements to the Python references across:
  - Grid helpers (point-in-polygon, edge-distance, vertex-distance)
  - Large-disc scan (centroid + bias scoring)
  - Small-disc scan (edge/vertex/snugness/bias scoring)
  - End-to-end through _nest_discs_polygon for several polygon shapes,
    including rectangles, triangles, L-shapes, hexagons, and concave shapes

Plus an informational timing measurement for a realistic scrap-shape scenario.
"""

import sys
import os
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
import svg_engine
from svg_engine import (
    _point_in_polygon, _distance_to_nearest_edge, _distance_to_nearest_vertex,
    _points_in_polygon_grid, _distances_grid_to_polygon_edges,
    _distances_grid_to_vertices,
    _find_best_polygon_large_python, _find_best_polygon_large_numpy,
    _find_best_polygon_small_python, _find_best_polygon_small_numpy,
    _nest_discs_polygon, _HAS_NUMPY,
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
    print("\n!! numpy not available - skipping polygon parity test")
    sys.exit(0)

import numpy as np


# Polygon shapes used throughout the tests --------------------------------
RECT = [(0.0, 0.0), (200.0, 0.0), (200.0, 150.0), (0.0, 150.0)]
TRIANGLE = [(0.0, 0.0), (200.0, 0.0), (100.0, 173.2)]
HEXAGON = [(100.0, 0.0), (200.0, 50.0), (200.0, 150.0),
           (100.0, 200.0), (0.0, 150.0), (0.0, 50.0)]
# L-shape (concave)
L_SHAPE = [(0.0, 0.0), (150.0, 0.0), (150.0, 80.0),
           (80.0, 80.0), (80.0, 200.0), (0.0, 200.0)]
# Tall narrow strip
STRIP = [(0.0, 0.0), (80.0, 0.0), (80.0, 250.0), (0.0, 250.0)]
# Concave U-shape
U_SHAPE = [(0.0, 0.0), (200.0, 0.0), (200.0, 200.0),
           (140.0, 200.0), (140.0, 60.0), (60.0, 60.0),
           (60.0, 200.0), (0.0, 200.0)]


# ============================================================
print("\n=== Grid helpers vs scalar references ===")
# ============================================================

def _check_grid_pip(label, polygon):
    """Confirm vectorized point-in-poly matches the scalar version for a grid of points."""
    cxs = np.arange(-10.0, 220.0, 5.0)
    cys = np.arange(-10.0, 260.0, 5.0)
    grid = _points_in_polygon_grid(cxs, cys, polygon)
    bad = 0
    for yi, cy in enumerate(cys):
        for xi, cx in enumerate(cxs):
            scalar = _point_in_polygon(float(cx), float(cy), polygon)
            if bool(grid[yi, xi]) != bool(scalar):
                bad += 1
    check(f"{label}: point-in-polygon grid matches scalar ({bad} mismatches)", bad == 0)


def _check_grid_edges(label, polygon):
    cxs = np.arange(0.0, 220.0, 7.0)
    cys = np.arange(0.0, 260.0, 7.0)
    grid = _distances_grid_to_polygon_edges(cxs, cys, polygon)
    max_diff = 0.0
    for yi, cy in enumerate(cys):
        for xi, cx in enumerate(cxs):
            scalar = _distance_to_nearest_edge(float(cx), float(cy), polygon)
            diff = abs(float(grid[yi, xi]) - scalar)
            if diff > max_diff:
                max_diff = diff
    check(f"{label}: edge-distance grid matches scalar (max diff={max_diff:.2e})",
          max_diff < 1e-9)


def _check_grid_vertices(label, polygon):
    cxs = np.arange(0.0, 220.0, 7.0)
    cys = np.arange(0.0, 260.0, 7.0)
    grid = _distances_grid_to_vertices(cxs, cys, polygon)
    max_diff = 0.0
    for yi, cy in enumerate(cys):
        for xi, cx in enumerate(cxs):
            scalar = _distance_to_nearest_vertex(float(cx), float(cy), polygon)
            diff = abs(float(grid[yi, xi]) - scalar)
            if diff > max_diff:
                max_diff = diff
    check(f"{label}: vertex-distance grid matches scalar (max diff={max_diff:.2e})",
          max_diff < 1e-9)


for name, poly in [("RECT", RECT), ("TRIANGLE", TRIANGLE), ("HEXAGON", HEXAGON),
                    ("L_SHAPE", L_SHAPE), ("U_SHAPE", U_SHAPE)]:
    _check_grid_pip(name, poly)
    _check_grid_edges(name, poly)
    _check_grid_vertices(name, poly)


# ============================================================
print("\n=== Direct parity: large-disc scan ===")
# ============================================================

def _bbox(polygon):
    xs = [p[0] for p in polygon]
    ys = [p[1] for p in polygon]
    return min(xs), max(xs), min(ys), max(ys)


def _centroid(polygon):
    n = len(polygon)
    return (sum(p[0] for p in polygon) / n, sum(p[1] for p in polygon) / n)


def _parity_large(label, polygon, r, placed_discs, bias_target=None,
                   bias_weight=0.0, spacing_mm=1.0, step=1):
    mnx, mxx, mny, mxy = _bbox(polygon)
    cx_c, cy_c = _centroid(polygon)
    if bias_target is None:
        bx, by = cx_c, cy_c
    else:
        bx, by = bias_target
    py = _find_best_polygon_large_python(
        r, placed_discs, polygon, mnx, mxx, mny, mxy, cx_c, cy_c,
        bx, by, bias_weight, spacing_mm, step)
    nz = _find_best_polygon_large_numpy(
        r, placed_discs, polygon, mnx, mxx, mny, mxy, cx_c, cy_c,
        bx, by, bias_weight, spacing_mm, step)
    if py is None and nz is None:
        check(f"{label}: both None", True)
        return
    if py is None or nz is None:
        check(f"{label}: one None (py={py}, np={nz})", False)
        return
    same = abs(py[0] - nz[0]) < 1e-9 and abs(py[1] - nz[1]) < 1e-9
    check(f"{label}: large positions match (py={py}, np={nz})", same)


# Empty / single placements on each shape, no bias
for name, poly in [("RECT", RECT), ("TRIANGLE", TRIANGLE), ("HEXAGON", HEXAGON),
                    ("L_SHAPE", L_SHAPE), ("U_SHAPE", U_SHAPE)]:
    _parity_large(f"{name} r=10, empty, center", poly, 10, [])
    _parity_large(f"{name} r=15, empty, center", poly, 15, [])
    _parity_large(f"{name} r=20, empty, center", poly, 20, [])

# With placed discs (collision)
placed_1 = [(20.0, 100.0, 75.0, 10.0)]
_parity_large("RECT r=10, 1 placed, center", RECT, 10, placed_1)
_parity_large("HEXAGON r=15, 1 placed, center", HEXAGON, 15, placed_1)
_parity_large("L_SHAPE r=8, 1 placed, center", L_SHAPE, 8, placed_1)

# With bias
for bias_name, target, weight in [
    ("nw", (0.0, 0.0), 2.0),
    ("ne", (200.0, 0.0), 2.0),
    ("se", (200.0, 150.0), 2.0),
    ("sw", (0.0, 150.0), 2.0),
]:
    _parity_large(f"RECT bias={bias_name} r=12, empty", RECT, 12, [],
                   bias_target=target, bias_weight=weight)
    _parity_large(f"L_SHAPE bias={bias_name} r=10, empty", L_SHAPE, 10, [],
                   bias_target=target, bias_weight=weight)

# Crowded — multiple placed discs
placed_crowd = [
    (15.0, 30.0, 30.0, 7.5),
    (15.0, 60.0, 30.0, 7.5),
    (15.0, 90.0, 30.0, 7.5),
    (15.0, 30.0, 60.0, 7.5),
    (15.0, 60.0, 60.0, 7.5),
]
_parity_large("RECT crowded r=12", RECT, 12, placed_crowd)
_parity_large("HEXAGON crowded r=10", HEXAGON, 10, placed_crowd)

# Disc too big to fit anywhere
_parity_large("STRIP r=50 (doesn't fit)", STRIP, 50, [])


# ============================================================
print("\n=== Direct parity: small-disc scan ===")
# ============================================================

def _parity_small(label, polygon, r, placed_discs, bias_target=None,
                   bias_weight=0.0, spacing_mm=1.0, step=1):
    mnx, mxx, mny, mxy = _bbox(polygon)
    cx_c, cy_c = _centroid(polygon)
    if bias_target is None:
        bx, by = cx_c, cy_c
    else:
        bx, by = bias_target
    py = _find_best_polygon_small_python(
        r, placed_discs, polygon, mnx, mxx, mny, mxy,
        bx, by, bias_weight, spacing_mm, step)
    nz = _find_best_polygon_small_numpy(
        r, placed_discs, polygon, mnx, mxx, mny, mxy,
        bx, by, bias_weight, spacing_mm, step)
    if py is None and nz is None:
        check(f"{label}: both None", True)
        return
    if py is None or nz is None:
        check(f"{label}: one None (py={py}, np={nz})", False)
        return
    same = abs(py[0] - nz[0]) < 1e-9 and abs(py[1] - nz[1]) < 1e-9
    check(f"{label}: small positions match (py={py}, np={nz})", same)


for name, poly in [("RECT", RECT), ("TRIANGLE", TRIANGLE), ("HEXAGON", HEXAGON),
                    ("L_SHAPE", L_SHAPE), ("U_SHAPE", U_SHAPE)]:
    _parity_small(f"{name} r=5, empty, center", poly, 5, [])
    _parity_small(f"{name} r=8, empty, center", poly, 8, [])

_parity_small("RECT small r=5, 5 placed", RECT, 5, placed_crowd)
_parity_small("L_SHAPE small r=5, 5 placed", L_SHAPE, 5, placed_crowd)
_parity_small("HEXAGON small r=6, 1 placed", HEXAGON, 6, placed_1)

# Small disc with bias
_parity_small("RECT small r=5, bias=ne", RECT, 5, [],
                bias_target=(200, 0), bias_weight=2.0)
_parity_small("L_SHAPE small r=5, bias=sw", L_SHAPE, 5, [],
                bias_target=(0, 200), bias_weight=2.0)


# ============================================================
print("\n=== End-to-end parity through _nest_discs_polygon ===")
# ============================================================

def _nest_e2e_polygon(label, pads, material, polygon, settings):
    """Run _nest_discs_polygon with python and numpy paths, compare placements."""
    saved = svg_engine._HAS_NUMPY
    try:
        svg_engine._HAS_NUMPY = False
        placed_py, fp_py, ft_py = _nest_discs_polygon(pads, material, settings, polygon)
        svg_engine._HAS_NUMPY = True
        placed_np, fp_np, ft_np = _nest_discs_polygon(pads, material, settings, polygon)
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

# Small fixed set on each polygon shape
pads_small_set = [{'size': 20.0, 'qty': 3}]
for name, poly in [("RECT", RECT), ("TRIANGLE", TRIANGLE), ("HEXAGON", HEXAGON),
                    ("L_SHAPE", L_SHAPE), ("U_SHAPE", U_SHAPE)]:
    _nest_e2e_polygon(f"{name}: 3 felt pads 20mm", pads_small_set, "felt", poly, settings)

# Mix of large + small (triggers both scan functions)
pads_mixed = [{'size': float(s), 'qty': 1} for s in (12, 14, 16, 18, 20, 22, 24, 26)]
_nest_e2e_polygon("RECT mixed sizes 12-26", pads_mixed, "felt", RECT, settings)
_nest_e2e_polygon("L_SHAPE mixed sizes 12-26", pads_mixed, "felt", L_SHAPE, settings)
_nest_e2e_polygon("HEXAGON mixed sizes 12-26", pads_mixed, "felt", HEXAGON, settings)

# Card material on a paper-shaped polygon (approx A4 with a cutout corner)
A4_CUT = [(0.0, 0.0), (210.0, 0.0), (210.0, 250.0),
          (160.0, 250.0), (160.0, 297.0), (0.0, 297.0)]
pads_card = [{'size': float(s), 'qty': 1} for s in range(7, 30)]
_nest_e2e_polygon("Card on A4-with-cut polygon", pads_card, "card", A4_CUT, settings)

# All radial biases
for bias in ("center", "nw", "ne", "sw", "se"):
    settings["edge_bias"] = bias
    _nest_e2e_polygon(f"bias={bias}, 4 pads on RECT", [{'size': 20.0, 'qty': 4}],
                       "felt", RECT, settings)

# Cardinal biases (still use the polygon scan)
for bias in ("n", "s", "e", "w"):
    settings["edge_bias"] = bias
    _nest_e2e_polygon(f"bias={bias} on RECT, 4 pads", [{'size': 20.0, 'qty': 4}],
                       "felt", RECT, settings)

# Max-quantity fill
settings["edge_bias"] = "center"
pads_max = [{'size': 25.0, 'qty': 2}, {'size': 12.0, 'qty': 'max'}]
_nest_e2e_polygon("RECT max-qty fill", pads_max, "felt", RECT, settings)


# ============================================================
print("\n=== Timing: realistic polygon scrap scenario (informational) ===")
# ============================================================

# An irregular scrap with a typical card pad set
SCRAP_POLY = [(0.0, 0.0), (180.0, 0.0), (200.0, 20.0), (200.0, 200.0),
              (50.0, 220.0), (0.0, 180.0)]
pads_realworld = [{'size': float(s), 'qty': 1} for s in range(7, 40)]

settings["edge_bias"] = "center"

svg_engine._HAS_NUMPY = False
t0 = time.perf_counter()
placed_py, _, _ = _nest_discs_polygon(pads_realworld, "card", settings, SCRAP_POLY)
t_py = time.perf_counter() - t0

svg_engine._HAS_NUMPY = True
t0 = time.perf_counter()
placed_np, _, _ = _nest_discs_polygon(pads_realworld, "card", settings, SCRAP_POLY)
t_np = time.perf_counter() - t0

print(f"  Python: {t_py:.2f}s ({len(placed_py)} pads placed)")
print(f"  Numpy:  {t_np:.2f}s ({len(placed_np)} pads placed)")
if t_np > 0:
    print(f"  Speedup: {t_py / t_np:.1f}x")

check("numpy at least 2x faster than Python on polygon scrap scenario",
      t_py / max(t_np, 0.001) >= 2.0)
check(f"timing-test polygon placements identical ({len(placed_py)} each)",
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
