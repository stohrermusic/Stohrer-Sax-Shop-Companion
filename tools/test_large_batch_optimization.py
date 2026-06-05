"""
Sanity tests for the scrap-mode large-batch optimizer.

Verifies that:
  - try_nest_partial(optimize=True) returns same or more pads than the
    default greedy on representative large pad sets.
  - Reproducible across runs (seeded RNG).
  - Doesn't crash on edge cases (empty pads, tiny scraps, polygon shapes).
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from svg_engine import try_nest_partial, _multistart_nest  # noqa: E402
from config import DEFAULT_SETTINGS  # noqa: E402

passed = 0
failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS  {name}")
        passed += 1
    else:
        print(f"  FAIL  {name}")
        failed += 1


# A representative large pad set (≥ 75 pads, mixed sizes) on a Falcon-
# bed-sized rectangle.
large_pads = [
    {'size': float(s), 'qty': 10}
    for s in (8.0, 10.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0, 26.0)
]  # 100 pads total
settings = DEFAULT_SETTINGS.copy()
settings['edge_bias'] = 'center'

# Default greedy
t0 = time.perf_counter()
placed_default, remaining_default, any_default = try_nest_partial(
    large_pads, 'felt', 400.0, 415.0, settings, optimize=False)
t_default = time.perf_counter() - t0

# Optimized multistart
t0 = time.perf_counter()
placed_opt, remaining_opt, any_opt = try_nest_partial(
    large_pads, 'felt', 400.0, 415.0, settings, optimize=True)
t_opt = time.perf_counter() - t0

print(f"\nDefault: {len(placed_default)} placed in {t_default:.2f}s")
print(f"Optimized: {len(placed_opt)} placed in {t_opt:.2f}s")
print(f"Improvement: +{len(placed_opt) - len(placed_default)} pads "
      f"({100 * (len(placed_opt) - len(placed_default)) / max(len(placed_default), 1):.1f}%)")

check("optimize=True placed >= default",
       len(placed_opt) >= len(placed_default))
check("optimize=True placed at least something", any_opt)
check("optimize=True returned consistent remaining count",
       sum(p['qty'] for p in remaining_opt) == 100 - len(placed_opt))

# Reproducibility: running twice should give identical results
placed_opt2, _, _ = try_nest_partial(
    large_pads, 'felt', 400.0, 415.0, settings, optimize=True)
check("optimize=True is reproducible (same RNG seed)",
       len(placed_opt) == len(placed_opt2))

# Tight case: pad set DENSER than what fits easily on the sheet.
# This is where multistart actually pays off vs. default greedy.
dense_pads = [
    {'size': float(s), 'qty': 15}
    for s in (12.0, 16.0, 20.0, 24.0, 28.0)
]  # 75 pads
t0 = time.perf_counter()
placed_dense_d, _, _ = try_nest_partial(
    dense_pads, 'felt', 250.0, 250.0, settings, optimize=False)
t_dd = time.perf_counter() - t0
t0 = time.perf_counter()
placed_dense_o, _, _ = try_nest_partial(
    dense_pads, 'felt', 250.0, 250.0, settings, optimize=True)
t_do = time.perf_counter() - t0
print("\nDense case (75 pads on 250x250mm):")
print(f"  default:   {len(placed_dense_d)} placed in {t_dd:.2f}s")
print(f"  optimized: {len(placed_dense_o)} placed in {t_do:.2f}s")
print(f"  delta: +{len(placed_dense_o) - len(placed_dense_d)} pads")
check("dense case: optimize >= default",
       len(placed_dense_o) >= len(placed_dense_d))

# Polygon path
A4 = [(0.0, 0.0), (210.0, 0.0), (210.0, 297.0), (0.0, 297.0)]
placed_poly_default, _, _ = try_nest_partial(
    large_pads, 'felt', 210.0, 297.0, settings,
    polygon=A4, optimize=False)
placed_poly_opt, _, _ = try_nest_partial(
    large_pads, 'felt', 210.0, 297.0, settings,
    polygon=A4, optimize=True)

print(f"\nA4 polygon default: {len(placed_poly_default)} placed")
print(f"A4 polygon optimized: {len(placed_poly_opt)} placed")
check("polygon optimize=True placed >= polygon default",
       len(placed_poly_opt) >= len(placed_poly_default))

# Edge case: small pad set (below threshold) — optimization still works
# but doesn't add value
small_pads = [{'size': 18.0, 'qty': 5}]
placed_small, _, _ = try_nest_partial(
    small_pads, 'felt', 400.0, 415.0, settings, optimize=True)
check("small pad set with optimize=True still places pads",
       len(placed_small) >= 1)

# Edge case: empty pads
placed_empty, remaining_empty, any_empty = try_nest_partial(
    [], 'felt', 400.0, 415.0, settings, optimize=True)
check("empty pads with optimize=True returns no placements",
       len(placed_empty) == 0)
check("empty pads with optimize=True returns no remaining",
       len(remaining_empty) == 0)
check("empty pads with optimize=True any_placed is False",
       any_empty is False)

# ---------------------------------------------------------------------
# Variable-shape scrap safety.
#
# In scrap mode the polygon is re-captured per scrap, so the shape is
# different on every Generate / Frame & Cut click, while the session's
# large-batch `optimize` flag stays on for the whole run. The optimizer
# must therefore be fully stateless across calls — each scrap has to nest
# into ONLY its own current shape, with nothing leaking from the prior
# scrap. These prove it.
# ---------------------------------------------------------------------

def _point_in_polygon(x, y, poly):
    """Ray-casting point-in-polygon (matches the nester's own test)."""
    inside = False
    n = len(poly)
    j = n - 1
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[j]
        if ((yi > y) != (yj > y)) and \
                (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def _round_placed(placed):
    return [(round(s, 6), round(cx, 6), round(cy, 6), round(r, 6))
            for (s, cx, cy, r) in placed]


def _all_inside(placed, poly):
    return all(_point_in_polygon(cx, cy, poly) for (_s, cx, cy, _r) in placed)


shape_square = [(0.0, 0.0), (200.0, 0.0), (200.0, 200.0), (0.0, 200.0)]
shape_tri = [(0.0, 0.0), (260.0, 0.0), (0.0, 260.0)]
shape_pent = [(0.0, 40.0), (120.0, 0.0), (240.0, 60.0),
              (200.0, 220.0), (30.0, 200.0)]
scrap_pads = [{'size': float(s), 'qty': 12}
              for s in (10.0, 14.0, 18.0, 22.0)]  # 48 pads

# (a) Order independence: nest the pentagon standalone, then nest it
# again AFTER running a square and a triangle through the optimizer.
# Identical placements prove no state carried over between scraps.
pent_standalone, _, _ = try_nest_partial(
    scrap_pads, 'felt', 240.0, 220.0, settings, polygon=shape_pent, optimize=True)
sq_placed, _, _ = try_nest_partial(
    scrap_pads, 'felt', 200.0, 200.0, settings, polygon=shape_square, optimize=True)
tri_placed, _, _ = try_nest_partial(
    scrap_pads, 'felt', 260.0, 260.0, settings, polygon=shape_tri, optimize=True)
pent_after_others, _, _ = try_nest_partial(
    scrap_pads, 'felt', 240.0, 220.0, settings, polygon=shape_pent, optimize=True)

check("optimizer stateless across scraps (pentagon order-independent)",
      _round_placed(pent_standalone) == _round_placed(pent_after_others))

# (b) Every placed disc lands inside the shape it was nested into, across
# three distinct shapes. A leaked placement from a prior shape would fall
# outside.
check("square scrap: all pads inside the square", _all_inside(sq_placed, shape_square))
check("triangle scrap: all pads inside the triangle", _all_inside(tri_placed, shape_tri))
check("pentagon scrap: all pads inside the pentagon",
      _all_inside(pent_standalone, shape_pent))

# (c) Full variable-shape session: 100 pads, optimize on, three different
# scraps in sequence (feeding each scrap's remainder to the next, exactly
# like a real session). Pads are conserved and nothing lands off-shape.
session_pads = [{'size': float(s), 'qty': 10}
                for s in (8.0, 10.0, 12.0, 14.0, 16.0,
                          18.0, 20.0, 22.0, 24.0, 26.0)]  # 100 pads
total_start = sum(p['qty'] for p in session_pads)
scrap_shapes = [
    (shape_square, 200.0, 200.0),
    (shape_tri, 260.0, 260.0),
    (shape_pent, 240.0, 220.0),
]
session_remaining = session_pads
session_total_placed = 0
session_off_shape = 0
for poly, w, h in scrap_shapes:
    s_placed, session_remaining, _ = try_nest_partial(
        session_remaining, 'felt', w, h, settings, polygon=poly, optimize=True)
    session_total_placed += len(s_placed)
    if not _all_inside(s_placed, poly):
        session_off_shape += 1
session_remaining_total = sum(p['qty'] for p in session_remaining)
print(f"\nVariable-shape session: placed {session_total_placed}, "
      f"{session_remaining_total} remaining of {total_start}")
check("variable-shape session conserves pads (placed + remaining == start)",
      session_total_placed + session_remaining_total == total_start)
check("variable-shape session: no pad placed off its scrap shape",
      session_off_shape == 0)

print(f"\n=== Summary: {passed} passed, {failed} failed ===")
assert all([try_nest_partial, _multistart_nest])  # silence ruff
sys.exit(0 if failed == 0 else 1)
