"""
Sanity tests for scrap-mode partial nesting (try_nest_partial).

Verifies that:
  - Results are deterministic (same input, same placement).
  - Edge cases don't crash (empty pads, small sets) and the remaining
    count is consistent with what was placed.
  - Nesting is stateless across scraps: in scrap mode the polygon is
    re-captured per scrap, so every call must nest into ONLY its own
    shape, with nothing leaking from the prior one.
  - A multi-scrap session with three different shapes conserves pads and
    never places one off its scrap.

History: this file replaced test_large_batch_optimization.py when the
multistart "large-batch optimization" was removed (2026-09-20). On real
lists it moved material usage by 0-1.3%: the gaps between big discs are
smaller than the smallest pad in the list, and adding small sizes fills
them with plain largest-first. The shape-safety cases below predate the
removal and still matter.

Run:
    python tools/test_scrap_partial_nest.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from svg_engine import try_nest_partial  # noqa: E402
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


settings = DEFAULT_SETTINGS.copy()
settings['edge_bias'] = 'center'

# A representative large pad set (100 pads, mixed sizes) on a Falcon-bed
# sized rectangle.
large_pads = [
    {'size': float(s), 'qty': 10}
    for s in (8.0, 10.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0, 26.0)
]

placed_a, remaining_a, any_a = try_nest_partial(large_pads, 'felt', 400.0, 415.0, settings)
placed_b, remaining_b, any_b = try_nest_partial(large_pads, 'felt', 400.0, 415.0, settings)
print(f"\n100 felt pads on 400x415: {len(placed_a)} placed")
check("places at least something", any_a)
check("deterministic: same input gives the same placement", placed_a == placed_b)
check("remaining count is consistent with what was placed",
      sum(p['qty'] for p in remaining_a) == 100 - len(placed_a))

# Edge case: small pad set
placed_small, _, _ = try_nest_partial([{'size': 18.0, 'qty': 5}], 'felt', 400.0, 415.0, settings)
check("small pad set places pads", len(placed_small) >= 1)

# Edge case: empty pads
placed_empty, remaining_empty, any_empty = try_nest_partial([], 'felt', 400.0, 415.0, settings)
check("empty pads returns no placements", len(placed_empty) == 0)
check("empty pads returns no remaining", len(remaining_empty) == 0)
check("empty pads any_placed is False", any_empty is False)


# ---------------------------------------------------------------------
# Variable-shape scrap safety.
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

# (a) Order independence: nest the pentagon standalone, then again AFTER
# a square and a triangle. Identical placements prove no state carried
# over between scraps.
pent_standalone, _, _ = try_nest_partial(
    scrap_pads, 'felt', 240.0, 220.0, settings, polygon=shape_pent)
sq_placed, _, _ = try_nest_partial(
    scrap_pads, 'felt', 200.0, 200.0, settings, polygon=shape_square)
tri_placed, _, _ = try_nest_partial(
    scrap_pads, 'felt', 260.0, 260.0, settings, polygon=shape_tri)
pent_after_others, _, _ = try_nest_partial(
    scrap_pads, 'felt', 240.0, 220.0, settings, polygon=shape_pent)

check("stateless across scraps (pentagon order-independent)",
      _round_placed(pent_standalone) == _round_placed(pent_after_others))

# (b) Every placed disc lands inside the shape it was nested into.
check("square scrap: all pads inside the square", _all_inside(sq_placed, shape_square))
check("triangle scrap: all pads inside the triangle", _all_inside(tri_placed, shape_tri))
check("pentagon scrap: all pads inside the pentagon",
      _all_inside(pent_standalone, shape_pent))

# (c) Full variable-shape session: 100 pads across three different scraps
# in sequence, feeding each scrap's remainder to the next, exactly like a
# real session. Pads are conserved and nothing lands off-shape.
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
        session_remaining, 'felt', w, h, settings, polygon=poly)
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
sys.exit(0 if failed == 0 else 1)
