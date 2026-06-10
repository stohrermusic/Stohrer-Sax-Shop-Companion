"""
SVG <-> G-code parity tests.

The two engines render the same discs independently, and history shows they
drift (the dart wave math and the engraving Y-placement both diverged before
the June 2026 review). These tests pin the contract: what the SVG/preview
shows is what the G-code cuts.

Covers:
  1. Dart wave parity: _generate_star_points (gcode) vs calculate_star_path
     (svg) across the full shape-factor spectrum.
  2. Pad engraving label placement parity through the real
     generate_*_from_placed pipelines, for all three placement modes.
  3. Engine purity: gcode_engine/svg_engine importable without tkinter.
"""

import sys
import os
import re
import copy
import math
import tempfile
import subprocess

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS
from svg_engine import calculate_star_path, generate_svg_from_placed
from gcode_engine import _generate_star_points, generate_gcode_from_placed

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


# ============================================================
print("\n=== Dart wave parity (svg path vs gcode points) ===")
# ============================================================

def parse_star_path(path_d):
    """Parse the 'M x y L x y ... Z' path emitted by calculate_star_path."""
    nums = [float(t) for t in path_d.replace("M", " ").replace("L", " ")
            .replace("Z", " ").split()]
    return list(zip(nums[0::2], nums[1::2]))


CX, CY, OUTER_R, INNER_R, NUM_POINTS = 50.0, 50.0, 20.0, 16.0, 12

for sf in (0.0, 0.25, 0.5, 0.75, 1.0):
    svg_pts = parse_star_path(
        calculate_star_path(CX, CY, OUTER_R, INNER_R,
                            num_points=NUM_POINTS, shape_factor=sf))
    gc_pts = _generate_star_points(CX, CY, OUTER_R, INNER_R, NUM_POINTS, sf)

    check(f"shape_factor={sf}: same number of points",
          len(svg_pts) == len(gc_pts))

    # Pointwise: both engines walk the same theta sequence, so points
    # must coincide (svg path is formatted to 3 decimals -> 0.002 tol).
    max_dev = max(math.hypot(sx - gx, sy - gy)
                  for (sx, sy), (gx, gy) in zip(svg_pts, gc_pts))
    check(f"shape_factor={sf}: max point deviation < 0.002mm "
          f"(got {max_dev:.5f})", max_dev < 0.002)


# ============================================================
print("\n=== Pad engraving placement parity (full pipeline) ===")
# ============================================================

SHEET_W, SHEET_H = 100.0, 100.0
PAD_SIZE, DISC_CX, DISC_CY, DISC_R = 40.0, 50.0, 40.0, 19.625
HOLE_DIA = 6.0
FONT_SIZE = 3.0  # DEFAULT_SETTINGS felt engraving font size


def svg_label_center_y(svg_text):
    """Visual center of the engraved label in the SVG (Y-down, mm).

    SVG anchors text at its baseline; the engines place the baseline at
    visual_center + 0.35 * font_size.
    """
    texts = re.findall(r'<text[^>]*\sy="([0-9.]+)mm"', svg_text)
    assert len(texts) == 1, f"expected exactly 1 text element, got {len(texts)}"
    return float(texts[0]) - 0.35 * FONT_SIZE


def gcode_label_center_y(gcode_text):
    """Visual center of the engraved label in the G-code (Y-up, mm).

    Engraving stroke points are the ones NOT lying on the cut circle or
    the hole circle (allowing for kerf offset via a 0.6mm margin).
    """
    gc_cy = SHEET_H - DISC_CY  # disc center in flipped coords
    # Only the shape-drawing body counts: the footer contains a G1
    # return-to-origin travel move that isn't part of any shape.
    body = gcode_text.split("\n; return to origin")[0]
    ys = []
    for m in re.finditer(r'^G1 X([0-9.-]+)Y([0-9.-]+)', body, re.M):
        x, y = float(m.group(1)), float(m.group(2))
        dist = math.hypot(x - DISC_CX, y - gc_cy)
        if abs(dist - DISC_R) > 0.6 and abs(dist - HOLE_DIA / 2) > 0.6:
            ys.append(y)
    assert ys, "no engraving stroke points found in gcode"
    return (min(ys) + max(ys)) / 2


for mode, value in (("from_outside", 2.5), ("from_inside", 4.0), ("centered", 0)):
    settings = copy.deepcopy(DEFAULT_SETTINGS)
    settings["engraving_on"] = True
    settings["engraving_location"]["felt"] = {"mode": mode, "value": value}
    placed = [(PAD_SIZE, DISC_CX, DISC_CY, DISC_R)]

    with tempfile.TemporaryDirectory() as td:
        svg_path = os.path.join(td, "parity.svg")
        gc_path = os.path.join(td, "parity.gcode")
        generate_svg_from_placed(placed, "felt", SHEET_W, SHEET_H,
                                 svg_path, HOLE_DIA, settings)
        generate_gcode_from_placed(placed, "felt", SHEET_W, SHEET_H,
                                   gc_path, HOLE_DIA, settings)
        with open(svg_path) as f:
            svg_center = svg_label_center_y(f.read())
        with open(gc_path) as f:
            gc_center = gcode_label_center_y(f.read())

    expected_gc = SHEET_H - svg_center  # Y-flip
    err = abs(expected_gc - gc_center)
    check(f"{mode}: gcode label center matches flipped svg center "
          f"(err {err:.3f}mm)", err < 0.3)

    # And the label must be on the correct SIDE of the disc center
    # (the pre-fix bug mirrored it to the opposite side).
    svg_side = svg_center - DISC_CY          # negative = above center (Y-down)
    gc_side = gc_center - (SHEET_H - DISC_CY)  # positive = above center (Y-up)
    check(f"{mode}: label on same physical side of disc center",
          (svg_side < 0) == (gc_side > 0))


# ============================================================
print("\n=== Engine purity: importable without tkinter ===")
# ============================================================

repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
probe = ("import sys; sys.modules['tkinter'] = None; "
         "import svg_engine, gcode_engine; print('OK')")
result = subprocess.run([sys.executable, "-c", probe],
                        capture_output=True, text=True, cwd=repo)
check("svg_engine + gcode_engine import with tkinter blocked "
      f"(stderr: {result.stderr.strip()[:120] or 'none'})",
      result.returncode == 0 and "OK" in result.stdout)


# ============================================================
print(f"\n=== Total: {passed} passed, {failed} failed ===")
# ============================================================

sys.exit(0 if failed == 0 else 1)
