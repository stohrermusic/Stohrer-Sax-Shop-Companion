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
print("\n=== 'mid' / 'int' placement with and without a hole (both engines) ===")
# ============================================================
# Matt (2026-10-26): "mid" (settings mode "centered") puts the label halfway
# between the center hole's edge and the cut edge, or halfway between the
# CENTER and the cut edge when the pad gets no hole. Until then a no-hole
# pad used a phantom 1.75 mm hole radius, and the G-code engine placed
# "int"/"mid" labels from the PRESET hole size even on pads too small for
# a hole, so the two outputs disagreed there. Both engines now resolve the
# disc's hole the same way: preset size only when pad >= min_hole_size.


def gcode_label_center_y_h(gcode_text, hole_dia):
    """gcode_label_center_y for a given hole size (0 = no hole circle)."""
    gc_cy = SHEET_H - DISC_CY
    body = gcode_text.split("\n; return to origin")[0]
    ys = []
    for m in re.finditer(r'^G1 X([0-9.-]+)Y([0-9.-]+)', body, re.M):
        x, y = float(m.group(1)), float(m.group(2))
        dist = math.hypot(x - DISC_CX, y - gc_cy)
        if abs(dist - DISC_R) > 0.6 and (hole_dia <= 0 or abs(dist - hole_dia / 2) > 0.6):
            ys.append(y)
    assert ys, "no engraving stroke points found in gcode"
    return (min(ys) + max(ys)) / 2


# Darted leather: "out" measures from the dart valley (the star's inner
# radius), not the cut edge, so one value sits the label just inside the
# valley at every size (Matt, 2026-10-26). From the valley the value is the
# gap to the TOP of the number (its middle sits value + font/2 inside), so
# 0 means touching rather than half into the darts; from the cut edge it is
# still the distance to the middle. Both engines must agree.
from svg_engine import dart_valley_radius, get_disc_diameter  # noqa: E402
from config import get_sizing_for_size, get_dart_settings_for_size  # noqa: E402
for pad, from_valley in ((12.0, True), (30.0, True), (30.0, False)):
    settings = copy.deepcopy(DEFAULT_SETTINGS)
    settings["engraving_on"] = True
    settings["dart_engraving_on"] = True
    settings["dart_threshold"] = 100.0            # every pad darted
    settings["darted_out_from_valley"] = from_valley
    settings["engraving_location"]["darted_leather"] = {"mode": "from_outside", "value": 3.0}
    r = get_disc_diameter(pad, "leather", settings) / 2
    sizing = get_sizing_for_size(pad, settings)
    valley = dart_valley_radius(pad, r, sizing, get_dart_settings_for_size(pad, settings), settings)
    placed = [(pad, DISC_CX, DISC_CY, r)]
    with tempfile.TemporaryDirectory() as td:
        svg_path = os.path.join(td, "dart.svg")
        gc_path = os.path.join(td, "dart.gcode")
        generate_svg_from_placed(placed, "leather", SHEET_W, SHEET_H, svg_path, 3.0, settings)
        generate_gcode_from_placed(placed, "leather", SHEET_W, SHEET_H, gc_path, 3.0, settings)
        with open(svg_path) as f:
            svg_text = f.read()
        with open(gc_path) as f:
            gc_text = f.read()
    font = float(re.search(r'font-size="([0-9.]+)mm"', svg_text).group(1))
    svg_center = float(re.search(r'<text[^>]*\sy="([0-9.]+)mm"', svg_text).group(1)) - 0.35 * font
    svg_off = DISC_CY - svg_center
    gc_cy = SHEET_H - DISC_CY
    # The engraving block is everything between the first and second
    # '; Cut @' lines (hole and outline cuts follow), so no geometric
    # filtering is needed to tell label strokes from dart strokes.
    engraving_block = gc_text.split('; Cut @')[1]
    ys = [float(mm.group(2)) for mm in re.finditer(r'^G1 X([0-9.-]+)Y([0-9.-]+)', engraving_block, re.M)]
    gc_off = (min(ys) + max(ys)) / 2 - gc_cy
    ref = valley if from_valley else r
    expected = ref - 3.0 - (font / 2 if from_valley else 0)
    check(f"darted {pad} mm, from_valley={from_valley}: svg label {svg_off:.3f} from center == "
          f"{expected:.3f} ({'valley' if from_valley else 'cut edge'} {ref:.3f} - 3.0"
          f"{' - font/2' if from_valley else ''})",
          abs(svg_off - expected) < 0.05)
    if from_valley:
        # The burned digits' top edge: value from the valley, less the stroke
        # font's own margin inside its box (0.05 x font), never more.
        top = max(ys) - gc_cy
        check(f"darted {pad} mm: gcode digits' top is 3.0 mm inside the valley "
              f"(gap {valley - top:.3f})", 3.0 <= valley - top <= 3.0 + 0.1 * font)
    check(f"darted {pad} mm, from_valley={from_valley}: gcode label agrees with svg "
          f"(err {abs(gc_off - svg_off):.3f} mm)", abs(gc_off - svg_off) < 0.3)

# ============================================================
print("\n=== Dart tip over the label (both engines) ===")
# ============================================================
# Matt (2026-10-26): a dart tip always sits at 12 o'clock, over the size
# label, so the number gets a tip's room rather than a notch. Before this
# the darts started at 3 o'clock, which put a notch over the label on every
# size whose dart count is 2 mod 4 (10 and 12 under his preset).
from svg_engine import DART_TIP_ANGLE  # noqa: E402

# Engine helpers: the G-code star (Y-up, tip at -DART_TIP_ANGLE) is the
# mirror image of the SVG star (Y-down, tip at DART_TIP_ANGLE), any count.
for n in (12, 13, 14):
    svg_pts = parse_star_path(calculate_star_path(CX, CY, OUTER_R, INNER_R, num_points=n,
                                                  shape_factor=0.5, tip_angle=DART_TIP_ANGLE))
    gc_pts = _generate_star_points(CX, CY, OUTER_R, INNER_R, n, 0.5, tip_angle=-DART_TIP_ANGLE)
    mirrored = [(x, 2 * CY - y) for x, y in reversed(gc_pts)]
    dev = max(math.hypot(sx - gx, sy - gy) for (sx, sy), (gx, gy) in zip(svg_pts, mirrored))
    check(f"{n} darts: gcode star is the svg star mirrored (dev {dev:.4f})", dev < 0.002)


def tip_over_center(points, cx, cy, up):
    """Distance from the centre of the outline point straight above it."""
    above = [(x, y) for x, y in points if abs(x - cx) < 1e-3 and ((y > cy) if up else (y < cy))]
    assert above, "no outline point straight above the centre"
    return max(math.hypot(x - cx, y - cy) for x, y in above)


settings = copy.deepcopy(DEFAULT_SETTINGS)
settings["dart_threshold"] = 100.0
for mat_cfg in settings["gcode_settings"].values():
    mat_cfg["kerf_width"] = 0.0
counts_seen = set()
for pad in (7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 20.0, 25.0, 30.0, 39.5):
    r = get_disc_diameter(pad, "leather", settings) / 2
    cx, cy = SHEET_W / 2, SHEET_H / 2
    with tempfile.TemporaryDirectory() as td:
        svg_path = os.path.join(td, "tip.svg")
        gc_path = os.path.join(td, "tip.gcode")
        generate_svg_from_placed([(pad, cx, cy, r)], "leather", SHEET_W, SHEET_H, svg_path, 0, settings)
        generate_gcode_from_placed([(pad, cx, cy, r)], "leather", SHEET_W, SHEET_H, gc_path, 0, settings)
        svg_text = open(svg_path).read()
        gc_text = open(gc_path).read()
    svg_star = parse_star_path(re.search(r'<path[^>]*\sd="([^"]+)"', svg_text).group(1))
    counts_seen.add(((len(svg_star) - 1) // 8) % 4)
    svg_top = tip_over_center(svg_star, cx, cy, up=False)
    cut_block = gc_text.split('; Cut @')[-1].split("\n; return to origin")[0]
    gc_star = [(float(m.group(1)), float(m.group(2)))
               for m in re.finditer(r'^G[01] X([0-9.-]+) ?Y([0-9.-]+)', cut_block, re.M)]
    gc_top = tip_over_center(gc_star, cx, SHEET_H - cy, up=True)
    check(f"darted {pad} mm: a dart tip over the label in svg ({svg_top:.3f}) and "
          f"gcode ({gc_top:.3f}), cut radius {r:.3f}",
          abs(svg_top - r) < 0.01 and abs(gc_top - r) < 0.01)
check(f"the sizes above include dart counts both 0 and 2 mod 4 (saw {sorted(counts_seen)})",
      counts_seen == {0, 2})

HOLE_SCENARIOS = (
    # label, hole option (mm), min_hole_size, hole radius the label should use
    ("hole 6.0, pad 40 gets it", 6.0, 16.5, 3.0),
    ("hole option None", 0.0, 16.5, 0.0),
    ("hole 6.0 but pad below min_hole_size", 6.0, 50.0, 0.0),
)
for label, hole_dia, min_hole, exp_hole_r in HOLE_SCENARIOS:
    for mode, value, expected_offset in (
            ("centered", 0, (DISC_R + exp_hole_r) / 2),
            ("from_inside", 4.0, exp_hole_r + 4.0)):
        settings = copy.deepcopy(DEFAULT_SETTINGS)
        settings["engraving_on"] = True
        settings["min_hole_size"] = min_hole
        settings["engraving_location"]["felt"] = {"mode": mode, "value": value}
        placed = [(PAD_SIZE, DISC_CX, DISC_CY, DISC_R)]
        with tempfile.TemporaryDirectory() as td:
            svg_path = os.path.join(td, "mid.svg")
            gc_path = os.path.join(td, "mid.gcode")
            generate_svg_from_placed(placed, "felt", SHEET_W, SHEET_H,
                                     svg_path, hole_dia, settings)
            generate_gcode_from_placed(placed, "felt", SHEET_W, SHEET_H,
                                       gc_path, hole_dia, settings)
            with open(svg_path) as f:
                svg_center = svg_label_center_y(f.read())
            with open(gc_path) as f:
                # Only a hole that is actually cut leaves a circle to filter out.
                cut_hole = hole_dia if PAD_SIZE >= min_hole else 0
                gc_center = gcode_label_center_y_h(f.read(), cut_hole)
        svg_off = DISC_CY - svg_center             # above center, Y-down
        gc_off = gc_center - (SHEET_H - DISC_CY)   # above center, Y-up
        check(f"{mode}, {label}: svg label {svg_off:.3f} mm from center, "
              f"expected {expected_offset:.3f}", abs(svg_off - expected_offset) < 0.05)
        check(f"{mode}, {label}: gcode label agrees with svg "
              f"(err {abs(gc_off - svg_off):.3f} mm)", abs(gc_off - svg_off) < 0.3)



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
