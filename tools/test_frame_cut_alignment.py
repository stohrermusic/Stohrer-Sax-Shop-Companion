"""
Frame & Cut alignment: the cut must land where the frame showed it.

The framing pass traces the scrap OUTLINE, Y-flipped with the outline's
height, and the cut is streamed into the same work frame (same G92). A
camera-captured polygon is INSET from that outline, and the cut generator
flips with the height of whatever polygon it's handed — so handing it the
inset polygon shifted every disc `inset` mm toward the machine front of
where the frame showed it (v2.5 through v2.7; 5 mm on a 5 mm inset, which
is the front row of pads on the scrap's edge). Found 2026-09-20.

Fix: on_frame_and_cut passes flip_height_mm = the outline's height. This
file checks the geometry at the engine level with the same G92 arithmetic
as on_frame_and_cut, including the negative control, and checks that the
call site actually passes the override.

Headless: no Tk. Run:
    python tools/test_frame_cut_alignment.py
"""
import inspect
import os
import re
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DEFAULT_SETTINGS  # noqa: E402
from gcode_engine import generate_gcode_from_placed, generate_polygon_framing_gcode  # noqa: E402

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
INSET = 5.0

# A 100 x 80 mm scrap on the bed, machine Y-UP mm, as camera capture hands it
# over: the outline, and the same shape inset by 5 mm.
outline_up = [(200.0, 100.0), (300.0, 100.0), (300.0, 180.0), (200.0, 180.0)]
inset_up = [(205.0, 105.0), (295.0, 105.0), (295.0, 175.0), (205.0, 175.0)]
SCRAP_CENTER = (250.0, 140.0)

# _set_custom_polygon_from_y_up: both polygons normalised to the OUTLINE's
# bbox-min-x and max-y, then Y-flipped into storage (Y-down).
min_x = min(p[0] for p in outline_up)
max_y = max(p[1] for p in outline_up)
custom_polygon = [(x - min_x, max_y - y) for x, y in inset_up]        # nesting boundary
custom_polygon_outline = [(x - min_x, max_y - y) for x, y in outline_up]  # framing trace

# One disc at the centre of the inset polygon, in storage coordinates.
cx = sum(p[0] for p in custom_polygon) / 4
cy = sum(p[1] for p in custom_polygon) / 4
placed = [(20.0, cx, cy, 10.0)]


def cut_center_in_work_frame(flip_height_mm):
    """Generate the cut and return the disc centre in the file's work frame."""
    fd, path = tempfile.mkstemp(suffix=".gcode")
    os.close(fd)
    try:
        generate_gcode_from_placed(placed, 'felt', 393.7, 393.7, path, 0, settings,
                                   polygon=custom_polygon, flip_height_mm=flip_height_mm)
        lines = open(path).read().splitlines()
    finally:
        os.remove(path)
    i0 = next(i for i, l in enumerate(lines) if l.startswith("; Layer C00"))
    # Stop before the footer: its return-to-origin move is a G1 X0Y0 too.
    i1 = next(i for i, l in enumerate(lines) if i > i0 and l.startswith("; return to origin"))
    pts = [(float(m.group(1)), float(m.group(2)))
           for l in lines[i0:i1] for m in [re.match(r"G1 X([-\d.]+)Y([-\d.]+)", l)] if m]
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return ((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2)


# The framing frame and G92, exactly as on_frame_and_cut computes them: the
# outline flipped with its own height, G92 at its leftmost-lowest vertex,
# and the head physically jogged to that corner of the scrap.
ymax_storage = max(p[1] for p in custom_polygon_outline)
flipped_outline = [(x, ymax_storage - y) for x, y in custom_polygon_outline]
lb = min(flipped_outline, key=lambda p: p[0] ** 2 + p[1] ** 2)
jog = (200.0, 100.0)   # the scrap's physical bottom-left corner on the bed


def to_machine(x, y):
    return (jog[0] + (x - lb[0]), jog[1] + (y - lb[1]))


frame_pts = [(float(m.group(1)), float(m.group(2)))
             for l in generate_polygon_framing_gcode(custom_polygon_outline)
             for m in [re.match(r"G[01] X([-\d.]+) Y([-\d.]+)", l)] if m]
frame_m = [to_machine(x, y) for x, y in frame_pts]
fx = (min(p[0] for p in frame_m), max(p[0] for p in frame_m))
fy = (min(p[1] for p in frame_m), max(p[1] for p in frame_m))
check("framing pass traces the scrap where it physically is",
      abs(fx[0] - 200) < 1e-6 and abs(fx[1] - 300) < 1e-6
      and abs(fy[0] - 100) < 1e-6 and abs(fy[1] - 180) < 1e-6)

# --- the fix: flip the cut with the outline's height --------------------
good = to_machine(*cut_center_in_work_frame(ymax_storage))
print(f"  disc centre with the outline flip : ({good[0]:.2f}, {good[1]:.2f})  expected {SCRAP_CENTER}")
check("cut disc lands at the centre of the framed scrap (x)", abs(good[0] - SCRAP_CENTER[0]) < 0.05)
check("cut disc lands at the centre of the framed scrap (y)", abs(good[1] - SCRAP_CENTER[1]) < 0.05)

# --- negative control: the pre-fix default (flip with the inset) ----------
old = to_machine(*cut_center_in_work_frame(None))
print(f"  disc centre with the inset flip    : ({old[0]:.2f}, {old[1]:.2f})  = {SCRAP_CENTER[1] - old[1]:.2f} mm toward the front")
check("without the override the disc lands `inset` mm toward the machine front (the v2.5-v2.7 bug)",
      abs((SCRAP_CENTER[1] - old[1]) - INSET) < 0.05 and abs(old[0] - SCRAP_CENTER[0]) < 0.05)

# --- the call site passes the override ------------------------------------
import main  # noqa: E402
src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
check("on_frame_and_cut passes flip_height_mm to the cut generator",
      "flip_height_mm=flip_height_mm" in src)
check("on_frame_and_cut derives it from the outline, falling back to the polygon",
      "self.custom_polygon_outline or self.custom_polygon" in src.split("flip_height_mm =")[0])
check("on_frame_and_cut keeps the streamed cut", "_keep_last_cut_gcode(" in src)

print(f"\n=== Summary: {passed} passed, {failed} failed ===")
sys.exit(0 if failed == 0 else 1)
