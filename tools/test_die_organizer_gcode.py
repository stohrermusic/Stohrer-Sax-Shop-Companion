"""Die Organizer G-code: the bundled drawings, read and cut by this app.

The organizer was SVG only. Since 2026-11-09 the same two drawings also
come out as G-code (and through Frame & Cut), so nobody needs LightBurn for
them. This suite pins that path:

  - the reader takes rect, circle and ellipse with a translate or a pure
    rotation, and FAILS LOUDLY on anything else (a path, a group, a scale,
    a unit suffix, a missing plate), so a replacement drawing is never
    half read and mis-cut
  - every element of each drawing is in the G-code, once: an independent
    reading of the SVG (a regex pass written here, not the app's reader)
    agrees with the cut paths after the Y flip
  - openings first, the plate's edge last; kerf inward on the openings and
    outward on the edge, so both come out at drawn size; corners are true
    arcs; the bounding box is 230 x 330 plus the kerf; everything is on the
    sheet; the laser is off at the end
  - cut speed, power and passes come from the chosen material
  - a sheet smaller than the plate is refused, and so is a slot narrower
    than the kerf
  - through the real section: the material choice, the sheet fields, the
    Done box, and the section text no longer sending everyone to LightBurn

Engine checks are headless. The GUI ones self-skip without a display.
Isolated profile (SAXSHOP_CONFIG_DIR).
"""
import copy
import math
import os
import re
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-organizer-gcode-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import config  # noqa: E402
import gcode_engine as ge  # noqa: E402
import svg_engine as se  # noqa: E402

results = []


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        box_guard.settle()      # a question this check did not expect fails it
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}{box_guard.reset()}")
        traceback.print_exc()
        results.append(False)


PLATE_W, PLATE_H, MARGIN = 230.0, 330.0, 2.0
OUT = tempfile.mkdtemp(prefix="ssc-organizer-out-")


def settings(kerf=0.2, material="basswood"):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    s["gcode_settings"][material]["kerf_width"] = kerf
    return s


def asset(variant):
    return os.path.join(REPO, "tooling_assets", f"die_organizer_{variant}.svg")


def own_reading(variant):
    """The drawing read HERE, with regexes, into where each element should
    come out: (kind, is_plate, box), with box(d) the element's bounding box
    on the plate (y down) once its outline is moved d mm outward. Nothing
    from the app. Most elements are moved with translate(); two finger-grip
    rectangles and the oval in the upper plate are turned with a matrix()."""
    text = open(asset(variant), encoding="utf-8").read()
    vb = [float(v) for v in re.search(r'viewBox="([^"]+)"', text).group(1).split()]
    out = []
    for kind, attrs in re.findall(r'<(rect|circle|ellipse)\b([^>]*)/>', text):
        a = dict(re.findall(r'([\w:-]+)="([^"]*)"', attrs))
        tr = a["transform"]
        nums = [float(v) for v in re.findall(r'-?\d+\.?\d*(?:[eE]-?\d+)?', tr)]
        if tr.startswith("translate"):
            cos_t, sin_t, tx, ty = 1.0, 0.0, nums[0], nums[1]
        else:
            cos_t, sin_t, tx, ty = nums[0], nums[1], nums[4], nums[5]
        tx, ty = tx - vb[0], ty - vb[1]
        if kind == "rect":
            x, y, w, h = (float(a[k]) for k in ("x", "y", "width", "height"))
            lx, ly = x + w / 2, y + h / 2                    # the rectangle's own centre, turned onto the plate
            cx, cy = tx + cos_t * lx - sin_t * ly, ty + sin_t * lx + cos_t * ly

            def box(d, cx=cx, cy=cy, w=w, h=h, c=abs(cos_t), s=abs(sin_t)):
                ex, ey = (w / 2 + d) * c + (h / 2 + d) * s, (w / 2 + d) * s + (h / 2 + d) * c
                return cx - ex, cy - ey, cx + ex, cy + ey
            out.append(("rect", (w, h) == (PLATE_W, PLATE_H), box))
        elif kind == "circle":
            r = float(a["r"])
            out.append(("circle", False, lambda d, tx=tx, ty=ty, r=r: (tx - r - d, ty - r - d, tx + r + d, ty + r + d)))
        else:
            rx, ry = float(a["rx"]), float(a["ry"])
            # Half-extents of a turned ellipse; its parallel curve's are d more.
            ex = math.sqrt((rx * cos_t) ** 2 + (ry * sin_t) ** 2)
            ey = math.sqrt((rx * sin_t) ** 2 + (ry * cos_t) ** 2)
            out.append(("ellipse", False,
                        lambda d, tx=tx, ty=ty, ex=ex, ey=ey: (tx - ex - d, ty - ey - d, tx + ex + d, ty + ey + d)))
    return out


def strokes_of(path):
    """A G-code file's cut strokes, by layer: [(layer, [(x, y), ...]), ...]."""
    strokes, layer, pending = [], None, None
    x = y = None
    for ln in open(path).read().splitlines():
        if ln.startswith("; Layer "):
            layer = ln.split()[-1]
            continue
        if ln.startswith("; return to origin"):
            break                       # the footer's move home is not a cut
        if ln.startswith(";"):
            continue
        m = re.match(r'G([01])\b', ln)
        if not m:
            continue
        mx, my = re.search(r'X(-?\d+\.?\d*)', ln), re.search(r'Y(-?\d+\.?\d*)', ln)
        if mx is None and my is None:
            continue
        x = float(mx.group(1)) if mx else x
        y = float(my.group(1)) if my else y
        if m.group(1) == "0":
            pending = [(x, y)]
        else:
            if pending is not None:
                strokes.append((layer, pending))
                pending = None
            if strokes:
                strokes[-1][1].append((x, y))
    return [(lay, pts) for lay, pts in strokes if lay in ("C01", "C02")]


def bbox(pts):
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    return min(xs), min(ys), max(xs), max(ys)


def generate(variant, kerf=0.2, material="basswood", sheet=(300.0, 400.0), s=None):
    path = os.path.join(OUT, f"{variant}_{material}_{kerf}.gcode")
    info = ge.generate_die_organizer_gcode(variant, material, path, s or settings(kerf, material),
                                           sheet_width_mm=sheet[0], sheet_height_mm=sheet[1])
    return path, info


# ---------------------------------------------------------------- the reader

def reader_counts_and_plate():
    up, low = se.read_die_organizer_drawing("upper"), se.read_die_organizer_drawing("lower")
    for d in (up, low):
        assert (d["width"], d["height"]) == (PLATE_W, PLATE_H)
        assert d["outer"]["kind"] == "rect" and d["outer"]["r"] == 10.0
    kinds = lambda d: sorted(s["kind"] for s in d["inner"])   # noqa: E731
    assert kinds(up).count("rect") == 111 and kinds(up).count("circle") == 4 and kinds(up).count("ellipse") == 1
    assert kinds(low) == ["circle"] * 4
    ell = [s for s in up["inner"] if s["kind"] == "ellipse"][0]
    assert abs(math.degrees(ell["angle"]) - 89.145) < 0.01, "the oval is turned about 89 degrees"
    turned = [s for s in up["inner"] if s["kind"] == "rect" and abs(s["angle"]) > 1e-6]
    assert len(turned) == 2 and all((s["w"], s["h"]) == (60.604397, 15.494) for s in turned), \
        "two 60.6 x 15.5 rectangles are turned with a matrix()"


check("reader: both plates are 230 x 330 with one rounded edge; 111 slots, 4 holes and the turned oval in the upper",
      reader_counts_and_plate)


def write_svg(body, viewbox="0 0 100 100"):
    path = os.path.join(OUT, f"bad_{len(os.listdir(OUT))}.svg")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{viewbox}">{body}</svg>')
    return path


PLATE = '<rect transform="translate(50,50)" x="-50" y="-50" width="100" height="100" rx="5"/>'


def reader_fails_loudly():
    good = write_svg(PLATE + '<circle transform="translate(20,20)" r="2" cx="0" cy="0"/>')
    assert len(se.read_die_organizer_drawing(path=good)["inner"]) == 1, "fixture: the plain drawing must read"
    bad = {
        "a path": PLATE + '<path d="M0,0 L10,10"/>',
        "a group": PLATE + '<g><circle r="2" cx="5" cy="5"/></g>',
        "a line": PLATE + '<line x1="0" y1="0" x2="5" y2="5"/>',
        "a scale": PLATE + '<circle transform="scale(2)" r="2" cx="5" cy="5"/>',
        "a rotate()": PLATE + '<circle transform="rotate(30)" r="2" cx="5" cy="5"/>',
        "a matrix that scales": PLATE + '<circle transform="matrix(2,0,0,2,5,5)" r="2" cx="0" cy="0"/>',
        "a matrix that skews": PLATE + '<circle transform="matrix(1,0,0.5,1,5,5)" r="2" cx="0" cy="0"/>',
        "a mirror": PLATE + '<circle transform="matrix(1,0,0,-1,5,5)" r="2" cx="0" cy="0"/>',
        "a unit suffix": PLATE + '<circle transform="translate(5,5)" r="2mm" cx="0" cy="0"/>',
        "a circle with no radius": PLATE + '<circle transform="translate(5,5)" cx="0" cy="0"/>',
        "different rx and ry": '<rect transform="translate(50,50)" x="-50" y="-50" width="100" height="100" '
                               'rx="5" ry="2"/>',
        "no plate": '<circle transform="translate(5,5)" r="2" cx="0" cy="0"/>',
        "two plates": PLATE + PLATE,
        "a child element": PLATE + '<circle transform="translate(5,5)" r="2"><title>x</title></circle>',
    }
    for what, body in bad.items():
        try:
            se.read_die_organizer_drawing(path=write_svg(body))
        except se.OrganizerDrawingError:
            continue
        raise AssertionError(f"a drawing with {what} was read instead of refused")
    for what, vb in (("no viewBox", ""), ("a short viewBox", "0 0 100")):
        try:
            se.read_die_organizer_drawing(path=write_svg(PLATE, viewbox=vb))
        except se.OrganizerDrawingError:
            continue
        raise AssertionError(f"a drawing with {what} was read")
    # A pure turn is fine: that is what the upper plate's oval uses.
    c, s_ = math.cos(0.4), math.sin(0.4)
    turned = write_svg(PLATE + f'<ellipse transform="matrix({c},{s_},{-s_},{c},50,50)" rx="20" ry="8" cx="0" cy="0"/>')
    assert abs(se.read_die_organizer_drawing(path=turned)["inner"][0]["angle"] - 0.4) < 1e-9


check("reader: a path, group, line, scale, skew, mirror, unit suffix or missing plate is refused, never half read",
      reader_fails_loudly)


# ------------------------------------------------------------------ the G-code

def every_element_is_cut_once_where_the_drawing_puts_it():
    for variant in ("upper", "lower"):
        for kerf in (0.0, 0.2, 0.35):
            h = kerf / 2
            path, info = generate(variant, kerf)
            want = own_reading(variant)
            got = strokes_of(path)
            assert len(got) == len(want), f"{variant}: {len(got)} cut paths for {len(want)} drawn elements"
            assert sum(info["counts"].values()) == len(want)
            used = set()
            assert sum(1 for _k, plate, _b in want if plate) == 1
            for kind, plate, box in want:
                x0, y0, x1, y1 = box(h if plate else -h)   # the edge grows, an opening shrinks
                # Onto the sheet: 2 mm in from the corner, y turned up.
                ex0, ex1 = MARGIN + x0, MARGIN + x1
                ey0, ey1 = MARGIN + PLATE_H - y1, MARGIN + PLATE_H - y0
                match = [i for i, (_lay, pts) in enumerate(got) if i not in used
                         and max(abs(a - b) for a, b in zip(bbox(pts), (ex0, ey0, ex1, ey1))) < 2e-3]
                assert match, f"{variant} kerf {kerf}: no cut path for the {kind} at {(x0, y0, x1, y1)}"
                used.add(match[0])
                assert got[match[0]][0] == ("C02" if plate else "C01"), f"{kind} is on the wrong layer"
                pts = got[match[0]][1]
                assert pts[0] == pts[-1], f"the {kind} is not a closed cut"
            assert len(used) == len(got)


check("G-code: every drawn element is one closed cut where an independent reading of the SVG puts it, kerf applied",
      every_element_is_cut_once_where_the_drawing_puts_it)


def inner_first_then_the_edge():
    for variant in ("upper", "lower"):
        path, _info = generate(variant, 0.2)
        got = strokes_of(path)
        layers = [lay for lay, _p in got]
        assert layers == ["C01"] * (len(got) - 1) + ["C02"], "the openings come first and the plate's edge last"
        x0, y0, x1, y1 = bbox(got[-1][1])
        assert abs((x1 - x0) - (PLATE_W + 0.2)) < 2e-3 and abs((y1 - y0) - (PLATE_H + 0.2)) < 2e-3
        text = open(path).read()
        assert text.index("; Layer C01") < text.index("; Layer C02")


check("G-code: all the slots and holes are cut before the plate's own edge", inner_first_then_the_edge)


def bounding_box_sheet_and_laser_off():
    for variant in ("upper", "lower"):
        for kerf in (0.0, 0.2):
            path, info = generate(variant, kerf)
            pts = [p for _lay, s in strokes_of(path) for p in s]
            x0, y0, x1, y1 = bbox(pts)
            h = kerf / 2
            assert max(abs(x0 - (MARGIN - h)), abs(y0 - (MARGIN - h)),
                       abs(x1 - (MARGIN + PLATE_W + h)), abs(y1 - (MARGIN + PLATE_H + h))) < 2e-3, (x0, y0, x1, y1)
            assert (info["width"], info["height"]) == (PLATE_W + 2 * MARGIN, PLATE_H + 2 * MARGIN)
            assert x0 >= 0 and y0 >= 0 and x1 <= info["width"] and y1 <= info["height"], "a cut leaves the sheet"
            lines = open(path).read().splitlines()
            last_cut = max(i for i, ln in enumerate(lines) if re.match(r"G1 X", ln))
            assert "M5" in lines[last_cut:], "the laser must be switched off after the last cut"
            assert not [ln for ln in lines if re.search(r"F0(?![\d.])", ln)], "a zero feed was written"


check("G-code: the cut is 230 x 330 plus the kerf, 2 mm in from the corner, on the sheet, laser off at the end",
      bounding_box_sheet_and_laser_off)


def kerf_goes_inward_on_openings_and_outward_on_the_edge():
    widths = {}
    for kerf in (0.0, 0.3):
        path, _info = generate("upper", kerf)
        got = strokes_of(path)
        slots = sorted(round(min(bbox(p)[2] - bbox(p)[0], bbox(p)[3] - bbox(p)[1]), 3)
                       for lay, p in got if lay == "C01")
        widths[kerf] = (slots, bbox(got[-1][1]))
    narrow0, narrow3 = widths[0.0][0][0], widths[0.3][0][0]
    assert abs(narrow0 - 3.175) < 1e-3, f"with no kerf the die slots are the drawn 3.175 mm, got {narrow0}"
    assert abs(narrow3 - (3.175 - 0.3)) < 1e-3, f"the path of a 3.175 slot is a kerf narrower, got {narrow3}"
    e0, e3 = widths[0.0][1], widths[0.3][1]
    assert abs((e3[2] - e3[0]) - (e0[2] - e0[0]) - 0.3) < 2e-3, "the edge's path is a kerf wider"
    # The holes: 3.2 mm drawn (r 1.6).
    path, _info = generate("lower", 0.3)
    holes = [p for lay, p in strokes_of(path) if lay == "C01"]
    assert all(abs((bbox(p)[2] - bbox(p)[0]) - (3.2 - 0.3)) < 2e-3 for p in holes), [bbox(p) for p in holes]


check("kerf: slot and hole paths are a kerf smaller, the edge a kerf bigger, and at kerf 0 all are as drawn",
      kerf_goes_inward_on_openings_and_outward_on_the_edge)


def corners_are_true_arcs():
    kerf = 0.2
    path, _info = generate("lower", kerf)
    edge = strokes_of(path)[-1][1]
    r = 10.0 + kerf / 2
    # Corner centres of the plate on the sheet (the centres do not move with kerf).
    centres = [(MARGIN + cx, MARGIN + cy) for cx in (10.0, PLATE_W - 10.0) for cy in (10.0, PLATE_H - 10.0)]
    on_arc = 0
    for x, y in edge:
        in_x = MARGIN + 10.0 - 1e-6 <= x <= MARGIN + PLATE_W - 10.0 + 1e-6
        in_y = MARGIN + 10.0 - 1e-6 <= y <= MARGIN + PLATE_H - 10.0 + 1e-6
        if in_x or in_y:
            continue                                    # on a straight side
        d = min(math.hypot(x - cx, y - cy) for cx, cy in centres)
        assert abs(d - r) < 1.5e-3, f"corner point {(x, y)} is {d:.4f} from its centre, not {r}"
        on_arc += 1
    assert on_arc >= 4 * 20, f"only {on_arc} points round four corners: not finely linearized"
    # No chord sags more than the tolerance from the true arc.
    worst = 0.0
    for (ax, ay), (bx, by) in zip(edge, edge[1:]):
        mx, my = (ax + bx) / 2, (ay + by) / 2
        if (MARGIN + 10.0 < mx < MARGIN + PLATE_W - 10.0) or (MARGIN + 10.0 < my < MARGIN + PLATE_H - 10.0):
            continue
        worst = max(worst, r - min(math.hypot(mx - cx, my - cy) for cx, cy in centres))
    assert worst <= se.TOOLPATH_ARC_TOLERANCE_MM + 1e-3, f"a corner chord sags {worst:.4f} mm"
    # A hole is a circle about its centre.
    hole = strokes_of(path)[0][1]
    hx0, hy0, hx1, hy1 = bbox(hole)
    cx, cy = (hx0 + hx1) / 2, (hy0 + hy1) / 2
    assert all(abs(math.hypot(x - cx, y - cy) - (1.6 - kerf / 2)) < 1.5e-3 for x, y in hole)


check("corners and holes are true arcs, finely linearized (no chord further than the tolerance from the curve)",
      corners_are_true_arcs)


def the_oval_is_offset_along_its_normals():
    up = se.read_die_organizer_drawing("upper")
    ell = [s for s in up["inner"] if s["kind"] == "ellipse"][0]
    plain = se.organizer_shape_path(ell, 0.0)
    inset = se.organizer_shape_path(ell, -0.15)
    # Every point of the inset path is 0.15 mm from the drawn oval (not a scaled oval).
    def dist_to(pts, p):
        best = 1e9
        for a, b in zip(pts, pts[1:]):
            vx, vy = b[0] - a[0], b[1] - a[1]
            t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / (vx * vx + vy * vy)))
            best = min(best, math.hypot(p[0] - a[0] - t * vx, p[1] - a[1] - t * vy))
        return best
    worst = max(abs(dist_to(plain, p) - 0.15) for p in inset[::7])
    assert worst < 3e-3, f"the oval's kerf path is up to {worst:.4f} mm off a true parallel curve"
    scaled = [(ell["tx"] + (x - ell["tx"]) * (1 - 0.15 / ell["ry"]), ell["ty"] + (y - ell["ty"]) * (1 - 0.15 / ell["ry"]))
              for x, y in plain]
    assert max(abs(dist_to(plain, p) - 0.15) for p in scaled[::7]) > 0.05, \
        "fixture: a scaled oval should NOT be a parallel curve, or this check proves nothing"


check("the turned oval is moved along its normals (a true parallel curve), not scaled", the_oval_is_offset_along_its_normals)


def cut_settings_come_from_the_material():
    s = settings(0.2, "acrylic")
    s["gcode_settings"]["acrylic"].update(cut_speed=321, cut_power=77, cut_passes=2, air_assist_cut=False)
    path, _info = generate("lower", 0.2, "acrylic", s=s)
    text = open(path).read()
    assert "F321" in text and "S770" in text, "cut speed and power are not the material's"
    assert "M8" not in text, "air assist was switched on though the material has it off"
    one = settings(0.2, "acrylic")
    one["gcode_settings"]["acrylic"].update(cut_speed=321, cut_power=77, cut_passes=1, air_assist_cut=False)
    p1, _i = generate("lower", 0.2, "acrylic", s=one)
    n1 = len([ln for ln in open(p1).read().splitlines() if ln.startswith("G1 X")])
    n2 = len([ln for ln in text.splitlines() if ln.startswith("G1 X")])
    assert n2 > n1 * 1.8, f"two passes should about double the cutting moves ({n1} -> {n2})"
    s["gcode_settings"]["acrylic"]["cut_power"] = 0
    try:
        generate("lower", 0.2, "acrylic", s=s)
    except ValueError:
        return
    raise AssertionError("a material with no cut power was cut at S0")


check("cut speed, power, passes and air come from the chosen material; one with no cut settings is refused",
      cut_settings_come_from_the_material)


def refusals():
    need_w, need_h = se.die_organizer_min_sheet("upper")
    assert (need_w, need_h) == (PLATE_W + 2 * MARGIN, PLATE_H + 2 * MARGIN)
    for sheet in ((200.0, 400.0), (300.0, 300.0), (233.9, 334.0)):
        try:
            generate("upper", 0.2, sheet=sheet)
        except ValueError as e:
            assert "234" in str(e) and "334" in str(e), str(e)
            continue
        raise AssertionError(f"a {sheet} sheet took a 230 x 330 plate")
    generate("upper", 0.2, sheet=(need_w, need_h))          # exactly big enough is fine
    try:
        generate("upper", 3.5)                               # wider than the 3.175 mm die slots
    except se.OrganizerKerfError as e:
        assert isinstance(e, se.OrganizerDrawingError) and isinstance(e, ValueError)
        # The opening it names is one the kerf really bridges (the first in
        # drawing order: the 3.2 mm pin holes come before the 3.175 mm slots).
        assert e.kerf_mm == 3.5 and e.what in ("slot", "hole") and float(e.opening.split()[-1]) < 3.5,             (e.kerf_mm, e.what, e.opening)
        assert "kerf" in str(e) and "small" in str(e), str(e)
    else:
        raise AssertionError("a kerf wider than the die slots was accepted")
    # A negative kerf is refused, at the toolpaths and at the generator, and
    # says so with no opening named (there is nothing too narrow for it).
    for fn in (lambda: se.die_organizer_toolpaths("upper", -0.1), lambda: generate("upper", -0.1),
               lambda: generate("lower", -2.0)):
        try:
            fn()
        except se.OrganizerKerfError as e:
            assert e.opening is None and e.kerf_mm < 0 and "negative" in str(e), (e.opening, e.kerf_mm, str(e))
        else:
            raise AssertionError("a negative kerf was accepted")
    assert se.die_organizer_toolpaths("upper", 0.0)["counts"]["rect"] == 112, "kerf 0 is fine"
    # The slots themselves: narrower than the kerf is refused in so many words.
    slot = {"kind": "rect", "angle": 0.0, "tx": 0.0, "ty": 0.0, "x": 0.0, "y": 0.0, "w": 54.0, "h": 3.175, "r": 0.0}
    try:
        se.organizer_shape_path(slot, -3.2 / 2)
    except se.OrganizerKerfError as e:
        assert e.what == "slot" and e.opening == "54 x 3.175" and e.kerf_mm == 3.2, (e.what, e.opening, e.kerf_mm)
    else:
        raise AssertionError("a slot narrower than the kerf was given a path")
    assert len(se.organizer_shape_path(slot, -0.1)) == 5, "a square slot moved inward stays four square corners"


check("a sheet smaller than the plate is refused, and so is a kerf wider than the slots", refusals)


def the_svg_is_still_a_byte_copy():
    for variant in ("upper", "lower"):
        out = os.path.join(OUT, f"copy_{variant}.svg")
        se.generate_die_organizer_svg(variant, out, None)
        assert open(out, "rb").read() == open(asset(variant), "rb").read()


check("Generate SVG is still a byte-for-byte copy of the bundled drawing", the_svg_is_still_a_byte_copy)


# ------------------------------------------------------------------------ GUI

def gui():
    import tkinter as tk
    from tkinter import ttk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import main as main_mod
    import tooling_tab
    import ui_dialogs
    app = main_mod.PadSVGGeneratorApp(root)
    saved, boxes = [], []

    def fake_save(**kw):
        path = os.path.join(OUT, kw.get("initialfile"))
        saved.append(path)
        return path
    tooling_tab.filedialog.asksaveasfilename = fake_save
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))

    def section_has_a_material_and_a_sheet():
        frame = app._tooling_sections["die_organizer"]
        combos = [w for row in frame.winfo_children() for w in row.winfo_children() if isinstance(w, ttk.Combobox)]
        assert len(combos) == 1 and list(combos[0].cget("values")) == ["Basswood", "Acrylic"],\
            [list(c.cget("values")) for c in combos]
        assert app.organizer_material_var.get() == "Basswood" and app._organizer_material_key() == "basswood"
        texts = []

        def walk(w):
            for c in w.winfo_children():
                if isinstance(c, (tk.Label, tk.Button)):
                    texts.append(str(c.cget("text")))
                walk(c)
        walk(frame)
        assert "Generate G-code" in texts and "Generate SVG" in texts
        info = [t for t in texts if t.startswith("Plate: 230")]
        assert info and "Generate G-code" in info[0] and "LightBurn" in info[0], info
        assert not [t for t in texts if "Open the SVG in LightBurn or your laser software to cut" in t], \
            "the section still says LightBurn is the way to cut it"
        app.organizer_material_var.set("Acrylic")
        app._update_tooling_settings()
        assert app.settings["tooling_settings"]["organizer_material"] == "acrylic", "stored as the key, not the label"
        app.organizer_material_var.set("Basswood")
    check("GUI: the section offers a material (basswood first) and a sheet size, and no longer points only at LightBurn",
          section_has_a_material_and_a_sheet)

    def generate_through_the_button():
        app.organizer_units_var.set("mm")
        app.organizer_width_var.set("300")
        app.organizer_height_var.set("400")
        app.organizer_filename_var.set("org")
        for variant in ("upper", "lower"):
            app.organizer_variant_var.set(variant)
            del boxes[:]
            app._on_generate_organizer_gcode()
            root.update()
            assert not [b for b in boxes if b[0] == "showerror"], boxes
            path = saved[-1]
            assert os.path.basename(path) == f"org_{variant}.gcode" and os.path.getsize(path) > 0
            want, _info = generate(variant, app.settings["gcode_settings"]["basswood"]["kerf_width"],
                                   s=app.settings)
            assert open(path).read() == open(want).read(), "the button's file is not the engine's output"
            done = [b for b in boxes if b[0] == "showinfo"]
            assert len(done) == 1 and done[0][2].endswith(ui_dialogs.gcode_file_note()), done
            assert path in done[0][2]
        del boxes[:]
        app._on_generate_organizer_svg()
        root.update()
        done = [b for b in boxes if b[0] == "showinfo"]
        assert len(done) == 1 and ui_dialogs.gcode_file_note() not in done[0][2], "the SVG box must not carry the G-code line"
    check("GUI: Generate G-code writes the engine's file for each plate; the Done box ends with the laser-file line",
          generate_through_the_button)

    def a_small_sheet_is_refused_before_the_save_dialog():
        app.organizer_variant_var.set("upper")
        # Too small in both directions, in one direction only (each way),
        # by a hair, not a number, zero.
        for w, h, units in (("200", "300", "mm"), ("9", "12", "in"), ("200", "400", "mm"), ("400", "200", "mm"),
                            ("233.9", "334", "mm"), ("234", "333.9", "mm"), ("9.2", "14", "in"),
                            ("wide", "400", "mm"), ("0", "400", "mm")):
            app.organizer_units_var.set(units)
            app.organizer_width_var.set(w)
            app.organizer_height_var.set(h)
            del boxes[:]
            before = len(saved)
            app._on_generate_organizer_gcode()
            root.update()
            assert len(saved) == before, f"{w} x {h} {units}: reached the save dialog"
            errs = [b for b in boxes if b[0] == "showerror"]
            assert len(errs) == 1, boxes
            if w not in ("wide", "0"):
                assert errs[0][1] == "Sheet too small" and "234" in errs[0][2] and "334" in errs[0][2], \
                    f"{w} x {h} {units}: {errs}"
        app.organizer_units_var.set("in")
        app.organizer_width_var.set("12")
        app.organizer_height_var.set("18")
        del boxes[:]
        app._on_generate_organizer_gcode()
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], "12 x 18 in is big enough"
    check("GUI: a sheet smaller than 234 x 334 mm is refused in plain words before anything is saved",
          a_small_sheet_is_refused_before_the_save_dialog)

    def a_refusal_leaves_no_file():
        app.organizer_units_var.set("mm")
        app.organizer_width_var.set("300")
        app.organizer_height_var.set("400")
        app.organizer_filename_var.set("refused")
        keep = app.settings["gcode_settings"]["basswood"]["kerf_width"]
        app.settings["gcode_settings"]["basswood"]["kerf_width"] = 3.5
        try:
            del boxes[:]
            app._on_generate_organizer_gcode()
            root.update()
        finally:
            app.settings["gcode_settings"]["basswood"]["kerf_width"] = keep
        errs = [b for b in boxes if b[0] == "showerror"]
        assert errs and not [b for b in boxes if b[0] == "showinfo"], boxes
        assert not os.path.exists(os.path.join(OUT, "refused_upper.gcode")), "a refused job left a file behind"
        # The message is the tab's, in the user's words (translated), not the
        # engine's English: it names the material, the kerf, the opening and
        # where to change it.
        where = app._material_settings_place("basswood")
        assert errs[0][1] == "Invalid Input" and errs[0][2].startswith(
            "The Basswood kerf (3.5 mm) is too wide for an opening of ") and errs[0][2].endswith(
            f"in the organizer plate. Check it in {where}."), errs[0]
        named = re.search(r"opening of ([\d.]+(?: x [\d.]+)?) mm", errs[0][2])
        assert named and float(named.group(1).split()[-1]) < 3.5 and "narrow" not in errs[0][2], errs[0][2]
        # A negative kerf, likewise.
        app.settings["gcode_settings"]["basswood"]["kerf_width"] = -0.2
        try:
            del boxes[:]
            app._on_generate_organizer_gcode()
            root.update()
        finally:
            app.settings["gcode_settings"]["basswood"]["kerf_width"] = keep
        errs = [b for b in boxes if b[0] == "showerror"]
        assert errs and errs[0][2] == f"The Basswood kerf is negative (-0.2 mm). Check it in {where}.", errs
        assert not os.path.exists(os.path.join(OUT, "refused_upper.gcode"))
        # No cut settings gets its sentence; any other ValueError (a kerf that
        # is not a number) does not: it reaches the user as what it is.
        for value, want, forbidden in (("cut_speed", 0, None), ("kerf_width", "wide", "has no cut speed and power set")):
            saved_value = app.settings["gcode_settings"]["basswood"][value]
            app.settings["gcode_settings"]["basswood"][value] = want
            try:
                del boxes[:]
                app._on_generate_organizer_gcode()
                root.update()
            finally:
                app.settings["gcode_settings"]["basswood"][value] = saved_value
            errs = [b for b in boxes if b[0] == "showerror"]
            assert len(errs) == 1, (value, boxes)
            if forbidden is None:
                assert "Basswood has no cut speed and power set" in errs[0][2], errs[0]
            else:
                assert forbidden not in errs[0][2] and "wide" in errs[0][2],                     f"a kerf of 'wide' was reported as missing cut settings or not named: {errs[0]}"
            assert not os.path.exists(os.path.join(OUT, "refused_upper.gcode"))
    check("GUI: a job the engine refuses (kerf wider than the slots, or negative) says so in the tab's own words "
          "and leaves no file", a_refusal_leaves_no_file)

    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
