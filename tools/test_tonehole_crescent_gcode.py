"""Tonehole Crescent G-code: cut by this app, to size, from a known material.

The crescent was SVG and STL only. Since 2026-11-09 it has a G-code path
(and Frame & Cut), cut from a material the section now asks for, cork by
default. The SVG cannot allow for kerf because it does not know the
material; the G-code does, so it must. This suite pins:

  - the kerf path is the outline's exact parallel curve: over a grid of
    tonehole sizes, percentages, inner-edge curves and kerfs, every point
    of the path is half a kerf from the design outline (measured here
    against the two design arcs, not with the app's own offset code), its
    area is the design area plus perimeter x kerf/2 plus pi (kerf/2)^2,
    and every point of the design outline is half a kerf inside it. So a
    beam of that kerf leaves the designed part.
  - with kerf 0 the path IS the outline (cork ships with kerf 0)
  - the horns never cross or close; a crescent whose horn gap is narrower
    than the kerf is refused; a thin crescent is never refused for being thin
  - the path stays within half a kerf of the design (the horns are wrapped,
    not run on to a far-off mitre point)
  - the G-code: numbers before any cut, one closed cut per crescent, the
    material's cut and engraving settings, the SVG's layout and numbering
    rule, and the same physical side up as the SVG after the Y flip
  - a material with no engraving settings (cork) gets no numbers and never
    an F0 line; the cuts are all there
  - through the real section: the material choice (cork first), the Done
    box, the refusals

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
_PROFILE = tempfile.mkdtemp(prefix="ssc-crescent-gcode-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import config  # noqa: E402
import gcode_engine as ge  # noqa: E402
import svg_engine as se  # noqa: E402

results = []
OUT = tempfile.mkdtemp(prefix="ssc-crescent-out-")
TOL = se.TOOLPATH_ARC_TOLERANCE_MM


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


# ---- geometry written here, independent of the app's offset code -------------

def polygon_area(pts):
    return abs(sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(pts, pts[1:]))) / 2.0


def horn_angles(g):
    """Right horn seen from the tonehole centre and from the inner circle's."""
    return math.atan2(g["tip_y"], g["tip_x"]), math.atan2(g["tip_y"] - g["offset"], g["tip_x"])


def on_arc(angle, a_right):
    """Is `angle` on the arc that runs from the left horn over the
    mouthpiece side (-y) to the right horn? Both design arcs do."""
    # The arc is every direction NOT in the open wedge between the horns on the far side.
    d = (angle - math.pi / 2 + math.pi) % (2 * math.pi) - math.pi      # measured from straight down (+y)
    half_gap = math.pi / 2 - a_right                                    # half the wedge left open
    return abs(d) >= half_gap - 1e-12


def dist_to_design(g, p):
    """Distance from p to the design outline: the outer arc, the inner arc,
    or a horn tip, whichever is nearest."""
    R, Rc, off, tx, ty = g["radius"], g["inner_radius"], g["offset"], g["tip_x"], g["tip_y"]
    a_r, b_r = horn_angles(g)
    best = min(math.hypot(p[0] - tx, p[1] - ty), math.hypot(p[0] + tx, p[1] - ty))
    if on_arc(math.atan2(p[1], p[0]), a_r):
        best = min(best, abs(math.hypot(p[0], p[1]) - R))
    if on_arc(math.atan2(p[1] - off, p[0]), b_r):
        best = min(best, abs(math.hypot(p[0], p[1] - off) - Rc))
    return best


def dist_to_path(path, p):
    best = 1e9
    for a, b in zip(path, path[1:]):
        vx, vy = b[0] - a[0], b[1] - a[1]
        den = vx * vx + vy * vy
        t = 0.0 if den == 0 else max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / den))
        best = min(best, math.hypot(p[0] - a[0] - t * vx, p[1] - a[1] - t * vy))
    return best


def inside(path, p):
    hit = False
    for a, b in zip(path, path[1:]):
        if (a[1] > p[1]) != (b[1] > p[1]) and p[0] < a[0] + (p[1] - a[1]) * (b[0] - a[0]) / (b[1] - a[1]):
            hit = not hit
    return hit


def design_numbers(g, percent):
    """(area, perimeter) of the designed crescent, from the design itself."""
    a_r, b_r = horn_angles(g)
    area = percent / 100.0 * math.pi * g["radius"] ** 2
    perimeter = g["radius"] * (math.pi + 2 * a_r) + g["inner_radius"] * (math.pi + 2 * b_r)
    return area, perimeter


GRID = [(d, pct, curve) for d in (8.0, 20.0, 43.0, 60.0, 120.0)
        for pct in (0.5, 2.0, 5.0, 10.0, 20.0, 40.0, 75.0) for curve in (0, 40, 75, 90, 100)]


def kerf_zero_is_the_outline():
    for d, pct, curve in GRID:
        g = se.tonehole_crescent_geometry(d, pct, curve)
        path = se.tonehole_crescent_toolpath(g, 0.0)
        assert path[0] == path[-1], "the path is not closed"
        worst = max(dist_to_design(g, p) for p in path)
        assert worst < 1e-9, f"{d} mm {pct}% curve {curve}: a kerf-0 point is {worst} off the outline"
        area, perimeter = design_numbers(g, pct)
        assert abs(polygon_area(path) - area) <= perimeter * TOL * 1.1 + 1e-9, \
            f"{d} mm {pct}%: kerf-0 path encloses {polygon_area(path):.4f}, the design is {area:.4f} mm2"


check("kerf 0 is a no-op: the path is the design outline, enclosing the design area (175 crescents)",
      kerf_zero_is_the_outline)


def the_path_is_the_exact_parallel_curve():
    cases = refused = 0
    for d, pct, curve in GRID:
        g = se.tonehole_crescent_geometry(d, pct, curve)
        for kerf in (0.1, 0.2, 0.5):
            h = kerf / 2
            try:
                path = se.tonehole_crescent_toolpath(g, kerf)
            except se.CrescentKerfError as e:
                assert e.reason == "gap" and 2 * g["tip_x"] <= kerf, \
                    f"{d} mm {pct}% curve {curve} kerf {kerf}: refused for {e.reason} with a {2 * g['tip_x']:.3f} mm gap"
                refused += 1
                continue
            cases += 1
            what = f"{d} mm {pct}% curve {curve} kerf {kerf}"
            # 1. Every point of the path is half a kerf from the design outline.
            worst = max(abs(dist_to_design(g, p) - h) for p in path)
            assert worst < 1e-7, f"{what}: a path point is {worst:.2e} off half a kerf from the outline"
            # 2. So it never strays: no mitre running off past the horns.
            reach = max(math.hypot(p[0], p[1]) for p in path)
            assert reach <= g["radius"] + h + 1e-9, f"{what}: the path reaches {reach - g['radius']:.3f} mm past the tonehole"
            # 3. Its area is the design's plus the band a parallel curve adds.
            area, perimeter = design_numbers(g, pct)
            want = area + perimeter * h + math.pi * h * h
            assert abs(polygon_area(path) - want) <= (perimeter + 2 * math.pi * h) * TOL * 1.1 + 1e-6, \
                f"{what}: path area {polygon_area(path):.4f}, an exact offset is {want:.4f}"
            # 4. Shrunk by half a kerf it is the design: each design point is
            #    inside the path and half a kerf from it.
            outer, inner = se.tonehole_crescent_outline(g, 24)
            for q in outer + inner:                     # the two horn tips included
                dq = dist_to_path(path, q)
                assert abs(dq - h) <= TOL * 1.2 + 1e-9, f"{what}: design point {q} is {dq:.4f} from the path, not {h}"
            for q in outer[2:-2:5]:
                assert inside(path, (q[0] * (1 - 1e-6), q[1] * (1 - 1e-6))), f"{what}: the design is not inside the path"
            # 5. The horns neither cross nor close.
            assert g["tip_x"] > h, what
            left = min(p[0] for p in path if dist_to_design(g, p) > -1 and math.hypot(p[0] + g["tip_x"], p[1] - g["tip_y"]) <= h + 1e-9)
            right = max(p[0] for p in path if math.hypot(p[0] - g["tip_x"], p[1] - g["tip_y"]) <= h + 1e-9)
            assert left < 0 < right, f"{what}: the horn wraps crossed"
    assert cases >= 480 and refused <= 45, (cases, refused)


check("the kerf path is the outline's exact parallel curve: half a kerf away everywhere, right area, design inside",
      the_path_is_the_exact_parallel_curve)


def a_mitre_would_have_run_off():
    # Why the horns are wrapped: on a thin crescent the two offset circles
    # cross far past the horn. The wrapped path never leaves the design by
    # more than half a kerf.
    g = se.tonehole_crescent_geometry(43.0, 2.0)
    h = 0.1
    R, Rc, off = g["radius"], g["inner_radius"], g["offset"]
    y_m = (off * off + (R + h) ** 2 - (Rc - h) ** 2) / (2 * off)
    x_m = math.sqrt((R + h) ** 2 - y_m ** 2)
    assert math.hypot(x_m - g["tip_x"], y_m - g["tip_y"]) > 6.0, "fixture: the mitre point should be far off here"
    path = se.tonehole_crescent_toolpath(g, 0.2)
    assert max(dist_to_design(g, p) for p in path) <= h + 1e-7
    assert max(p[1] for p in path) <= g["tip_y"] + h + 1e-9, "the path runs on below the horns"


check("thin crescents: the path wraps each horn and stays within half a kerf (a mitre would run 6 mm past)",
      a_mitre_would_have_run_off)


def refusals_are_only_for_a_gap_narrower_than_the_kerf():
    for d in (8.0, 43.0, 120.0):
        for pct in (0.1, 0.5, 1.0):
            g = se.tonehole_crescent_geometry(d, pct)
            se.tonehole_crescent_toolpath(g, 0.5)          # thin is fine: the path only moves away from the part
    # The horns are never closer than about two thirds of the tonehole's
    # diameter (75 %, roundest edge), so with the largest kerf allowed only
    # a tonehole under 1.5 mm can be refused: a guard, not a limit anyone
    # will meet on a saxophone.
    tight = se.tonehole_crescent_geometry(10.0, 75.0, 100)
    assert 0.60 < 2 * tight["tip_x"] / 10.0 < 0.70, 2 * tight["tip_x"] / 10.0
    closed = se.tonehole_crescent_geometry(1.2, 75.0, 100)
    assert 2 * closed["tip_x"] <= 1.0, "fixture: need a crescent whose horns are under 1 mm apart"
    try:
        se.tonehole_crescent_toolpath(closed, 1.0)
    except se.CrescentKerfError as e:
        assert e.reason == "gap" and abs(e.size_mm - 2 * closed["tip_x"]) < 1e-9 and e.kerf_mm == 1.0
        assert "narrower than the kerf" in str(e)
    else:
        raise AssertionError("a crescent whose horns are closer together than the kerf was given a path")
    se.tonehole_crescent_toolpath(closed, 0.0)             # with no kerf there is nothing to bridge
    # The threshold is the gap itself, on both sides: a kerf a hair under
    # the gap cuts, a hair over is refused. (Comparing the kerf with half
    # the gap would refuse the first and nobody noticed.)
    small = se.tonehole_crescent_geometry(1.4, 75.0, 100)
    gap_mm = 2 * small["tip_x"]
    assert 0.5 < gap_mm < se.CRESCENT_MAX_KERF_MM, f"fixture: a gap inside the kerf range, got {gap_mm}"
    assert se.tonehole_crescent_toolpath(small, gap_mm - 0.001), "a kerf just narrower than the gap must cut"
    assert se.tonehole_crescent_toolpath(small, gap_mm * 0.6), "a kerf well inside the gap must cut"
    for too_wide in (gap_mm + 0.001, gap_mm):
        try:
            se.tonehole_crescent_toolpath(small, too_wide)
        except se.CrescentKerfError as e:
            assert e.reason == "gap" and abs(e.size_mm - gap_mm) < 1e-9, (e.reason, e.size_mm)
        else:
            raise AssertionError(f"a kerf of {too_wide} mm bridged a {gap_mm:.3f} mm gap and was accepted")
    g = se.tonehole_crescent_geometry(43.0, 20.0)
    for bad in (-0.1, se.CRESCENT_MAX_KERF_MM + 0.01, 5.0):
        try:
            se.tonehole_crescent_toolpath(g, bad)
        except se.CrescentKerfError as e:
            assert e.reason == "kerf"
        else:
            raise AssertionError(f"a kerf of {bad} mm was accepted")


check("refused only when the horn gap is narrower than the kerf (or the kerf is out of range); never for being thin",
      refusals_are_only_for_a_gap_narrower_than_the_kerf)


# ---------------------------------------------------------------------- G-code

def settings(material="acrylic", kerf=None, **changes):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    if kerf is not None:
        s["gcode_settings"][material]["kerf_width"] = kerf
    s["gcode_settings"][material].update(changes)
    return s


def layers(path):
    """[(layer name, [stroke, ...])] with each stroke a list of (x, y)."""
    out, pending = [], None
    x = y = None
    for ln in open(path).read().splitlines():
        if ln.startswith("; Layer "):
            out.append((ln.split()[-1], []))
            continue
        if ln.startswith("; return to origin"):
            break
        m = re.match(r"G([01])\b", ln)
        if ln.startswith(";") or not m or not out:
            if m and m.group(1) == "0":
                mx, my = re.search(r"X(-?\d+\.?\d*)", ln), re.search(r"Y(-?\d+\.?\d*)", ln)
                if mx and my:
                    x, y = float(mx.group(1)), float(my.group(1))
                    pending = [(x, y)]
            continue
        mx, my = re.search(r"X(-?\d+\.?\d*)", ln), re.search(r"Y(-?\d+\.?\d*)", ln)
        if mx is None and my is None:
            continue
        x = float(mx.group(1)) if mx else x
        y = float(my.group(1)) if my else y
        if m.group(1) == "0":
            pending = [(x, y)]
        else:
            if pending is not None:
                out[-1][1].append(pending)
                pending = None
            out[-1][1][-1].append((x, y))
    return out


LIST = [{"tonehole_d": 43.0, "percent": 20.0, "curve": 75, "qty": 2},
        {"tonehole_d": 41.5, "percent": 10.0, "curve": 60, "qty": 1},
        {"tonehole_d": 20.0, "percent": 35.0, "curve": 90, "qty": 3}]


def write(entries, material="acrylic", s=None, name="c"):
    path = os.path.join(OUT, f"{name}_{material}_{len(os.listdir(OUT))}.gcode")
    info = ge.generate_tonehole_crescent_gcode(entries, material, path, s or settings(material))
    return path, info


def numbers_first_then_one_closed_cut_each():
    s = settings("acrylic", kerf=0.2, engraving_mode="line", engraving_speed=1500, engraving_power=12)
    path, info = write(LIST, s=s)
    got = layers(path)
    assert [name for name, _s in got] == ["C03", "C02"], "numbers (C03) then cuts (C02), nothing else"
    parts = info["parts"]
    assert len(parts) == 6 and info["numbers_skipped"] is False and info["kerf"] == 0.2
    cuts = got[1][1]
    assert len(cuts) == len(parts), f"{len(cuts)} cut strokes for {len(parts)} crescents"
    H = info["height"]
    for p, stroke in zip(parts, cuts):
        assert stroke[0] == stroke[-1], "a crescent's cut is not one closed path"
        want = [(round(p["cx"] + x, 3), round(H - (p["cy"] + y), 3))
                for x, y in se.tonehole_crescent_toolpath(p["geom"], 0.2)]
        assert len(stroke) == len(want) and max(
            max(abs(a[0] - b[0]), abs(a[1] - b[1])) for a, b in zip(stroke, want)) < 1.5e-3, \
            "the cut is not the kerf path of its own crescent at its own place"
    text = open(path).read()
    assert "F1500" in text and "S120" in text, "the numbers are not at the material's engraving speed and power"
    mat = s["gcode_settings"]["acrylic"]
    assert f"F{mat['cut_speed']}" in text and f"S{int(mat['cut_power'] * 10)}" in text
    assert text.index("; Layer C03") < text.index("; Layer C02")
    # Each number is under its own part, on the scrap, and is that entry's number.
    numbered = [p for p in parts]
    strokes = got[0][1]
    for p in numbered:
        mine = [pt for st in strokes for pt in st if abs(pt[0] - p["label_x"]) < 2.0
                and abs(pt[1] - (H - (p["label_y"] - 0.35 * se.CRESCENT_LABEL_FONT_MM))) < 2.2]
        assert mine, f"no number engraved under the part at {(p['cx'], p['cy'])}"
        lowest_cut = min(y for x, y in cuts[parts.index(p)])
        assert max(y for _x, y in mine) < lowest_cut, "a number runs into its crescent"
    assert sorted({p["number"] for p in parts}) == [1, 2, 3]


check("G-code: list numbers are engraved first, under their parts; then one closed kerf path per crescent",
      numbers_first_then_one_closed_cut_each)


def each_numbers_box_matches_the_svg_and_sits_in_its_own_strip():
    """The SVG anchors its <text> at the baseline; the G-code's stroke font
    is centre-anchored, so the G-code moves the centre up by 0.35 x font.
    With that sign the wrong way round the numbers sit 2.1 mm too low and
    run out of their strip. Pinned here against the SVG file's own text
    and against the cut of the part each number belongs to, neither read
    through the engine's placement formula."""
    f, gap = se.CRESCENT_LABEL_FONT_MM, se.CRESCENT_LABEL_GAP_MM
    kerf = 0.2
    s = settings("acrylic", kerf=kerf, engraving_mode="line", engraving_speed=1500, engraving_power=12)
    path, info = write(LIST, s=s, name="box")
    svg_path = os.path.join(OUT, "box.svg")
    se.generate_tonehole_crescent_svg(LIST, svg_path, s)
    H = info["height"]
    with open(svg_path, encoding="utf-8") as fh:
        svg = fh.read()
    texts = [(float(x), float(y), int(n)) for x, y, n in
             re.findall(r'<text[^>]*x="([\d.]+)mm"[^>]*y="([\d.]+)mm"[^>]*>(\d+)</text>', svg)]
    parts = info["parts"]
    assert len(texts) == len(parts) == 6, f"fixture: {len(texts)} texts in the SVG for {len(parts)} parts"
    got = layers(path)
    assert [name for name, _s in got] == ["C03", "C02"]
    number_strokes, cuts = got[0][1], got[1][1]
    # Each stroke of a number belongs to the SVG text nearest its centre
    # (the parts are many millimetres apart, the glyphs are 3 mm).
    boxes = {i: [] for i in range(len(texts))}
    for stroke in number_strokes:
        xs = [p[0] for p in stroke]
        ys = [p[1] for p in stroke]
        cx, cy_svg = (min(xs) + max(xs)) / 2, H - (min(ys) + max(ys)) / 2
        i = min(range(len(texts)), key=lambda k: (texts[k][0] - cx) ** 2 + (texts[k][1] - cy_svg) ** 2)
        boxes[i].append((min(xs), min(ys), max(xs), max(ys)))
    for i, (tx, ty, number) in enumerate(texts):
        p = parts[i]
        assert number == p["number"], (number, p["number"])
        assert boxes[i], f"no strokes near the SVG's number {number} at ({tx}, {ty})"
        x0 = min(b[0] for b in boxes[i])
        y0 = min(b[1] for b in boxes[i])
        x1 = max(b[2] for b in boxes[i])
        y1 = max(b[3] for b in boxes[i])
        centre_svg_y = H - (y0 + y1) / 2
        # Against the SVG: the glyph box's centre is 0.35 x font above the text baseline.
        assert abs(centre_svg_y - (ty - 0.35 * f)) < 0.02, (
            f"number {number}: its G-code box is centred {centre_svg_y:.3f} mm down the sheet, "
            f"the SVG's baseline is at {ty:.3f} (centre should be {ty - 0.35 * f:.3f})")
        assert abs((x0 + x1) / 2 - tx) < 0.5 * f, f"number {number}: not centred on the SVG's x {tx}"
        assert 0.8 * f < (y1 - y0) < 1.0 * f, f"number {number}: glyph box {y1 - y0:.2f} mm tall for a {f} mm font"
        # Against its own part: entirely below the part's lowest cut point
        # (G-code y runs up), and inside the strip the layout reserves
        # under the part (gap + font), kerf allowed for.
        cut_bottom = min(y for _x, y in cuts[i])
        assert y1 < cut_bottom - 0.5, f"number {number}: its top ({y1:.2f}) runs into its part (cut bottom {cut_bottom:.2f})"
        assert y0 > cut_bottom - (gap + f) - kerf / 2 - 0.05, (
            f"number {number}: its bottom ({y0:.2f}) is below the strip under its part "
            f"({cut_bottom - (gap + f):.2f})")


check("G-code: each number's box matches the SVG's text (centre 0.35 x font above the baseline) and sits in the "
      "strip under its own part", each_numbers_box_matches_the_svg_and_sits_in_its_own_strip)


def same_layout_and_numbering_rule_as_the_svg():
    for entries, numbered in ((LIST, True), ([dict(LIST[0], qty=3)], False), ([dict(LIST[1])], False)):
        path, info = write(entries, s=settings("acrylic", kerf=0.0))
        svg_parts, w, h = se.layout_tonehole_crescents(entries)
        assert (info["width"], info["height"]) == (w, h)
        assert [(p["cx"], p["cy"], p["number"]) for p in info["parts"]] == \
            [(p["cx"], p["cy"], p["number"]) for p in svg_parts]
        names = [name for name, _s in layers(path)]
        assert names == (["C03", "C02"] if numbered else ["C02"]), \
            f"{len(entries)} entries: layers {names} (numbers only when the list has more than one entry)"


check("G-code uses the SVG's layout, and numbers the parts only when the list has more than one entry",
      same_layout_and_numbering_rule_as_the_svg)


def svg_and_gcode_put_the_mouthpiece_side_the_same_way_up():
    entries = [{"tonehole_d": 43.0, "percent": 20.0, "curve": 75, "qty": 1},
               {"tonehole_d": 30.0, "percent": 55.0, "curve": 95, "qty": 1}]
    s = settings("acrylic", kerf=0.0)
    gpath, info = write(entries, s=s)
    spath = os.path.join(OUT, "parity.svg")
    se.generate_tonehole_crescent_svg(entries, spath, s)
    svg = open(spath, encoding="utf-8").read()
    paths = re.findall(r'd="M ([-\d.]+),([-\d.]+) A ([-\d.]+),[-\d.]+ 0 (\d) 1 ([-\d.]+),[-\d.]+ '
                       r'A ([-\d.]+),[-\d.]+ 0 (\d) 0', svg)
    assert len(paths) == 2, "fixture: two crescent paths expected in the SVG"
    cuts = layers(gpath)[-1][1]
    H = info["height"]
    for (lx, ty, R, big_outer, rx, Rc, big_inner), stroke in zip(paths, cuts):
        lx, ty, R, rx, Rc = float(lx), float(ty), float(R), float(rx), float(Rc)
        half = (rx - lx) / 2
        cx = (lx + rx) / 2
        # The SVG's own arcs: centre of the tonehole circle and of the inner circle.
        oy = ty - (1 if big_outer == "1" else -1) * math.sqrt(max(0.0, R * R - half * half))
        iy = ty - (1 if big_inner == "1" else -1) * math.sqrt(max(0.0, Rc * Rc - half * half))
        # Turn the G-code back down (y = H - Y) and it must lie on those two arcs.
        for X, Y in stroke:
            x, y = X, H - Y
            d = min(abs(math.hypot(x - cx, y - oy) - R), abs(math.hypot(x - cx, y - iy) - Rc))
            assert d < 2e-3, f"G-code point {(X, Y)} is {d:.4f} mm off the SVG's outline"
        # Mouthpiece side: the top of the tonehole circle. Up in the SVG (smallest y),
        # so up on the bed (largest Y).
        # (The topmost point of the arc need not be a vertex of the
        # linearized path, so this one is good to the arc tolerance.)
        assert abs(max(Y for _X, Y in stroke) - (H - (oy - R))) < TOL + 1e-3, "the crescent is upside down on the bed"
        assert max(Y for _X, Y in stroke) > H - ty + R / 4, "the bulge should be well above the horns"
        assert abs(min(Y for _X, Y in stroke) - (H - ty)) < 2e-3, "the horns should be the lowest points"
        xs = [X for X, _Y in stroke]
        assert abs((min(xs) + max(xs)) / 2 - cx) < 2e-3, "mirrored or shifted left to right"


check("parity: turned back over, the G-code lies on the SVG's two arcs, mouthpiece side up in both",
      svg_and_gcode_put_the_mouthpiece_side_the_same_way_up)


def kerf_comes_from_the_material_and_moves_the_cut():
    widths = {}
    for kerf in (0.0, 0.15, 0.4):
        path, info = write([dict(LIST[0], qty=1)], s=settings("acrylic", kerf=kerf))
        stroke = layers(path)[-1][1][0]
        xs = [x for x, _y in stroke]
        widths[kerf] = max(xs) - min(xs)
        assert info["kerf"] == kerf and f"kerf {kerf:g} mm compensated" in open(path).read()
    assert abs(widths[0.0] - 43.0) < 2e-3, f"at kerf 0 the cut is the tonehole's 43 mm across, got {widths[0.0]}"
    assert abs(widths[0.15] - 43.15) < 2e-3 and abs(widths[0.4] - 43.4) < 2e-3, widths


check("the kerf is the material's: a 43 mm crescent's path is 43 mm across at kerf 0 and 43.15 at 0.15",
      kerf_comes_from_the_material_and_moves_the_cut)


def cork_as_shipped_cuts_everything_and_engraves_nothing():
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    path, info = write(LIST, "cork", s=s)
    assert info["numbers_skipped"] is True and info["kerf"] == 0.0
    got = layers(path)
    assert [name for name, _s in got] == ["C02"], "cork has no engraving settings: cuts only"
    assert len(got[0][1]) == 6, "all six crescents are still cut"
    text = open(path).read()
    assert not re.search(r"F0(?![\d.])", text), "a zero feed was written"
    assert "F1000" in text and "S300" in text, "cork cuts at its tested 1000 mm/min and 30 %"
    # A single entry is not numbered anyway: nothing was skipped.
    _p, info1 = write([dict(LIST[0])], "cork", s=s)
    assert info1["numbers_skipped"] is False
    # Give cork engraving values and the numbers are back.
    s["gcode_settings"]["cork"].update(engraving_speed=3000, engraving_power=10)
    path2, info2 = write(LIST, "cork", s=s)
    assert info2["numbers_skipped"] is False and [n for n, _s in layers(path2)] == ["C03", "C02"]
    # Speed alone, or power alone, is still "not set".
    for speed, power in ((3000, 0), (0, 10)):
        s["gcode_settings"]["cork"].update(engraving_speed=speed, engraving_power=power)
        path3, info3 = write(LIST, "cork", s=s)
        assert info3["numbers_skipped"] is True and "; Layer C03" not in open(path3).read()
    # Filled mode reads the filled pair.
    s["gcode_settings"]["cork"].update(engraving_mode="filled", engraving_speed=3000, engraving_power=10,
                                       filled_engraving_speed=0, filled_engraving_power=0)
    _p4, info4 = write(LIST, "cork", s=s)
    assert info4["numbers_skipped"] is True


check("cork as shipped (engraving 0/0): every crescent cut, no numbers, no F0; numbers return with real values",
      cork_as_shipped_cuts_everything_and_engraves_nothing)


def a_material_with_no_cut_settings_is_refused():
    s = settings("acrylic", cut_power=0)
    try:
        write(LIST, s=s)
    except ValueError as e:
        assert not isinstance(e, se.CrescentKerfError)
        return
    raise AssertionError("crescents were cut at S0")


check("a material with no cut speed or power is refused, not cut at zero", a_material_with_no_cut_settings_is_refused)


# -------------------------------------------------------------------------- GUI

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
        path = os.path.join(OUT, "ui_" + kw.get("initialfile"))
        saved.append(path)
        return path
    tooling_tab.filedialog.asksaveasfilename = fake_save
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))
    note = ui_dialogs.gcode_file_note()

    def fields(d="43", pct="20", qty="1"):
        app.crescent_list.clear()
        app._refresh_crescent_list()
        app.crescent_tonehole_var.set(d)
        app.crescent_percent_var.set(pct)
        app.crescent_curve_var.set(75)
        app.crescent_qty_var.set(qty)

    def click(fn):
        del boxes[:]
        before = len(saved)
        fn()
        root.update()
        return saved[before:], list(boxes)

    def the_section_asks_for_a_material_cork_first():
        frame = app._tooling_sections["tonehole_crescent"]
        combos = [w for row in frame.winfo_children() for w in row.winfo_children() if isinstance(w, ttk.Combobox)]
        assert len(combos) == 1, combos
        assert list(combos[0].cget("values")) == ["Cork", "Acrylic", "Basswood", "Leather", "Felt", "Card"]
        assert app.crescent_material_var.get() == "Cork" and app._crescent_material_key() == "cork"
        app.crescent_material_var.set("Acrylic")
        app._update_tooling_settings()
        assert app.settings["tooling_settings"]["crescent_material"] == "acrylic", "stored as the key"
        app.crescent_material_var.set("Cork")
        app._update_tooling_settings()
        buttons = [str(w.cget("text")) for row in frame.winfo_children() for w in row.winfo_children()
                   if isinstance(w, tk.Button)]
        assert buttons.count("Generate G-code") == 1 and "Generate SVG" in buttons and "Generate STL" in buttons
    check("GUI: the section has a material choice, cork first and the default, saved as its key",
          the_section_asks_for_a_material_cork_first)

    def one_crescent_from_the_fields():
        fields()
        app.crescent_material_var.set("Cork")
        new, shown = click(app._on_generate_crescent_gcode)
        assert [os.path.basename(p) for p in new] == ["ui_crescent_43mm_20pct.gcode"], new
        want, _info = write([{"tonehole_d": 43.0, "percent": 20.0, "curve": 75.0, "qty": 1}], "cork", s=app.settings)
        assert open(new[0]).read() == open(want).read(), "the button's file is not the engine's output"
        done = [b for b in shown if b[0] == "showinfo"]
        assert len(done) == 1 and not [b for b in shown if b[0] == "showerror"], shown
        assert done[0][2] == f"Saved to: {new[0]}\n\n{note}", done[0][2]
    check("GUI: Generate G-code writes the engine's file for the crescent in the fields; plain Done box plus the laser line",
          one_crescent_from_the_fields)

    def a_cork_list_says_the_parts_will_not_be_numbered():
        fields()
        for d, pct, q in (("41.5", "10", "2"), ("43", "12", "1")):
            app.crescent_tonehole_var.set(d)
            app.crescent_percent_var.set(pct)
            app.crescent_qty_var.set(q)
            app._on_add_crescent()
        app.crescent_tonehole_var.set("20")          # the fields now hold something else: the list wins
        try:
            app.crescent_material_var.set("Cork")
            new, shown = click(app._on_generate_crescent_gcode)
            assert [os.path.basename(p) for p in new] == ["ui_tonehole_crescents.gcode"], new
            assert len(layers(new[0])[-1][1]) == 3, "2 + 1 crescents from the list"
            done = [b for b in shown if b[0] == "showinfo"]
            assert len(done) == 1
            msg = done[0][2]
            assert "will not be numbered" in msg and "Cork" in msg and "Tooling > Options > Settings" in msg, msg
            assert msg.endswith(note) and msg.startswith("Saved to: ")
            # Acrylic has engraving settings: numbered, and no such line.
            app.crescent_material_var.set("Acrylic")
            new, shown = click(app._on_generate_crescent_gcode)
            msg = [b for b in shown if b[0] == "showinfo"][0][2]
            assert "numbered" not in msg and msg.endswith(note)
            assert [n for n, _s in layers(new[0])] == ["C03", "C02"]
        finally:
            app.crescent_material_var.set("Cork")
            app.crescent_list.clear()
            app._refresh_crescent_list()
    check("GUI: a list cut from cork says in the Done box that the parts will not be numbered, and where to set it",
          a_cork_list_says_the_parts_will_not_be_numbered)

    def svg_and_stl_boxes_do_not_carry_the_laser_line():
        fields()
        for fn in (app._on_generate_crescent_svg, app._on_generate_crescent_stl):
            _new, shown = click(fn)
            done = [b for b in shown if b[0] == "showinfo"]
            assert len(done) == 1 and note not in done[0][2], done
    check("GUI: the SVG and STL Done boxes do not carry the G-code line", svg_and_stl_boxes_do_not_carry_the_laser_line)

    def refusals_in_plain_words_and_no_file():
        # Horns closer together than the kerf.
        fields(d="1.2", pct="75")
        app.crescent_curve_var.set(100)
        app.crescent_material_var.set("Acrylic")
        keep = app.settings["gcode_settings"]["acrylic"]["kerf_width"]
        app.settings["gcode_settings"]["acrylic"]["kerf_width"] = 1.0
        try:
            new, shown = click(app._on_generate_crescent_gcode)
        finally:
            app.settings["gcode_settings"]["acrylic"]["kerf_width"] = keep
        errs = [b for b in shown if b[0] == "showerror"]
        assert len(errs) == 1 and "narrower than the Acrylic kerf (1 mm)" in errs[0][2], shown
        assert not [b for b in shown if b[0] == "showinfo"] and not os.path.exists(new[0]), "a refused job left a file"
        # A bad field: refused before the save dialog, in the SVG's words.
        fields(d="0")
        new, shown = click(app._on_generate_crescent_gcode)
        g_err = [b for b in shown if b[0] == "showerror"]
        new_svg, shown_svg = click(app._on_generate_crescent_svg)
        assert new == [] and new_svg == [] and g_err == [b for b in shown_svg if b[0] == "showerror"] and g_err
        # No cut settings.
        fields()
        app.crescent_material_var.set("Cork")
        keep = app.settings["gcode_settings"]["cork"]["cut_power"]
        app.settings["gcode_settings"]["cork"]["cut_power"] = 0
        try:
            new, shown = click(app._on_generate_crescent_gcode)
        finally:
            app.settings["gcode_settings"]["cork"]["cut_power"] = keep
        errs = [b for b in shown if b[0] == "showerror"]
        assert len(errs) == 1 and "Cork has no cut speed and power set" in errs[0][2], shown
        assert not os.path.exists(new[0])
        # Any OTHER ValueError on the way (here a kerf that is not a number)
        # must not be dressed up as missing cut settings: it reaches the
        # user as what it is, through the handler's own error box.
        keep = app.settings["gcode_settings"]["cork"]["kerf_width"]
        app.settings["gcode_settings"]["cork"]["kerf_width"] = "wide"
        try:
            new, shown = click(app._on_generate_crescent_gcode)
        finally:
            app.settings["gcode_settings"]["cork"]["kerf_width"] = keep
        errs = [b for b in shown if b[0] == "showerror"]
        assert len(errs) == 1 and "has no cut speed and power set" not in errs[0][2],             f"a kerf of 'wide' was reported as missing cut settings: {shown}"
        assert "wide" in errs[0][2], f"the real problem is not named: {errs[0][2]!r}"
        assert not [b for b in shown if b[0] == "showinfo"] and not os.path.exists(new[0])
    check("GUI: a horn gap narrower than the kerf, a bad field, and a material with no cut settings are each refused; "
          "another ValueError is not called missing cut settings", refusals_in_plain_words_and_no_file)

    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
