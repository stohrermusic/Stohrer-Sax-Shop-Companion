"""Adhesive rings: fuzz the files that get written (2026-11-05).

test_adhesive_ring checks the helpers and the wiring on hand-picked sizes.
This one generates a few hundred random jobs and reads the SVG and G-code
FILES back, so a slip anywhere between the geometry helper and the bytes on
disk shows up. Seeded, so a failure names its case and reproduces.

Each job draws at random: a pad list (7.0 to 60 mm in half sizes, mixed
quantities), the sizing rules (program defaults, Matt's own preset from
presets/Matt mode.json, random universal values, random size ranges), darts
(universal, per-range, off), a ring width trim of either sign, kerf, cut
grouping, passes, compatibility mode, edge bias, labeled zones on or off,
and a rectangular sheet or a traced polygon (convex, concave, tilted).

Read back from the files, every job must satisfy:
  - one outer cut per ring, and its diameter is the card diameter
  - the inner cut follows wrap + trim (worked out here from the sizing
    rules, not taken from the helper under test); pads of 9.5 mm and
    under have none, nor does a larger pad whose hole would be under 2 mm
  - every ring has a number; an inside number is in its own hole with
    clearance, an outside number is under its own ring and clear of it
  - no ring's cut or number touches another ring's cut or number
  - everything is on the material (sheet, or inside the traced polygon)
  - SVG and G-code are the same points after the Y flip, kerf applied to
    the cuts only (inward inside, outward outside)
  - G-code order is number, inner, outer (per layer, or per ring), at the
    ring's cut settings, on the ring's one layer, laser off at the end

Headless: no Tk. Run:
    python tools/test_adhesive_ring_fuzz.py            # the standard run
    python tools/test_adhesive_ring_fuzz.py 400 7      # N cases, seed
"""
import copy
import json
import math
import os
import random
import re
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
os.environ.setdefault("SAXSHOP_CONFIG_DIR", tempfile.mkdtemp(prefix="ssc-ring-fuzz-"))

import config  # noqa: E402
import svg_engine as se  # noqa: E402
import gcode_engine as ge  # noqa: E402

RING = 'adhesive_ring'
N_CASES = int(sys.argv[1]) if len(sys.argv) > 1 else 160
SEED = int(sys.argv[2]) if len(sys.argv) > 2 else 20261105
TOL = 2.5e-3          # G-code is written to 3 decimals, SVG polylines to 4
results = []
stats = {"rings": 0, "inside": 0, "outside": 0, "solid": 0, "rect": 0, "polygon": 0,
         "matt": 0, "ranges": 0, "darts_off": 0, "dart_ranges": 0, "trim_neg": 0, "trim_pos": 0,
         "zones_on": 0, "by_pad": 0, "compat": 0, "partial": 0, "flip_override": 0, "sizes": set()}


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def matt_mode_preset():
    with open(os.path.join(REPO, "presets", "Matt mode.json"), encoding="utf-8") as f:
        raw = json.load(f)
    preset = raw if any(k in raw for k in config.SIZING_PRESET_KEYS) else raw[next(iter(raw))]
    return config.normalize_sizing_preset(preset)


# ------------------------------------------------- independent expectations ----

def expected_ring(pad, s):
    """(outer_r, inner_r or 0.0) for one ring, from the sizing rules alone.

    Deliberately does not call adhesive_ring_geometry or get_disc_diameter:
    card = pad - (felt offset + card-to-felt offset); width = the leather
    back wrap (base wrap, + the dart wrap bonus on a darted size, x the
    multiplier) + trim, never under 0.5. Pads of 9.5 mm and under are
    solid discs; above that a hole under 2 mm is dropped.
    """
    sizing = config.get_sizing_for_size(pad, s)
    outer_d = pad - (sizing["felt_offset"] + sizing["card_to_felt_offset"])
    dart = config.get_dart_settings_for_size(pad, s)
    bonus = dart.get("wrap_bonus", 0.75) if dart else 0.0
    if pad >= 45:
        base = 3.2
    elif pad >= 12:
        base = 1.2 + (pad - 12) * (2.0 / 33.0)
    elif pad >= 6:
        base = 1.0 + (pad - 6) * (0.2 / 6.0)
    else:
        base = 1.0
    width = max((base + bonus) * sizing["leather_wrap_multiplier"] + s["adhesive_ring_trim_mm"], 0.5)
    inner_d = outer_d - 2 * width
    return outer_d / 2, (inner_d / 2 if pad > 9.5 and inner_d >= 2.0 else 0.0)


def inside_polygon(x, y, poly):
    """Ray casting, written here rather than borrowed from svg_engine."""
    inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]
        if (y1 > y) != (y2 > y):
            if x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
                inside = not inside
    return inside


# ----------------------------------------------------------- file readers ----

def read_svg(path):
    text = open(path, encoding="utf-8").read()
    root = ET.fromstring(text)
    polys, circles, kinds, colors = [], [], set(), set()
    for el in root.iter():
        tag = el.tag.split('}')[-1]
        if tag in ("svg", "defs"):
            continue
        kinds.add(tag)
        colors.add(el.get("stroke"))
        if tag == "polyline":
            polys.append([tuple(float(v) for v in pt.split(',')) for pt in el.get("points").split()])
        elif tag == "circle":
            circles.append(tuple(float(el.get(k).replace("mm", "")) for k in ("cx", "cy", "r")))
    return text, root, polys, circles, kinds, colors


def read_gcode(path):
    text = open(path, encoding="utf-8").read()
    kinds = {"; Ring number": "number", "; Ring inner cut": "inner", "; Ring outer cut": "outer"}
    sections, cur = [], None
    for ln in text.splitlines():
        if ln.strip() in kinds:
            cur = (kinds[ln.strip()], [])
            sections.append(cur)
            continue
        if ln.startswith("; return to origin"):
            cur = None
        if cur is None:
            continue
        m = re.match(r"G([01]) X(-?[\d.]+)Y(-?[\d.]+)", ln)
        if not m:
            continue
        pt = (float(m.group(2)), float(m.group(3)))
        if m.group(1) == "0":
            cur[1].append([pt])
        else:
            cur[1][-1].append(pt)
    return text, sections


def circle_of(stroke):
    xs = [p[0] for p in stroke]
    ys = [p[1] for p in stroke]
    cx, cy = (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2
    r = (max(xs) - min(xs)) / 2
    for x, y in stroke:
        assert abs(math.hypot(x - cx, y - cy) - r) < TOL, "a cut in the G-code is not a circle"
    return cx, cy, r


# ------------------------------------------------------------ the checker ----

def verify_job(tag, s, placed, w, h, polygon, svg_path, gc_path, flip_height=None):
    """Assert every invariant on one job's written files.

    flip_height is what the G-code was flipped with when Frame & Cut's
    override was used (a camera scrap's outline is taller than the inset
    polygon the rings nest in); otherwise the polygon's or sheet's height."""
    text, root, polys, circles, kinds, colors = read_svg(svg_path)
    assert root.tag.split('}')[-1] == "svg", f"{tag}: not an SVG"
    assert kinds <= {"polyline", "circle"}, f"{tag}: unexpected SVG elements {kinds}"
    assert colors == {s["layer_colors"]["adhesive_ring_cut"]}, f"{tag}: off-layer strokes {colors}"
    if s.get("compatibility_mode"):
        assert 'mm"' not in text.split("<defs", 1)[1], f"{tag}: unit suffix in compatibility mode"

    kerf = s["gcode_settings"][RING]["kerf_width"] / 2
    passes = s["gcode_settings"][RING]["cut_passes"]
    flip = flip_height if flip_height is not None else (max(p[1] for p in polygon) if polygon else h)

    # --- rings, from the SVG circles ---------------------------------
    rings = []
    for pad, cx, cy, foot_r in placed:
        mine = sorted(r for x, y, r in circles if abs(x - cx) < 1e-6 and abs(y - cy) < 1e-6)
        want_outer, want_inner = expected_ring(pad, s)
        assert mine, f"{tag}: no cut for the {pad} ring at {(cx, cy)}"
        assert abs(mine[-1] - want_outer) < 1e-6, \
            f"{tag}: {pad} cut r {mine[-1]:.4f}, card r is {want_outer:.4f}"
        if want_inner > 0:
            assert len(mine) == 2 and abs(mine[0] - want_inner) < 1e-6, \
                f"{tag}: {pad} inner r {mine[:-1]}, wrap + trim gives {want_inner:.4f}"
        else:
            assert len(mine) == 1, f"{tag}: {pad} should be a solid disc, got circles {mine}"
            stats["solid"] += 1
        rings.append({"pad": pad, "cx": cx, "cy": cy, "foot": foot_r,
                      "outer": want_outer, "inner": want_inner, "number": []})
    assert len(circles) == sum(2 if r["inner"] > 0 else 1 for r in rings), f"{tag}: stray circles"

    # --- numbers: each stroke belongs to exactly one ring ------------------
    for stroke in polys:
        owners = [r for r in rings
                  if all(math.hypot(x - r["cx"], y - r["cy"]) <= r["foot"] + 1e-3 for x, y in stroke)]
        assert len(owners) == 1, f"{tag}: a number stroke lies in {len(owners)} ring footprints"
        owners[0]["number"].append(stroke)
    for r in rings:
        assert r["number"], f"{tag}: the {r['pad']} ring has no number"
        pts = [p for st in r["number"] for p in st]
        dists = [math.hypot(x - r["cx"], y - r["cy"]) for x, y in pts]
        height = max(y for _x, y in pts) - min(y for _x, y in pts)
        if max(dists) < r["outer"]:
            # inside: in its own hole, 1 mm clear of the inner cut, legible
            assert r["inner"] > 0 and max(dists) <= r["inner"] - 1.0 + 1e-3, \
                f"{tag}: {r['pad']} number is {r['inner'] - max(dists):.2f} mm from its inner cut"
            assert height >= 0.9 * 2.5 - 1e-3, f"{tag}: {r['pad']} inside number only {height:.2f} mm tall"
            # ... centered in the hole, and as large as fits (or at the 5 mm cap)
            bx = (min(x for x, _y in pts) + max(x for x, _y in pts)) / 2
            by = (min(y for _x, y in pts) + max(y for _x, y in pts)) / 2
            assert abs(bx - r["cx"]) < 1e-3 and abs(by - r["cy"]) < 1e-3, f"{tag}: {r['pad']} number off-center"
            assert abs(height - 0.9 * 5.0) < 1e-3 or abs(max(dists) - (r["inner"] - 1.0)) < 1e-3,                 f"{tag}: {r['pad']} inside number ({height / 0.9:.2f} mm) is neither capped nor filling its hole"
            stats["inside"] += 1
        else:
            # outside: all of it beyond the cut, under the ring (Y-down), centered
            assert min(dists) >= r["outer"] + 0.5, f"{tag}: {r['pad']} outside number touches its ring"
            top = min(y for _x, y in pts)
            assert abs(top - (r["cy"] + r["outer"] + 1.0)) < 1e-3, \
                f"{tag}: {r['pad']} number top {top - r['cy'] - r['outer']:.3f} mm under the cut, wanted 1.0"
            xs = [x for x, _y in pts]
            assert abs((min(xs) + max(xs)) / 2 - r["cx"]) < 1e-3, f"{tag}: {r['pad']} number not centered"
            assert abs(height - 0.9 * 2.5) < 1e-3, f"{tag}: {r['pad']} outside number {height:.2f} mm tall"
            # ... and only because the hole could not take it at this height.
            mx, my = (min(xs) + max(xs)) / 2, (min(y for _x, y in pts) + max(y for _x, y in pts)) / 2
            reach = max(math.hypot(x - mx, y - my) for x, y in pts)
            assert r["inner"] - 1.0 < reach - 1e-6,                 f"{tag}: {r['pad']} number is outside though its {2 * r['inner']:.2f} mm hole would hold it"
            stats["outside"] += 1
        r["box"] = (min(x for x, _y in pts), min(y for _x, y in pts),
                    max(x for x, _y in pts), max(y for _x, y in pts))
        # The strokes ARE this ring's pad size in the stroke font: the right
        # digits, single lines, upright and unmirrored (SVG is Y-down, the
        # font is drawn Y-up).
        label = f"{r['pad']:.1f}".rstrip('0').rstrip('.')
        unit = ge.get_text_strokes(label, 1.0, 0.0, 0.0)
        upts = [p for st in unit for p in st]
        ux = (min(p[0] for p in upts) + max(p[0] for p in upts)) / 2
        uy = (min(p[1] for p in upts) + max(p[1] for p in upts)) / 2
        font = height / (max(p[1] for p in upts) - min(p[1] for p in upts))
        bx, by = (r["box"][0] + r["box"][2]) / 2, (r["box"][1] + r["box"][3]) / 2
        assert len(r["number"]) == len(unit), \
            f"{tag}: {r['pad']} number has {len(r['number'])} strokes, '{label}' has {len(unit)}"
        for got, want in zip(r["number"], unit):
            assert len(got) == len(want), f"{tag}: {r['pad']} number is not '{label}' in the stroke font"
            for (gx, gy), (wx, wy) in zip(got, want):
                assert abs(gx - (bx + (wx - ux) * font)) < 1e-3 and abs(gy - (by - (wy - uy) * font)) < 1e-3, \
                    f"{tag}: {r['pad']} number is not an upright '{label}'"
        stats["rings"] += 1
        stats["sizes"].add(r["pad"])

    # --- nothing touches anything else ---------------------------------
    for i, a in enumerate(rings):
        for b in rings[i + 1:]:
            gap = math.hypot(a["cx"] - b["cx"], a["cy"] - b["cy"]) - a["outer"] - b["outer"]
            assert gap >= 0.9, f"{tag}: {a['pad']} and {b['pad']} cuts are {gap:.2f} mm apart"
            ax0, ay0, ax1, ay1 = a["box"]
            bx0, by0, bx1, by1 = b["box"]
            apart = ax1 + 0.5 < bx0 or bx1 + 0.5 < ax0 or ay1 + 0.5 < by0 or by1 + 0.5 < ay0
            assert apart, f"{tag}: the numbers of {a['pad']} and {b['pad']} overlap"
        for b in rings:
            if b is a:
                continue
            for stroke in a["number"]:
                for x, y in stroke:
                    gap = math.hypot(x - b["cx"], y - b["cy"]) - b["outer"]
                    assert gap >= 0.9, f"{tag}: {a['pad']}'s number is {gap:.2f} mm from a {b['pad']} cut"

    # --- on the material -------------------------------------------------
    def on_material(x, y, what):
        if polygon:
            assert inside_polygon(x, y, polygon), f"{tag}: {what} off the scrap at ({x:.2f}, {y:.2f})"
        else:
            assert 0 <= x <= w and 0 <= y <= h, f"{tag}: {what} off the sheet at ({x:.2f}, {y:.2f})"
    for r in rings:
        for k in range(48):
            a = 2 * math.pi * k / 48
            on_material(r["cx"] + (r["outer"] + kerf) * math.cos(a),
                        r["cy"] + (r["outer"] + kerf) * math.sin(a), f"{r['pad']} cut")
        for stroke in r["number"]:
            for x, y in stroke:
                on_material(x, y, f"{r['pad']} number")

    # --- the G-code file ---------------------------------------------------
    gtext, sections = read_gcode(gc_path)
    lines = gtext.splitlines()
    assert lines[-2:] == ["M5", "M2"], f"{tag}: file does not end with the laser off: {lines[-3:]}"
    assert "G1 S0" in lines and "M4" in lines and "M3" not in lines, f"{tag}: laser mode"
    assert not re.search(r"; Layer C(?!07)", gtext), f"{tag}: a layer other than the ring's"
    cfg = s["gcode_settings"][RING]
    heads = {ln for ln in lines if ln.startswith("; Cut @")}
    note = f", x{passes} passes" if passes > 1 else ""
    assert heads == {f"; Cut @ {cfg['cut_speed']} mm/min, {cfg['cut_power']}% power{note}"}, \
        f"{tag}: cut settings in the file: {heads}"

    order = [k for k, _st in sections]
    if s["gcode_cut_grouping"] == "pad":
        want = []
        for r in rings:
            want += ["number"] + (["inner"] if r["inner"] - kerf > 0 and r["inner"] > 0 else []) + ["outer"]
        assert order == want, f"{tag}: per-pad order {order[:9]}..."
    else:
        want = ["number"] + (["inner"] if any(r["inner"] > 0 for r in rings) else []) + ["outer"]
        assert order == want, f"{tag}: layer order {order}"

    def once(kind):
        """A step's strokes with repeated passes folded away, file order."""
        out = []
        for k, strokes in sections:
            if k == kind:
                assert len(strokes) % passes == 0, f"{tag}: {kind} strokes not a multiple of the passes"
                n = len(strokes) // passes
                for p in range(1, passes):
                    assert strokes[p * n:(p + 1) * n] == strokes[:n], f"{tag}: pass {p + 1} differs from pass 1"
                out.extend(strokes[:n])
        return out

    g_numbers = once("number")
    assert len(g_numbers) == len(polys), f"{tag}: {len(g_numbers)} G-code number strokes, SVG has {len(polys)}"
    for sp, gp in zip(polys, g_numbers):
        assert len(sp) == len(gp), f"{tag}: number stroke lengths differ"
        for (sx, sy), (gx, gy) in zip(sp, gp):
            assert abs(sx - gx) < TOL and abs((flip - sy) - gy) < TOL, \
                f"{tag}: number point SVG ({sx:.3f}, {sy:.3f}) vs G-code ({gx:.3f}, {gy:.3f})"

    g_outer = [circle_of(st) for st in once("outer")]
    g_inner = [circle_of(st) for st in once("inner")]
    # Both cuts of a ring are equally round: 72 segments each, closed.
    for kind in ("outer", "inner"):
        for st in once(kind):
            assert len(st) == 73 and st[0] == st[-1], f"{tag}: an {kind} cut has {len(st) - 1} segments"
    assert len(g_outer) == len(rings), f"{tag}: {len(g_outer)} outer cuts for {len(rings)} rings"
    for r, (gx, gy, gr) in zip(rings, g_outer):
        assert abs(gx - r["cx"]) < TOL and abs(gy - (flip - r["cy"])) < TOL, f"{tag}: {r['pad']} outer cut misplaced"
        assert abs(gr - (r["outer"] + kerf)) < TOL, \
            f"{tag}: {r['pad']} G-code cut r {gr:.3f}, card r + half kerf is {r['outer'] + kerf:.3f}"
    holed = [r for r in rings if r["inner"] > 0 and r["inner"] - kerf > 0]
    assert len(g_inner) == len(holed), f"{tag}: {len(g_inner)} inner cuts for {len(holed)} holes"
    for r, (gx, gy, gr) in zip(holed, g_inner):
        assert abs(gx - r["cx"]) < TOL and abs(gy - (flip - r["cy"])) < TOL, f"{tag}: {r['pad']} inner cut misplaced"
        assert abs(gr - (r["inner"] - kerf)) < TOL, f"{tag}: {r['pad']} inner cut r {gr:.3f}"

    for _k, strokes in sections:
        for st in strokes:
            for x, y in st:
                on_material(x, flip - y, "G-code move")
    stats["flip_override"] += flip_height is not None
    return len(rings)


# ------------------------------------------------------------ job builder ----

SIZES = [7.0 + 0.5 * i for i in range(107)]          # 7.0 ... 60.0
BIASES = ("center", "off", "n", "ne", "e", "se", "s", "sw", "w", "nw")


def random_settings(rng, matt):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    s["adhesive_ring_enabled"] = True
    kind = rng.choice(["default", "matt", "matt", "universal", "ranges"])
    if kind == "matt":
        s.update(copy.deepcopy(matt))
        s["adhesive_ring_enabled"] = True
        stats["matt"] += 1
    elif kind == "universal":
        s["felt_offset"] = round(rng.uniform(0.4, 1.2), 2)
        s["card_to_felt_offset"] = round(rng.uniform(0.0, 2.5), 2)
        s["leather_wrap_multiplier"] = round(rng.uniform(0.8, 1.3), 2)
    elif kind == "ranges":
        cut1, cut2 = sorted(rng.sample([10.0, 12.5, 16.0, 18.0, 24.0, 30.0, 40.0], 2))
        s["sizing_range_mode"] = "range"
        s["sizing_ranges"] = [
            {"min_size": lo, "max_size": hi,
             "felt_offset": round(rng.uniform(0.4, 1.2), 2),
             "card_to_felt_offset": round(rng.uniform(0.0, 2.0), 2),
             "leather_wrap_multiplier": round(rng.uniform(0.8, 1.3), 2),
             "min_hole_size": 16.5, "felt_thickness": 3.175, "felt_thickness_unit": "mm"}
            for lo, hi in ((0.0, cut1), (cut1 + 0.1, cut2))]        # above cut2: universal fallback
        stats["ranges"] += 1
    if kind != "matt":
        darts = rng.choice(["universal", "off", "range"])
        if darts == "off":
            s["darts_enabled"] = False
            stats["darts_off"] += 1
        elif darts == "universal":
            s["dart_threshold"] = rng.choice([12.0, 18.0, 25.0, 40.0])
            s["dart_wrap_bonus"] = round(rng.uniform(0.2, 1.2), 2)
        else:
            s["dart_range_mode"] = "range"
            s["dart_ranges"] = [
                {"min_size": 0.0, "max_size": 11.5, "overwrap": 0.5, "wrap_bonus": round(rng.uniform(0.2, 1.2), 2),
                 "frequency_multiplier": 1.0, "shape_factor": 0.5, "engraving_on": True},
                {"min_size": 14.0, "max_size": 22.0, "overwrap": 0.5, "wrap_bonus": round(rng.uniform(0.2, 1.2), 2),
                 "frequency_multiplier": 1.0, "shape_factor": 0.5, "engraving_on": True}]
            stats["dart_ranges"] += 1
    trim = rng.choice([0.0, round(rng.uniform(-1.2, -0.05), 2), round(rng.uniform(0.05, 1.5), 2)])
    s["adhesive_ring_trim_mm"] = trim
    stats["trim_neg" if trim < 0 else "trim_pos"] += 1 if trim else 0
    s["gcode_settings"][RING].update(
        cut_speed=rng.choice([600, 1500, 2400]), cut_power=rng.choice([20, 45, 62.5]),
        cut_passes=rng.choice([1, 1, 2]), kerf_width=rng.choice([0.0, 0.1, 0.2, 0.4]),
        air_assist_cut=rng.choice([True, False]))
    s["gcode_cut_grouping"] = rng.choice(["layer", "layer", "pad"])
    s["compatibility_mode"] = rng.random() < 0.25
    s["edge_bias"] = rng.choice(BIASES)
    s["zone_labels_enabled"] = rng.random() < 0.3       # must stand down for rings
    for mat in ("felt", "card", "leather"):
        s["gcode_settings"][mat]["engraving_mode"] = rng.choice(["line", "filled"])
    stats["by_pad"] += s["gcode_cut_grouping"] == "pad"
    stats["compat"] += s["compatibility_mode"]
    stats["zones_on"] += s["zone_labels_enabled"]
    return s


def random_pads(rng):
    n = rng.randint(3, 8)
    small = [z for z in SIZES if z <= 14.5]
    sizes = set(rng.sample(small, rng.randint(1, min(4, n))))
    while len(sizes) < n:
        sizes.add(rng.choice(SIZES))
    return [{'size': z, 'qty': rng.randint(1, 6 if z < 20 else 3 if z < 40 else 2)} for z in sorted(sizes)]


def random_sheet(rng, pads, s):
    """(w, h, polygon) with room for the job: about 2.4x the footprints' squares."""
    need = sum(p['qty'] * (se.get_disc_diameter(p['size'], RING, s) + 1.0) ** 2 for p in pads)
    biggest = max(se.get_disc_diameter(p['size'], RING, s) for p in pads) + 4.0
    area = need * rng.uniform(1.9, 3.0)
    aspect = rng.uniform(0.6, 1.7)
    w = max(math.sqrt(area * aspect), biggest * 1.3)
    h = max(area / w, biggest * 1.3)
    if rng.random() < 0.5:
        return round(w, 1), round(h, 1), None
    shape = rng.choice(["L", "trapezoid", "notch", "tilted"])
    if shape == "L":
        w, h = w * 1.25, h * 1.25
        cx, cy = w * rng.uniform(0.5, 0.7), h * rng.uniform(0.5, 0.7)
        poly = [(0, 0), (w, 0), (w, cy), (cx, cy), (cx, h), (0, h)]
    elif shape == "trapezoid":
        w = w * 1.15
        poly = [(0, 0), (w, 0), (w * rng.uniform(0.8, 0.95), h), (w * rng.uniform(0.05, 0.2), h)]
    elif shape == "notch":
        w, h = w * 1.1, h * 1.1
        poly = [(0, 0), (w, 0), (w, h), (w * 0.62, h), (w * 0.5, h * 0.62), (w * 0.38, h), (0, h)]
    else:
        ang = math.radians(rng.uniform(-8, 8))
        corners = [(0, 0), (w, 0), (w, h), (0, h)]
        rot = [(x * math.cos(ang) - y * math.sin(ang), x * math.sin(ang) + y * math.cos(ang)) for x, y in corners]
        mx, my = min(p[0] for p in rot), min(p[1] for p in rot)
        poly = [(x - mx, y - my) for x, y in rot]
    poly = [(round(x, 3), round(y, 3)) for x, y in poly]
    return max(p[0] for p in poly), max(p[1] for p in poly), poly


def run_case(i, matt, tmp):
    rng = random.Random(SEED * 1000 + i)
    s = random_settings(rng, matt)
    pads = random_pads(rng)
    w, h, polygon = random_sheet(rng, pads, s)
    tag = f"case {i} (seed {SEED})"
    placed, zones, fp, ft = se.nest_with_zones(pads, RING, w, h, s, polygon=polygon)
    assert zones == [], f"{tag}: zones produced for rings"
    if fp != ft:
        stats["partial"] += 1                 # too tight a sheet: check what did fit
    assert len(placed) >= max(1, ft // 2), f"{tag}: only {len(placed)} of {ft} placed on {w}x{h}"
    sp, gp = os.path.join(tmp, f"c{i}.svg"), os.path.join(tmp, f"c{i}.gcode")
    se.generate_svg_from_placed(placed, RING, w, h, sp, 3.5, s, polygon=polygon, zones=zones)
    # Frame & Cut flips a camera scrap with its OUTLINE's height, which is
    # a few mm more than the inset polygon's: try that on half the polygons.
    flip_override = None
    if polygon and rng.random() < 0.5:
        flip_override = round(max(p[1] for p in polygon) + rng.uniform(1.0, 8.0), 2)
    ge.generate_gcode_from_placed(placed, RING, w, h, gp, 3.5, s, polygon=polygon, zones=zones,
                                  flip_height_mm=flip_override)
    stats["polygon" if polygon else "rect"] += 1
    return verify_job(tag, s, placed, w, h, polygon, sp, gp, flip_height=flip_override)


def main():
    print(f"Adhesive rings: fuzzing the written files ({N_CASES} jobs, seed {SEED})")
    print("=" * 60)
    matt = matt_mode_preset()
    failures = []
    with tempfile.TemporaryDirectory() as tmp:
        for i in range(N_CASES):
            try:
                run_case(i, matt, tmp)
            except Exception as e:
                failures.append(f"{type(e).__name__}: {e}")
                if len(failures) <= 5:
                    traceback.print_exc()

        def all_jobs_hold():
            assert not failures, f"{len(failures)} of {N_CASES} jobs broke an invariant; first: {failures[0]}"
        check(f"{N_CASES} random jobs: every invariant holds on the written SVG and G-code", all_jobs_hold)

        def the_run_covered_the_ground():
            # A fuzz run that never drew a polygon or a negative trim proves nothing about them.
            for key, least in (("rings", N_CASES * 5), ("inside", N_CASES), ("outside", N_CASES),
                               ("solid", 5), ("rect", N_CASES // 4), ("polygon", N_CASES // 4),
                               ("matt", N_CASES // 6), ("ranges", 5), ("darts_off", 5), ("dart_ranges", 5),
                               ("trim_neg", N_CASES // 6), ("trim_pos", N_CASES // 6), ("zones_on", 5),
                               ("by_pad", 5), ("compat", 5), ("flip_override", N_CASES // 8)):
                assert stats[key] >= least, f"{key}: only {stats[key]}, wanted at least {least}"
            assert len(stats["sizes"]) >= 90, f"only {len(stats['sizes'])} of 107 half sizes drawn"
        check("the run covered both sheet types, Matt mode, ranges, darts off, both trim signs, "
              f"{len(stats['sizes'])} sizes, {stats['rings']} rings "
              f"({stats['inside']} numbered inside, {stats['outside']} outside, {stats['solid']} solid)",
              the_run_covered_the_ground)

        # ---- Matt's own preset, every half size, read back from the files ----
        def matt_mode_every_half_size():
            s = copy.deepcopy(config.DEFAULT_SETTINGS)
            s.update(copy.deepcopy(matt))
            s["adhesive_ring_enabled"] = True
            s["adhesive_ring_trim_mm"] = 0.0
            total, outside, solid = 0, [], []
            for chunk, (w, h, polygon) in enumerate((
                    (260.0, 200.0, None), (300.0, 260.0, None), (420.0, 400.0, None),
                    (330.0, 330.0, [(0, 0), (330, 0), (330, 150), (170, 150), (170, 330), (0, 330)]),
                    (400.0, 400.0, [(0, 0), (400, 0), (380, 400), (30, 400)]))):
                group = SIZES[chunk::5]
                pads = [{'size': z, 'qty': 2 if z < 30 else 1} for z in group]
                placed, zones, fp, ft = se.nest_with_zones(pads, RING, w, h, s, polygon=polygon)
                assert fp == ft, f"Matt mode chunk {chunk}: {fp} of {ft} placed"
                sp, gp = os.path.join(tmp, f"m{chunk}.svg"), os.path.join(tmp, f"m{chunk}.gcode")
                se.generate_svg_from_placed(placed, RING, w, h, sp, 3.5, s, polygon=polygon)
                ge.generate_gcode_from_placed(placed, RING, w, h, gp, 3.5, s, polygon=polygon)
                total += verify_job(f"Matt mode chunk {chunk}", s, placed, w, h, polygon, sp, gp)
                for z in group:
                    _o, inner = expected_ring(z, s)
                    geom = se.adhesive_ring_geometry(z, s)
                    if geom["placement"] == "outside":
                        outside.append(z)
                    if inner == 0:
                        solid.append(z)
            assert total == sum(2 if z < 30 else 1 for z in SIZES), total
            print(f"        Matt mode, 7.0-60 in half sizes: number outside on {sorted(outside)}; "
                  f"solid discs: {sorted(solid) or 'none'}")
            assert all(z <= 16.0 for z in outside), "an outside number on a big ring under Matt mode"
            assert sorted(solid) == [7.0, 7.5, 8.0, 8.5, 9.0, 9.5], f"Matt mode solid discs: {sorted(solid)}"
        check("Matt mode preset: all 107 half sizes from 7.0 to 60, rectangles and polygons", matt_mode_every_half_size)

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == '__main__':
    sys.exit(main())
