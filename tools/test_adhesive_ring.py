"""Adhesive rings (beta, 2026-11-05).

An opt-in fifth Pad Maker material: a ring of double-sided tape per pad, for
sticking a thin silicone wrap to the back of the card. Outer cut = the card
disc; ring width = the leather disc's back wrap for that size plus a trim;
the pad size is traced as single-stroke digits in the hole (on the waste
plug) or, where the hole is too small to read it, on the scrap just under
the ring. Lives behind Sizing Rules > Lesser-used settings.

Covers:
  - geometry: outer = card, width = leather wrap (dart bonus included) +
    trim of either sign; pads of 9.5 mm and under are solid discs, and so
    is any larger pad whose hole would be under 2 mm
  - the number: every ring gets one; inside and clear of the inner cut on
    big rings, outside and under the ring on small rings and solid discs;
    one constant decides; never filled, whatever engraving_mode says
  - nesting: the footprint holds ring and number; an outside number is
    clear of every neighbour's cut and on the material, on a rectangle and
    on a polygon; zones stand down; scrap-mode partial nests work
  - output: the SVG parses with ring and number on the ring's layer and no
    <text>; the cut radius is the card radius even where the footprint is
    larger; SVG and G-code draw the same points after the Y flip, the
    number upright and under the ring in both; G-code goes number, inner,
    outer, stays on the sheet, uses the ring's cut settings and kerf, and
    ends with the laser off; flip_height_mm is honoured
  - feature off (or keys absent) leaves the other materials' SVG and G-code
    byte-identical
  - presets: sizing keys captured and backfilled; G-code preset material
    with cut-only keys and the Default backfill; old settings files load
  - framing power: own entry, clamped
  - GUI (self-skips without a display): the enabling control is marked
    beta; "Wrap clearance" is a spinbox (0.1 steps, -5.0 to 5.0, tooltip on
    label, counter and unit) whose +1.0 / -1.0 move the inner cut in the
    written SVG by 2.0 mm of diameter, whose every tenth survives Apply,
    save and reopen, and which survives typed nonsense; the material
    checkbox, Layer Colors row, G-code section and Framing Power row exist
    only while enabled; the pad preview and the
    nesting preview draw the real ring; Frame & Cut takes rings as its one
    material and streams exactly what the file export writes, framing at
    the ring's own power with M4, in standard, polygon and scrap mode; job
    history records and reloads the material

Uses an isolated profile (SAXSHOP_CONFIG_DIR): never touches real settings.
"""
import copy
import glob
import json
import math
import os
import re
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-adhesive-ring-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

import config  # noqa: E402
import svg_engine as se  # noqa: E402
import gcode_engine as ge  # noqa: E402
from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

RING = 'adhesive_ring'
RING_COLOR = config.DEFAULT_SETTINGS["layer_colors"]["adhesive_ring_cut"]
EPS = 1e-6
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


def base_settings(**over):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    s.update(over)
    return s


def on(**over):
    return base_settings(adhesive_ring_enabled=True, **over)


# Real pad sizes. 7.0 and 8.5 are solid discs (9.5 and under always are),
# 10.0 is a ring whose hole is too small for a legible number, 12.5 has a
# wide label that doesn't fit its hole, and 16.0 / 20.0 / 35.5 take the
# number inside. Everything under 18 is a darted size on leather.
SIZES = (7.0, 8.5, 10.0, 12.5, 16.0, 20.0, 35.5)
OUTSIDE = (7.0, 8.5, 10.0, 12.5)
INSIDE = (16.0, 20.0, 35.5)
# Inside, but in a hole small enough that the clearance (not the 5 mm cap)
# sets the number's height.
INSIDE_TIGHT = (12.0, 13.0, 15.5)
PADS = [{'size': 7.0, 'qty': 4}, {'size': 8.5, 'qty': 3}, {'size': 10.0, 'qty': 3},
        {'size': 12.5, 'qty': 3}, {'size': 16.0, 'qty': 2}, {'size': 20.0, 'qty': 2},
        {'size': 35.5, 'qty': 1}]
W, H = 120.0, 90.0
# A concave (L-shaped) offcut, Y-down like custom_polygon.
SCRAP = [(0.0, 0.0), (110.0, 0.0), (110.0, 45.0), (55.0, 45.0), (55.0, 95.0), (0.0, 95.0)]


def expected_wrap(pad, s):
    """The leather disc's back wrap, computed the way get_disc_diameter does
    for leather, without going through the helper under test."""
    sizing = config.get_sizing_for_size(pad, s)
    dart = config.get_dart_settings_for_size(pad, s)
    bonus = dart.get("wrap_bonus", 0.75) if dart else 0.0
    return se.leather_back_wrap(pad, sizing["leather_wrap_multiplier"], extra_base=bonus)


def nest(s, pads=PADS, w=W, h=H, polygon=None):
    placed, zones, fp, ft = se.nest_with_zones(pads, RING, w, h, s, polygon=polygon)
    return placed, zones, fp, ft


def render(s, placed, material=RING, w=W, h=H, hole=3.5, **kw):
    with tempfile.TemporaryDirectory() as d:
        sp, gp = os.path.join(d, 'a.svg'), os.path.join(d, 'a.gcode')
        se.generate_svg_from_placed(placed, material, w, h, sp, hole, s)
        ge.generate_gcode_from_placed(placed, material, w, h, gp, hole, s, **kw)
        return open(sp, encoding='utf-8').read(), open(gp, encoding='utf-8').read()


def gcode_sections(text):
    """[(kind, [stroke, ...]), ...] in file order, kind in number / inner /
    outer, each stroke a list of (x, y). A repeated pass repeats strokes."""
    kinds = {"; Ring number": "number", "; Ring inner cut": "inner", "; Ring outer cut": "outer"}
    out, cur = [], None
    for ln in text.splitlines():
        if ln.strip() in kinds:
            cur = (kinds[ln.strip()], [])
            out.append(cur)
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
    return out


def svg_parts(text):
    """(polylines, circles) from a ring SVG: polylines as point lists,
    circles as (cx, cy, r), with every element's stroke color checked by
    the caller via svg_strokes()."""
    root = ET.fromstring(text)
    polys, circles = [], []
    for el in root.iter():
        tag = el.tag.split('}')[-1]
        if tag == "polyline":
            polys.append([tuple(float(v) for v in pt.split(',')) for pt in el.get("points").split()])
        elif tag == "circle":
            circles.append(tuple(float(el.get(k).replace('mm', '')) for k in ("cx", "cy", "r")))
    return polys, circles


def svg_strokes(text):
    root = ET.fromstring(text)
    return [(el.tag.split('}')[-1], el.get("stroke")) for el in root.iter()
            if el.tag.split('}')[-1] in ("polyline", "circle", "path", "line", "rect", "text")]


# ------------------------------------------------------------ defaults ----

def test_off_by_default_and_schema():
    d = config.DEFAULT_SETTINGS
    assert d["adhesive_ring_enabled"] is False
    assert d["adhesive_ring_trim_mm"] == 0.0
    assert RING in d["gcode_settings"] and RING in config.GCODE_PRESET_MATERIALS
    assert set(d["gcode_settings"][RING]) == set(config.GCODE_CUT_ONLY_KEYS), d["gcode_settings"][RING]
    assert d["laser_framing_power_by_material"][RING] == 10
    assert RING_COLOR in {h for _n, h in config.LIGHTBURN_COLORS}
    taken = {v for k, v in d["layer_colors"].items()
             if k != "adhesive_ring_cut" and not k.startswith("die_")}
    assert RING_COLOR not in taken, "ring layer shares a color with another pad layer"


# ------------------------------------------------------------ geometry ----

def test_outer_is_the_card_disc_and_width_is_the_leather_wrap():
    s = on()
    darted = [p for p in SIZES if config.get_dart_settings_for_size(p, s)]
    plain = [p for p in SIZES if not config.get_dart_settings_for_size(p, s)]
    assert darted and plain, "fixture needs darted and plain sizes"
    for pad in SIZES:
        g = se.adhesive_ring_geometry(pad, s)
        card_r = se.get_disc_diameter(pad, 'card', s) / 2
        assert math.isclose(g["outer_r"], card_r, abs_tol=EPS), f"{pad}: outer {g['outer_r']} != card {card_r}"
        wrap = expected_wrap(pad, s)
        assert math.isclose(g["width"], wrap, abs_tol=EPS), f"{pad}: width {g['width']} != wrap {wrap}"
        if pad > se.ADHESIVE_RING_SOLID_MAX_PAD_MM and 2 * (card_r - wrap) >= se.ADHESIVE_RING_MIN_HOLE_MM:
            assert math.isclose(g["inner_r"], card_r - wrap, abs_tol=EPS), f"{pad}: inner"
        else:
            assert g["inner_r"] == 0.0, f"{pad}: should be a solid disc"
    assert se.adhesive_ring_geometry(10.0, s)["inner_r"] > 0, "fixture: SIZES needs a small ring with a hole"
    # The dart wrap bonus is in the ring: same size with darts off is narrower by it.
    pad = 10.0
    bonus = config.get_dart_settings_for_size(pad, s)["wrap_bonus"]
    no_darts = se.adhesive_ring_geometry(pad, on(darts_enabled=False))
    with_darts = se.adhesive_ring_geometry(pad, s)
    assert bonus > 0
    assert math.isclose(with_darts["width"] - no_darts["width"], bonus * s["leather_wrap_multiplier"], abs_tol=EPS)
    # The wrap multiplier carries through.
    wide = se.adhesive_ring_geometry(20.0, on(leather_wrap_multiplier=1.5))
    assert math.isclose(wide["width"], 1.5 * se.adhesive_ring_geometry(20.0, s)["width"], abs_tol=EPS)


def test_trim_of_both_signs():
    pad = 20.0
    g0 = se.adhesive_ring_geometry(pad, on())
    wider = se.adhesive_ring_geometry(pad, on(adhesive_ring_trim_mm=0.6))
    narrower = se.adhesive_ring_geometry(pad, on(adhesive_ring_trim_mm=-0.6))
    for g in (wider, narrower):
        assert math.isclose(g["outer_r"], g0["outer_r"], abs_tol=EPS), "trim must not move the outer cut"
    assert math.isclose(wider["width"], g0["width"] + 0.6, abs_tol=EPS)
    assert math.isclose(wider["inner_r"], g0["inner_r"] - 0.6, abs_tol=EPS)
    assert math.isclose(narrower["width"], g0["width"] - 0.6, abs_tol=EPS)
    assert math.isclose(narrower["inner_r"], g0["inner_r"] + 0.6, abs_tol=EPS)
    # A trim can't erase the ring.
    gone = se.adhesive_ring_geometry(pad, on(adhesive_ring_trim_mm=-50.0))
    assert math.isclose(gone["width"], se.ADHESIVE_RING_MIN_WIDTH_MM, abs_tol=EPS)
    assert gone["inner_r"] < gone["outer_r"]


def _matt_mode():
    import json
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    with open(os.path.join(repo, "presets", "Matt mode.json"), encoding="utf-8") as f:
        raw = json.load(f)
    preset = raw if any(k in raw for k in config.SIZING_PRESET_KEYS) else raw[next(iter(raw))]
    s = on()
    s.update(config.normalize_sizing_preset(preset))
    s["adhesive_ring_enabled"] = True
    return s


def test_pads_of_9_5_and_under_are_solid_discs():
    # Matt, 2026-11-05: "9.5 and under solid is good". A rule on the pad
    # size, whatever the sizing rules make of the hole.
    assert se.ADHESIVE_RING_SOLID_MAX_PAD_MM == 9.5
    for label, s in (("program defaults", on()), ("Matt mode", _matt_mode()),
                     ("a wide card and thin wrap", on(card_to_felt_offset=0.0, leather_wrap_multiplier=0.5,
                                                      darts_enabled=False))):
        for pad in (7.0, 7.5, 8.0, 8.5, 9.0, 9.5):
            g = se.adhesive_ring_geometry(pad, s)
            assert g["inner_r"] == 0.0, f"{label}: {pad} should be a solid disc, hole r={g['inner_r']}"
            assert g["placement"] == "outside" and g["strokes"], f"{label}: {pad} keeps its number, under the disc"
        for pad in (10.0, 10.5, 12.0, 20.0):
            g = se.adhesive_ring_geometry(pad, s)
            assert g["inner_r"] > 0, f"{label}: {pad} should be a ring"
    # The rule is doing the work: without it these sizes would have had real holes.
    matt = _matt_mode()
    for pad in (7.0, 9.5):
        sizing = config.get_sizing_for_size(pad, matt)
        card_d = pad - (sizing["felt_offset"] + sizing["card_to_felt_offset"])
        hole = card_d - 2 * expected_wrap(pad, matt)
        assert hole > 3.0, f"fixture: Matt mode {pad} would have a {hole:.2f} mm hole"
    svg, g = render(matt, [(9.5, 30.0, 30.0, se.get_disc_diameter(9.5, RING, matt) / 2),
                           (10.0, 60.0, 30.0, se.get_disc_diameter(10.0, RING, matt) / 2)])
    _polys, circles = svg_parts(svg)
    assert len([c for c in circles if math.isclose(c[0], 30.0)]) == 1, "9.5: one circle, no inner cut"
    assert len([c for c in circles if math.isclose(c[0], 60.0)]) == 2, "10.0: outer and inner cut"
    secs = dict(gcode_sections(g))
    assert len(secs["inner"]) == 1 and len(secs["outer"]) == 2, (len(secs["inner"]), len(secs["outer"]))


def test_solid_disc_when_the_hole_would_be_under_2mm():
    assert se.ADHESIVE_RING_MIN_HOLE_MM == 2.0
    s = on()
    # Above 9.5 the hole decides. Walk one size across the threshold with the trim.
    pad = 10.0
    g = se.adhesive_ring_geometry(pad, s)
    hole = 2 * g["inner_r"]
    assert hole > 2.0, f"fixture: {pad} should have a hole, got {hole}"
    just_over = se.adhesive_ring_geometry(pad, on(adhesive_ring_trim_mm=(hole - 2.05) / 2))
    just_under = se.adhesive_ring_geometry(pad, on(adhesive_ring_trim_mm=(hole - 1.95) / 2))
    assert math.isclose(2 * just_over["inner_r"], 2.05, abs_tol=1e-6), just_over["inner_r"]
    assert just_under["inner_r"] == 0.0, f"a 1.95 mm hole must be dropped, got r={just_under['inner_r']}"


# ---------------------------------------------------------- the number ----

def _number_points(g):
    return [p for stroke in g["strokes"] for p in stroke]


def test_every_ring_gets_a_number_with_the_pad_size_text():
    s = on()
    size = 7.0
    while size <= 60.0:
        g = se.adhesive_ring_geometry(size, s)
        assert g["strokes"] and all(len(st) >= 2 for st in g["strokes"]), f"{size}: no number"
        assert g["text"] == f"{size:.1f}".rstrip('0').rstrip('.'), (size, g["text"])
        assert g["placement"] in ("inside", "outside")
        size += 0.5


def test_inside_number_is_in_the_hole_with_clearance():
    s = on()
    for pad in INSIDE_TIGHT:
        tight = se.adhesive_ring_geometry(pad, s)
        assert tight["font"] < se.ADHESIVE_RING_MAX_FONT_MM - 0.5, f"fixture: {pad} should be clearance-limited"
    for pad in INSIDE + INSIDE_TIGHT:
        g = se.adhesive_ring_geometry(pad, s)
        assert g["placement"] == "inside", f"{pad}: {g['placement']}"
        assert g["inner_r"] > 0
        limit = g["inner_r"] - se.ADHESIVE_RING_HOLE_CLEARANCE_MM
        reach = max(math.hypot(x, y) for x, y in _number_points(g))
        assert reach <= limit + EPS, f"{pad}: number reaches {reach}, limit {limit}"
        assert se.ADHESIVE_RING_LEGIBLE_FONT_MM - EPS <= g["font"] <= se.ADHESIVE_RING_MAX_FONT_MM + EPS
        # As large as fits, up to the cap.
        assert (math.isclose(g["font"], se.ADHESIVE_RING_MAX_FONT_MM, abs_tol=EPS)
                or math.isclose(reach, limit, abs_tol=1e-6)), f"{pad}: font {g['font']} not maximal"
        # Centered on the hole.
        x0, y0, x1, y1 = g["box"]
        assert math.isclose(x0 + x1, 0, abs_tol=EPS) and math.isclose(y0 + y1, 0, abs_tol=EPS), g["box"]
        assert math.isclose(g["footprint_r"], g["outer_r"], abs_tol=EPS)


def test_outside_number_sits_under_small_rings_and_solid_discs():
    s = on()
    assert se.adhesive_ring_geometry(7.0, s)["inner_r"] == 0.0     # solid disc
    assert se.adhesive_ring_geometry(10.0, s)["inner_r"] > 0       # ring, small hole
    for pad in OUTSIDE:
        g = se.adhesive_ring_geometry(pad, s)
        assert g["placement"] == "outside", f"{pad}: {g['placement']}"
        assert math.isclose(g["font"], se.ADHESIVE_RING_LEGIBLE_FONT_MM, abs_tol=EPS)
        pts = _number_points(g)
        ys = [y for _x, y in pts]
        xs = [x for x, _y in pts]
        # Y-down: under the ring is +y. The top of the number is the gap below the cut.
        assert math.isclose(min(ys), g["outer_r"] + se.ADHESIVE_RING_OUTSIDE_GAP_MM, abs_tol=EPS), \
            f"{pad}: top of number at {min(ys)}"
        assert math.isclose(min(xs) + max(xs), 0, abs_tol=EPS), f"{pad}: not centered under the ring"
        assert math.isclose(max(ys) - min(ys), 0.9 * g["font"], abs_tol=1e-6), "digits are 0.9 of the font cell tall"
        assert all(math.hypot(x, y) > g["outer_r"] for x, y in pts), "number overlaps the ring"


def test_one_constant_decides_inside_or_outside():
    s = on()
    keep = se.ADHESIVE_RING_LEGIBLE_FONT_MM
    try:
        fits = se.adhesive_ring_geometry(12.0, s)
        assert fits["placement"] == "inside" and fits["font"] < 3.0, fits
        assert se.adhesive_ring_geometry(10.0, s)["placement"] == "outside"
        se.ADHESIVE_RING_LEGIBLE_FONT_MM = 3.0
        g = se.adhesive_ring_geometry(12.0, s)
        assert g["placement"] == "outside" and math.isclose(g["font"], 3.0), g
        se.ADHESIVE_RING_LEGIBLE_FONT_MM = 0.5
        g = se.adhesive_ring_geometry(10.0, s)
        assert g["placement"] == "inside" and g["font"] >= 0.5, g
    finally:
        se.ADHESIVE_RING_LEGIBLE_FONT_MM = keep


def test_number_is_never_filled():
    line = on()
    filled = on(filled_overscan_enabled=True)
    for mat in filled["gcode_settings"]:
        filled["gcode_settings"][mat]["engraving_mode"] = "filled"
    filled["tooling_settings"]["engraving_mode"] = "filled"
    placed, _z, _fp, _ft = nest(line)
    svg_a, g_a = render(line, placed)
    svg_b, g_b = render(filled, placed)
    assert g_a == g_b, "engraving_mode changed the ring G-code"
    assert svg_a == svg_b, "engraving_mode changed the ring SVG"
    assert "<text" not in svg_a
    # The traced number is the stroke font: one polyline per glyph stroke.
    want = sum(len(ge.get_text_strokes(se.zone_label_text(p), 1.0, 0, 0)) for p, *_r in placed)
    polys, _c = svg_parts(svg_a)
    assert len(polys) == want, f"{len(polys)} polylines, stroke font has {want}"
    nums = [st for kind, strokes in gcode_sections(g_a) if kind == "number" for st in strokes]
    assert len(nums) == want, f"{len(nums)} G-code number strokes, stroke font has {want}"


# -------------------------------------------------------------- nesting ----

def test_footprint_is_what_nests():
    s = on()
    for pad in SIZES:
        g = se.adhesive_ring_geometry(pad, s)
        assert math.isclose(se.get_disc_diameter(pad, RING, s), 2 * g["footprint_r"], abs_tol=EPS)
        for bx in (g["box"][0], g["box"][2]):
            for by in (g["box"][1], g["box"][3]):
                if g["placement"] == "outside":
                    assert math.hypot(bx, by) <= g["footprint_r"] + EPS, f"{pad}: number box outside its footprint"
        assert g["footprint_r"] >= g["outer_r"] - EPS
    for pad in OUTSIDE:
        g = se.adhesive_ring_geometry(pad, s)
        assert g["footprint_r"] > g["outer_r"] + 1.0, f"{pad}: an outside number must widen the footprint"
    for pad in INSIDE:
        g = se.adhesive_ring_geometry(pad, s)
        assert math.isclose(g["footprint_r"], g["outer_r"], abs_tol=EPS)
    placed, _z, fp, ft = nest(s)
    assert fp == ft == sum(p['qty'] for p in PADS), (fp, ft)
    for pad, _cx, _cy, r in placed:
        assert math.isclose(r, se.adhesive_ring_geometry(pad, s)["footprint_r"], abs_tol=EPS), \
            "placed radius is the footprint"


def _layout_checks(s, placed, w=None, h=None, polygon=None):
    """Every number inside its own footprint, clear of every other ring's
    cut and number, and on the material; every cut on the material."""
    rings = [(pad, cx, cy, r, se.adhesive_ring_geometry(pad, s)) for pad, cx, cy, r in placed]
    assert any(g["placement"] == "outside" for *_a, g in rings), "fixture has no outside numbers"
    for i, (pad, cx, cy, r, g) in enumerate(rings):
        pts = [p for st in se.adhesive_ring_number_strokes(g, cx, cy) for p in st]
        for x, y in pts:
            assert math.hypot(x - cx, y - cy) <= r + EPS, f"{pad}: number leaves its footprint"
            if polygon:
                assert se._point_in_polygon(x, y, polygon), f"{pad}: number off the scrap at {(x, y)}"
            else:
                assert 0 <= x <= w and 0 <= y <= h, f"{pad}: number off the sheet at {(x, y)}"
        if polygon:
            assert se._circle_fits_in_polygon(cx, cy, g["outer_r"], polygon, spacing_mm=0), f"{pad}: cut off the scrap"
        else:
            assert g["outer_r"] <= cx <= w - g["outer_r"] and g["outer_r"] <= cy <= h - g["outer_r"], \
                f"{pad}: cut off the sheet"
        for j, (pad2, cx2, cy2, r2, g2) in enumerate(rings):
            if i == j:
                continue
            assert math.hypot(cx - cx2, cy - cy2) >= r + r2 + 1.0 - 1e-6, "footprints closer than the nest spacing"
            for x, y in pts:
                gap = math.hypot(x - cx2, y - cy2) - g2["outer_r"]
                assert gap >= 1.0 - 1e-6, f"{pad}'s number is {gap:.2f} mm from a {pad2} ring's cut"


def test_outside_numbers_are_clear_of_neighbours_on_a_rectangle():
    for bias in ("center", "off", "n", "sw", "e"):
        s = on(edge_bias=bias)
        placed, _z, fp, ft = nest(s)
        assert fp == ft, f"bias {bias}: {fp} of {ft} placed"
        _layout_checks(s, placed, w=W, h=H)


def test_outside_numbers_are_clear_of_neighbours_on_a_polygon():
    for bias in ("center", "sw"):
        s = on(edge_bias=bias)
        placed, _z, fp, ft = nest(s, polygon=SCRAP)
        assert fp == ft, f"bias {bias}: {fp} of {ft} placed"
        _layout_checks(s, placed, polygon=SCRAP)


def test_labeled_zones_stand_down_for_rings():
    zoned = on(zone_labels_enabled=True, zone_label_min_size=7.0, zone_label_max_size=12.5)
    placed_z, zones, fp, ft = nest(zoned)
    placed_plain, _z, _fp, _ft = nest(on())
    assert zones == [], "rings must not be grouped into zones"
    assert placed_z == placed_plain and fp == ft
    # ... and the same job on card still groups, so the switch is per material.
    _p, card_zones, _a, _b = se.nest_with_zones(PADS, 'card', 200, 150, zoned)
    assert card_zones, "fixture: card should have produced zones"
    svg, g = render(zoned, placed_z)
    assert "<rect" not in svg and "<text" not in svg


def test_scrap_mode_partial_nest():
    s = on()
    many = [{'size': 12.5, 'qty': 40}, {'size': 20.0, 'qty': 2}]
    placed, remaining, any_placed = se.try_nest_partial(many, RING, 70, 60, s)
    assert any_placed and 2 < len(placed) < 42, len(placed)
    assert len(placed) + sum(p['qty'] for p in remaining) == 42
    _layout_checks(s, placed, w=70, h=60)


def test_oversized_engraving_check_ignores_rings():
    class V:
        def __init__(self, v):
            self.v = v

        def get(self):
            return self.v
    s = on()
    # A font size for the ring can only come from a hand-edited config; even
    # then the ring's number sizes itself, so there is nothing to warn about.
    s["engraving_font_size"] = {m: 50.0 for m in ("felt", "card", "leather", "exact_size", RING)}
    got = se.check_for_oversized_engravings([{'size': 10.0, 'qty': 1}], {'card': V(True), RING: V(True)}, s)
    assert 'card' in got and RING not in got, got


# --------------------------------------------------------------- output ----

def test_svg_parses_with_ring_and_number_on_the_ring_layer():
    s = on()
    placed, _z, _fp, _ft = nest(s)
    svg, _g = render(s, placed)
    root = ET.fromstring(svg)
    assert root.tag.split('}')[-1] == "svg"
    assert root.get("width") == f"{W}mm" and root.get("viewBox") == f"0 0 {W} {H}"
    strokes = svg_strokes(svg)
    assert strokes and all(color == RING_COLOR for _tag, color in strokes), set(strokes)
    assert {tag for tag, _c in strokes} == {"polyline", "circle"}, {tag for tag, _c in strokes}
    polys, circles = svg_parts(svg)
    holes = sum(1 for p, *_r in placed if se.adhesive_ring_geometry(p, s)["inner_r"] > 0)
    assert len(circles) == len(placed) + holes, (len(circles), len(placed), holes)
    assert len(polys) >= len(placed)
    custom = on()
    custom["layer_colors"]["adhesive_ring_cut"] = "#00A000"
    svg2, _g2 = render(custom, placed)
    assert all(c == "#00A000" for _t, c in svg_strokes(svg2)), "ring layer color setting ignored"


def test_cut_radius_is_the_card_radius_even_when_the_footprint_is_larger():
    s = on()
    s["gcode_settings"][RING]["kerf_width"] = 0.0
    placed = [(7.0, 20.0, 20.0, se.get_disc_diameter(7.0, RING, s) / 2),
              (12.5, 50.0, 25.0, se.get_disc_diameter(12.5, RING, s) / 2),
              (20.0, 90.0, 30.0, se.get_disc_diameter(20.0, RING, s) / 2)]
    svg, g = render(s, placed)
    _polys, circles = svg_parts(svg)
    outer_sections = [st for kind, strokes in gcode_sections(g) if kind == "outer" for st in strokes]
    assert len(outer_sections) == 3
    for (pad, cx, cy, foot_r), stroke in zip(placed, outer_sections):
        card_r = se.get_disc_diameter(pad, 'card', s) / 2
        mine = [r for x, y, r in circles if math.isclose(x, cx) and math.isclose(y, cy)]
        assert math.isclose(max(mine), card_r, abs_tol=1e-9), f"{pad}: SVG cut r {max(mine)} != card {card_r}"
        if pad in (7.0, 12.5):
            assert foot_r > card_r + 1.0, "fixture: footprint should be larger than the cut here"
            assert not any(math.isclose(r, foot_r, abs_tol=1e-3) for r in mine), "footprint drawn as a cut"
        for x, y in stroke:
            assert math.isclose(math.hypot(x - cx, y - (H - cy)), card_r, abs_tol=2e-3), \
                f"{pad}: G-code cut radius is not the card radius"


def test_compatibility_mode_svg_is_unitless():
    s = on(compatibility_mode=True)
    placed, _z, _fp, _ft = nest(s)
    svg, _g = render(s, placed)
    body = svg.split("<defs", 1)[1]
    assert 'mm"' not in body, "compatibility mode must not write unit suffixes on ring elements"
    assert "tiny" not in svg
    ET.fromstring(svg)
    plain_svg, _g2 = render(on(), placed)
    assert 'r="' in plain_svg and re.search(r'<circle[^>]*r="[0-9.]+mm"', plain_svg), "default mode writes mm"


def test_svg_and_gcode_draw_the_same_ring_and_number():
    s = on()
    s["gcode_settings"][RING]["kerf_width"] = 0.0     # compare geometry, not kerf
    placed, _z, _fp, _ft = nest(s)
    svg, g = render(s, placed)
    polys, circles = svg_parts(svg)
    sections = gcode_sections(g)
    g_num = [st for kind, strokes in sections if kind == "number" for st in strokes]
    assert len(g_num) == len(polys), (len(g_num), len(polys))
    for sp, gp in zip(polys, g_num):
        assert len(sp) == len(gp), "stroke lengths differ"
        for (sx, sy), (gx, gy) in zip(sp, gp):
            assert abs(sx - gx) < 2e-3 and abs((H - sy) - gy) < 2e-3, f"SVG {(sx, sy)} vs G-code {(gx, gy)}"
    g_circ = [st for kind, strokes in sections if kind in ("inner", "outer") for st in strokes]
    assert len(g_circ) == len(circles), (len(g_circ), len(circles))
    for stroke in g_circ:
        # Each G-code circle is one of the SVG circles, flipped.
        x0 = (min(p[0] for p in stroke) + max(p[0] for p in stroke)) / 2
        y0 = (min(p[1] for p in stroke) + max(p[1] for p in stroke)) / 2
        rr = (max(p[0] for p in stroke) - min(p[0] for p in stroke)) / 2
        assert any(abs(cx - x0) < 2e-3 and abs((H - cy) - y0) < 2e-3 and abs(r - rr) < 2e-3
                   for cx, cy, r in circles), f"G-code circle at {(x0, y0)} r={rr} has no SVG twin"


def test_number_is_upright_and_on_the_same_side_in_both_outputs():
    # "7" in the stroke font is one stroke: top bar left to right, then the leg.
    assert ge.STROKE_FONT['7'] == [[(0.1, 0.95), (0.5, 0.95), (0.25, 0.05)]], "fixture: the 7 glyph changed"
    s = on()
    s["gcode_settings"][RING]["kerf_width"] = 0.0
    pad, cx, cy = 7.0, 30.0, 25.0
    g = se.adhesive_ring_geometry(pad, s)
    assert g["placement"] == "outside" and g["text"] == "7"
    svg, gc = render(s, [(pad, cx, cy, g["footprint_r"])])
    polys, _c = svg_parts(svg)
    assert len(polys) == 1
    (ax, ay), (bx, by), (lx, ly) = polys[0]
    assert math.isclose(ay, by, abs_tol=1e-6) and ax < bx, "SVG: top bar should run left to right"
    assert ly > ay, "SVG (Y-down): the leg goes down the page, so the 7 is upright"
    assert min(ay, ly) > cy + g["outer_r"], "SVG: the number is under the ring"
    nums = [st for kind, strokes in gcode_sections(gc) if kind == "number" for st in strokes]
    assert len(nums) == 1
    (ax, ay), (bx, by), (lx, ly) = nums[0]
    gcy = H - cy
    assert math.isclose(ay, by, abs_tol=1e-6) and ax < bx, "G-code: top bar should run left to right (not mirrored)"
    assert ly < ay, "G-code (Y-up): the leg goes toward the machine front, so the 7 is upright"
    assert max(ay, ly) < gcy - g["outer_r"], "G-code: the number is on the same side as in the SVG (under)"


def test_gcode_goes_number_inner_outer_on_the_sheet_and_ends_laser_off():
    s = on()
    placed, _z, _fp, _ft = nest(s)
    holes = sum(1 for p, *_r in placed if se.adhesive_ring_geometry(p, s)["inner_r"] > 0)
    _svg, g = render(s, placed)
    kinds = [k for k, _st in gcode_sections(g)]
    assert kinds == ["number", "inner", "outer"], f"layer grouping order: {kinds}"
    secs = dict(gcode_sections(g))
    assert len(secs["inner"]) == holes and len(secs["outer"]) == len(placed)
    lines = g.splitlines()
    assert lines[-2:] == ["M5", "M2"], lines[-3:]
    assert "M4" in lines and "M3" not in lines
    for kind, strokes in gcode_sections(g):
        for st in strokes:
            for x, y in st:
                assert 0 <= x <= W and 0 <= y <= H, f"{kind} move off the sheet: {(x, y)}"
    assert "; Layer C07" in g
    assert not re.search(r"; Layer C(?!07)", g), "everything belongs to the ring's one layer"

    by_pad = on(gcode_cut_grouping="pad")
    _svg2, g2 = render(by_pad, placed)
    kinds2 = [k for k, _st in gcode_sections(g2)]
    want = []
    for p, *_r in placed:
        want += ["number"] + (["inner"] if se.adhesive_ring_geometry(p, by_pad)["inner_r"] > 0 else []) + ["outer"]
    assert kinds2 == want, "per-pad grouping: number, inner, outer for each ring in turn"


def test_gcode_uses_the_ring_cut_settings_and_kerf():
    s = on()
    s["gcode_settings"][RING].update(cut_speed=777, cut_power=33, cut_passes=2, kerf_width=0.4,
                                     air_assist_cut=False)
    s["gcode_settings"]["felt"].update(cut_speed=111, cut_power=99)
    pad, cx, cy = 20.0, 40.0, 40.0
    g = se.adhesive_ring_geometry(pad, s)
    placed = [(pad, cx, cy, g["footprint_r"])]
    _svg, gc = render(s, placed)
    heads = [ln for ln in gc.splitlines() if ln.startswith("; Cut @")]
    assert len(heads) == 3 and all(h == "; Cut @ 777 mm/min, 33% power, x2 passes" for h in heads), heads
    # The number's passes all come first, the way engraving does, before any cut.
    assert [k for k, _st in gcode_sections(gc)] == ["number", "inner", "outer"]
    body_lines = gc.splitlines()
    last_number_pass = max(i for i, ln in enumerate(body_lines) if ln.strip() == "; Pass 2 of 2"
                           and i < body_lines.index("; Ring inner cut"))
    assert body_lines.index("; Ring number") < last_number_pass < body_lines.index("; Ring inner cut")
    assert "S330" in gc and "F777" in gc and "F111" not in gc
    body = gc.split("; Ring number", 1)[1].split("; return to origin")[0]
    assert "M9" in body and "M8" not in body, "air off must reach every ring step"
    secs = dict(gcode_sections(gc))
    assert len(secs["number"]) == 2 * len(g["strokes"]), "the number gets the cut passes too"
    gcy = H - cy
    for x, y in secs["outer"][0]:
        assert math.isclose(math.hypot(x - cx, y - gcy), g["outer_r"] + 0.2, abs_tol=2e-3), "kerf goes outward on the outer cut"
    for x, y in secs["inner"][0]:
        assert math.isclose(math.hypot(x - cx, y - gcy), g["inner_r"] - 0.2, abs_tol=2e-3), "kerf goes inward on the inner cut"
    no_kerf = copy.deepcopy(s)
    no_kerf["gcode_settings"][RING]["kerf_width"] = 0.0
    _svg0, gc0 = render(no_kerf, placed)
    assert dict(gcode_sections(gc0))["number"] == secs["number"], "the number must not be kerf compensated"
    # A settings dict with no ring block falls back to the ring defaults, never to felt.
    bare = copy.deepcopy(s)
    del bare["gcode_settings"][RING]
    _s, gcb = render(bare, placed)
    d = config.DEFAULT_SETTINGS["gcode_settings"][RING]
    assert f"; Cut @ {d['cut_speed']} mm/min, {d['cut_power']}% power" in gcb


def test_inner_cut_is_as_round_as_the_outer_cut():
    # A ring's inner cut is nearly as big as its outer cut, so it gets the
    # same number of segments. At the center-hole figure (36) a 51 mm pad's
    # inner cut had 3.7 mm flats, 0.08 mm off a true circle (2026-11-05).
    assert ge.ADHESIVE_RING_CUT_SEGMENTS == 72
    s = on()
    s["gcode_settings"][RING]["kerf_width"] = 0.0
    for pad in (10.0, 20.0, 35.5, 51.0, 60.0):
        g = se.adhesive_ring_geometry(pad, s)
        assert g["inner_r"] > 0, f"fixture: {pad} should have a hole"
        _svg, gc = render(s, [(pad, 45.0, 45.0, g["footprint_r"])])
        secs = dict(gcode_sections(gc))
        inner, outer = secs["inner"][0], secs["outer"][0]
        assert len(inner) == len(outer) == 73, f"{pad}: inner {len(inner) - 1} segments, outer {len(outer) - 1}"
        assert inner[0] == inner[-1] and outer[0] == outer[-1], "both cuts close on themselves"
        for stroke, radius in ((inner, g["inner_r"]), (outer, g["outer_r"])):
            flats = [math.hypot(x1 - x0, y1 - y0) for (x0, y0), (x1, y1) in zip(stroke, stroke[1:])]
            # chord of a 5 degree step, and how far its middle sits inside the true circle
            assert max(flats) <= 2 * radius * math.sin(math.pi / 72) + 2e-3, f"{pad}: {max(flats):.2f} mm flats"
            sag = radius * (1 - math.cos(math.pi / 72))
            assert sag < 0.03, f"{pad}: cut deviates {sag:.3f} mm from a circle"
    # The 51 mm pad that showed it: inner flats are now shorter than the outer cut's.
    g = se.adhesive_ring_geometry(51.0, s)
    _svg, gc = render(s, [(51.0, 45.0, 45.0, g["footprint_r"])])
    secs = dict(gcode_sections(gc))
    longest = [max(math.hypot(x1 - x0, y1 - y0) for (x0, y0), (x1, y1) in zip(st, st[1:]))
               for st in (secs["inner"][0], secs["outer"][0])]
    assert longest[0] < longest[1] < 2.3, longest


def test_flip_height_override_is_honoured():
    s = on()
    s["gcode_settings"][RING]["kerf_width"] = 0.0
    pad, cx, cy = 20.0, 45.0, 37.0
    g = se.adhesive_ring_geometry(pad, s)
    placed = [(pad, cx, cy, g["footprint_r"])]
    inset_poly = [(5.0, 5.0), (95.0, 5.0), (95.0, 75.0), (5.0, 75.0)]
    for kw, flip in (({}, H), ({"polygon": inset_poly}, 75.0), ({"polygon": inset_poly, "flip_height_mm": 80.0}, 80.0)):
        _svg, gc = render(s, placed, **kw)
        outer = dict(gcode_sections(gc))["outer"][0]
        y0 = (min(p[1] for p in outer) + max(p[1] for p in outer)) / 2
        assert math.isclose(y0, flip - cy, abs_tol=2e-3), f"{kw}: center y {y0}, wanted {flip - cy}"


def test_feature_off_leaves_other_materials_byte_identical():
    absent = base_settings()
    for k in ("adhesive_ring_enabled", "adhesive_ring_trim_mm"):
        del absent[k]
    del absent["gcode_settings"][RING]
    del absent["layer_colors"]["adhesive_ring_cut"]
    del absent["laser_framing_power_by_material"][RING]
    off = base_settings()
    enabled = on(adhesive_ring_trim_mm=1.5)
    enabled["gcode_settings"][RING].update(cut_speed=4321, kerf_width=0.9)
    enabled["layer_colors"]["adhesive_ring_cut"] = "#00A000"
    for extra in ({}, {"zone_labels_enabled": True}, {"locator_marks_enabled": True, "gcode_cut_grouping": "pad"}):
        for mat in ("felt", "card", "leather", "exact_size"):
            outs = []
            for s in (absent, off, enabled):
                s2 = copy.deepcopy(s)
                s2.update(extra)
                placed, zones, _fp, _ft = se.nest_with_zones(PADS, mat, 220, 160, s2)
                with tempfile.TemporaryDirectory() as d:
                    sp, gp = os.path.join(d, 'a.svg'), os.path.join(d, 'a.gcode')
                    se.generate_svg_from_placed(placed, mat, 220, 160, sp, 3.5, s2, zones=zones)
                    ge.generate_gcode_from_placed(placed, mat, 220, 160, gp, 3.5, s2, zones=zones)
                    outs.append((placed, open(sp, 'rb').read(), open(gp, 'rb').read()))
            assert outs[0] == outs[1] == outs[2], f"{mat} {extra}: output depends on the adhesive ring settings"
            assert outs[0][1] and outs[0][2]


# -------------------------------------------------------------- presets ----

def test_sizing_preset_captures_and_backfills():
    keys = ("adhesive_ring_enabled", "adhesive_ring_trim_mm")
    for k in keys:
        assert k in config.SIZING_PRESET_KEYS, k
    preset = config.settings_to_sizing_preset(on(adhesive_ring_trim_mm=-0.4))
    assert preset["adhesive_ring_enabled"] is True and preset["adhesive_ring_trim_mm"] == -0.4
    old = {k: v for k, v in preset.items() if k not in keys}
    fixed = config.normalize_sizing_preset(old)
    for k in keys:
        assert fixed[k] == config.DEFAULT_SETTINGS[k], f"{k} not backfilled"
    assert fixed == config.normalize_sizing_preset(config.settings_to_sizing_preset(base_settings()))


def test_gcode_preset_material_and_default_backfill():
    assert config.gcode_preset_keys_for(RING) == config.GCODE_CUT_ONLY_KEYS
    assert config.gcode_preset_keys_for("felt") == config.GCODE_PRESET_KEYS
    s = base_settings()
    boot = config.settings_to_gcode_presets(s)
    assert set(boot[RING]["Default"]) == set(config.GCODE_CUT_ONLY_KEYS)
    # An existing user's library, written before the material existed (main.py's loop).
    library = {m: {"Mine": config.material_settings_to_gcode_preset(s["gcode_settings"][m])}
               for m in ("felt", "card", "leather", "acrylic", "basswood")}
    before = copy.deepcopy(library)
    for mat in config.GCODE_PRESET_MATERIALS:
        if mat not in library or not library[mat]:
            library[mat] = boot.get(mat, {})
    assert library[RING] == {"Default": s["gcode_settings"][RING]}
    assert {m: library[m] for m in before} == before, "backfill touched an existing material"


def test_old_settings_file_loads_with_ring_defaults():
    old = base_settings()
    for k in ("adhesive_ring_enabled", "adhesive_ring_trim_mm"):
        del old[k]
    del old["gcode_settings"][RING]
    del old["layer_colors"]["adhesive_ring_cut"]
    del old["laser_framing_power_by_material"][RING]
    old["gcode_settings"]["felt"]["cut_speed"] = 613
    keep = config.SETTINGS_FILE
    path = os.path.join(_PROFILE, "old_settings.json")
    with open(path, "w") as f:
        json.dump(old, f)
    try:
        config.SETTINGS_FILE = path
        got = config.load_settings()
    finally:
        config.SETTINGS_FILE = keep
    d = config.DEFAULT_SETTINGS
    assert got["adhesive_ring_enabled"] is False and got["adhesive_ring_trim_mm"] == 0.0
    assert got["gcode_settings"][RING] == d["gcode_settings"][RING]
    assert got["gcode_settings"]["felt"]["cut_speed"] == 613
    assert got["layer_colors"]["adhesive_ring_cut"] == d["layer_colors"]["adhesive_ring_cut"]
    assert got["laser_framing_power_by_material"][RING] == 10


def test_framing_power_has_its_own_clamped_entry():
    s = base_settings()
    assert config.get_framing_power_s(RING, s) == 10
    s["laser_framing_power_by_material"][RING] = 23
    assert config.get_framing_power_s(RING, s) == 23
    assert config.get_framing_power_s("felt", s) == 10
    s["laser_framing_power_by_material"][RING] = 500
    assert config.get_framing_power_s(RING, s) == config.FRAMING_POWER_S_MAX
    lines = ge.generate_framing_gcode(0, 0, 50, 40, power_s=config.get_framing_power_s(RING, s))
    assert any(ln.startswith("M4 S30") for ln in lines) and not any(ln.startswith("M3") for ln in lines)


# ------------------------------------------------------------------ GUI ----

def gui_tests(root):
    import tkinter as tk
    import main as main_mod
    import ui_dialogs
    import falcon_sender
    from ui_dialogs import (OptionsWindow, LayerColorWindow, GcodeSettingsWindow,
                            PadPreviewWindow, NestingPreviewWindow, JobHistoryWindow)

    app = main_mod.PadSVGGeneratorApp(root)
    out_dir = tempfile.mkdtemp(prefix="ssc-adhesive-ring-out-")
    boxes = []
    main_mod.filedialog.askdirectory = lambda **k: out_dir
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(main_mod.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))

    def set_feature(enabled):
        app.settings["adhesive_ring_enabled"] = enabled
        app.update_ui_from_settings()
        root.update_idletasks()

    def options(settings, presets=None):
        class StubApp:
            def open_resonance_window(self):
                pass
        if presets is None:
            presets = {"Default": config.settings_to_sizing_preset(settings)}
        w = OptionsWindow(root, StubApp(), settings, lambda: None, lambda: None,
                          sizing_presets=presets, sizing_presets_save_callback=lambda: None)
        w.top.withdraw()
        return w

    # ---- Sizing Rules --------------------------------------------------
    def enabling_control_is_marked_beta_and_lives_in_lesser_used():
        w = options(base_settings())
        try:
            assert "beta" in w.ring_frame.cget("text").lower(), w.ring_frame.cget("text")
            assert w.ring_frame.master is w.lesser_body, "the section belongs in Lesser-used settings"
            assert w.ring_cb.master is w.ring_frame and w.ring_clearance_spin.winfo_toplevel() is w.top
            assert not w.lesser_open, "collapsed while nothing inside is active"
            assert str(w.ring_clearance_spin.cget("state")) == "disabled", "counter greyed while the feature is off"
            w.ring_enabled_var.set(True)
            w._toggle_ring_fields()
            assert str(w.ring_clearance_spin.cget("state")) == "normal"
        finally:
            w.top.destroy()
        w2 = options(on())
        try:
            assert w2.lesser_open, "opens itself when adhesive rings are on"
        finally:
            w2.top.destroy()
    check("GUI: the enabling control is marked beta, in Lesser-used settings, auto-open when on",
          enabling_control_is_marked_beta_and_lives_in_lesser_used)

    def form_round_trip_and_preset_detection():
        w = options(base_settings())
        try:
            w.ring_enabled_var.set(True)
            w.ring_trim_var.set(-0.3)
            snap = w._capture_form_to_dict()
            assert snap["adhesive_ring_enabled"] is True and snap["adhesive_ring_trim_mm"] == -0.3
            assert w._is_dirty(), "ticking the feature must count as an edit"
            w2 = options(base_settings())
            try:
                w2._apply_dict_to_form(snap)
                assert w2._capture_form_to_dict() == snap
            finally:
                w2.top.destroy()
        finally:
            w.top.destroy()
        # A preset saved before the feature existed still names itself.
        s = base_settings()
        old = {k: v for k, v in config.settings_to_sizing_preset(s).items()
               if not k.startswith("adhesive_ring_")}
        w3 = options(s, presets={"Old shop preset": old})
        try:
            assert w3._detect_active_preset() == "Old shop preset"
        finally:
            w3.top.destroy()
        # ... and a preset with rings on is told apart from it.
        s_on = on(adhesive_ring_trim_mm=0.25)
        w4 = options(s_on, presets={"Old shop preset": old, "Rings": config.settings_to_sizing_preset(s_on)})
        try:
            assert w4._detect_active_preset() == "Rings"
            s_applied = base_settings()
            w4.settings = s_applied
            w4.save_options()
            assert s_applied["adhesive_ring_enabled"] is True and s_applied["adhesive_ring_trim_mm"] == 0.25
        finally:
            try:
                w4.top.destroy()
            except tk.TclError:
                pass
    check("GUI: form round trip, Apply writes both keys, old presets still detected",
          form_round_trip_and_preset_detection)

    # ---- Wrap clearance: the stored trim, shown as a counter ----------------
    # Matt, 2026-11-05: "wrap clearance. then the tooltip says what it does";
    # "negative clearance as being inside and hidden, positive clearance as
    # being extra ring material". Same number and sign as the stored
    # adhesive_ring_trim_mm; only the control and its name changed.

    def type_clearance(w, text):
        w.ring_clearance_spin.delete(0, tk.END)
        w.ring_clearance_spin.insert(0, text)

    def wrap_clearance_is_a_spinbox_with_a_tooltip():
        w = options(on())
        try:
            spin = w.ring_clearance_spin
            assert isinstance(spin, tk.Spinbox), type(spin)
            assert not hasattr(w, "ring_trim_entry"), "the old entry is still there"
            assert w.ring_clearance_label.cget("text") == "Wrap clearance:", w.ring_clearance_label.cget("text")
            assert w.ring_clearance_unit.cget("text") == "mm"
            assert (float(spin.cget("from")), float(spin.cget("to")), float(spin.cget("increment"))) == \
                (-5.0, 5.0, 0.1), (spin.cget("from"), spin.cget("to"), spin.cget("increment"))
            assert spin.cget("format") == "%.1f"
            assert spin.master is w.ring_clearance_label.master is w.ring_clearance_unit.master
            assert str(spin.cget("state")) == "normal" and spin.get() == "0.0", spin.get()
            # The tooltip is on the label, the counter and the unit, and says what the sign does.
            tips = w.ring_clearance_tips
            assert [t.widget for t in tips] == [w.ring_clearance_label, spin, w.ring_clearance_unit]
            for wid in (w.ring_clearance_label, spin, w.ring_clearance_unit):
                assert wid.bind("<Enter>"), "no tooltip on part of the row"
            text = tips[0].text
            assert all(t.text == text for t in tips)
            for phrase in ("past the edge of the wrap", "Positive", "extra ring", "Negative", "hidden under the wrap"):
                assert phrase in text, f"tooltip does not say {phrase!r}: {text}"
            # The arrows: 0.1 at a time, exactly, one decimal shown, stopping at both ends.
            w.ring_trim_var.set(-5.0)
            seen, shown = [], []
            for _i in range(101):
                seen.append(w.ring_trim_var.get())
                shown.append(spin.get())
                spin.invoke("buttonup")
            assert seen == [round(-5.0 + 0.1 * i, 1) for i in range(101)], seen[:5]
            assert shown == [f"{v:.1f}" for v in seen], shown[:5]
            assert spin.get() == "5.0", "the up arrow must stop at 5.0"
            w.ring_trim_var.set(-5.0)
            spin.invoke("buttondown")
            assert spin.get() == "-5.0", "the down arrow must stop at -5.0"
            # Off: greyed, like before.
            w.ring_enabled_var.set(False)
            w._toggle_ring_fields()
            assert str(spin.cget("state")) == "disabled"
        finally:
            w.top.destroy()
        # The names that were tried and dropped are gone from what the user reads.
        repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        src = open(os.path.join(repo, "ui_dialogs.py"), encoding="utf-8").read()
        pot = open(os.path.join(repo, "locale", "saxshop.pot"), encoding="utf-8").read()
        for old_name in ("Ring width trim", "Inner edge clearance", "inner diameter adjustment"):
            assert old_name not in src and old_name not in pot, f"{old_name!r} is still user-facing text"
        assert 'msgid "Wrap clearance:"' in pot
    check("GUI: Wrap clearance is a spinbox, -5.0 to 5.0 in 0.1 steps, with the tooltip on label, counter and unit",
          wrap_clearance_is_a_spinbox_with_a_tooltip)

    def wrap_clearance_plus_and_minus_one_in_the_written_svg():
        # Typed into the counter, applied, generated, and read back from the file.
        pad = 20.0

        def cut_diameters(text):
            s = on()
            w = options(s)
            type_clearance(w, text)
            w.save_options()                       # Apply's last step; closes the dialog
            assert s["adhesive_ring_trim_mm"] == float(text), s["adhesive_ring_trim_mm"]
            g = se.adhesive_ring_geometry(pad, s)
            with tempfile.TemporaryDirectory() as d:
                path = os.path.join(d, "r.svg")
                se.generate_svg_from_placed([(pad, 40.0, 40.0, g["footprint_r"])], RING, 80, 80, path, 3.5, s)
                _polys, circles = svg_parts(open(path, encoding="utf-8").read())
            radii = sorted(r for _x, _y, r in circles)
            assert len(radii) == 2, radii
            return 2 * radii[0], 2 * radii[1]
        zero_in, zero_out = cut_diameters("0.0")
        plus_in, plus_out = cut_diameters("1.0")
        minus_in, minus_out = cut_diameters("-1.0")
        assert zero_out == plus_out == minus_out, "the outer cut must not move"
        assert math.isclose(zero_in - plus_in, 2.0, abs_tol=1e-9), \
            f"+1.0 is extra ring: the hole should be 2.0 mm smaller, got {zero_in - plus_in:.4f}"
        assert math.isclose(minus_in - zero_in, 2.0, abs_tol=1e-9), \
            f"-1.0 pulls the edge back: the hole should be 2.0 mm larger, got {minus_in - zero_in:.4f}"
        # Nothing else is shaping this size: not the 9.5 rule, the 2 mm cut-off or the 0.5 mm floor.
        assert pad > se.ADHESIVE_RING_SOLID_MAX_PAD_MM and plus_in > se.ADHESIVE_RING_MIN_HOLE_MM + 1
        assert (zero_out - minus_in) / 2 > se.ADHESIVE_RING_MIN_WIDTH_MM + 0.1
    check("GUI: Wrap clearance +1.0 makes the inner cut 2.0 mm smaller in the written SVG, -1.0 makes it 2.0 mm larger",
          wrap_clearance_plus_and_minus_one_in_the_written_svg)

    def every_tenth_survives_apply_save_and_reopen():
        # -5.0 ... 5.0: typed, applied (which saves it as a preset first), written
        # to disk, loaded back, and shown again by a newly opened dialog.
        class NamePrompt:
            def __init__(self, parent, existing_names, default_existing=None, title=None, intro=None):
                self.result = {"name": "Rings"}

        class StubApp:
            def open_resonance_window(self):
                pass

        def open_on(s, presets):
            w = OptionsWindow(root, StubApp(), s, lambda: None, lambda: config.save_settings(s),
                              sizing_presets=presets,
                              sizing_presets_save_callback=lambda: config.save_presets(presets, config.SIZING_PRESET_FILE))
            w.top.withdraw()
            return w

        keep = ui_dialogs.SaveSizingPresetDialog
        ui_dialogs.SaveSizingPresetDialog = NamePrompt
        w = open_on(on(), {"Default": config.settings_to_sizing_preset(base_settings())})
        try:
            for i in range(101):
                v = round(-5.0 + 0.1 * i, 1)
                type_clearance(w, f"{v:.1f}")
                box_guard.GUARD.expect("Save changes as a preset?", answer=True)
                w.on_apply()                       # unsaved edit -> save as "Rings" -> apply -> close
                assert not w.top.winfo_exists(), f"{v}: Apply did not go through"
                s = config.load_settings()
                presets = config.load_presets(config.SIZING_PRESET_FILE)
                assert s["adhesive_ring_trim_mm"] == v, f"{v}: settings file has {s['adhesive_ring_trim_mm']!r}"
                assert presets["Rings"]["adhesive_ring_trim_mm"] == v, f"{v}: preset file"
                w = open_on(s, presets)
                assert w.ring_clearance_spin.get() == f"{v:.1f}", f"{v}: reopened dialog shows {w.ring_clearance_spin.get()!r}"
                assert w.ring_trim_var.get() == v and w._capture_form_to_dict()["adhesive_ring_trim_mm"] == v
                assert w.active_preset_name == "Rings" and not w._is_dirty(), f"{v}: preset not detected on reopen"
        finally:
            ui_dialogs.SaveSizingPresetDialog = keep
            if w.top.winfo_exists():
                w.top.destroy()
    check("GUI: every Wrap clearance from -5.0 to 5.0 in 0.1 steps survives Apply, save and reopen unchanged",
          every_tenth_survives_apply_save_and_reopen)

    def stored_finer_value_is_not_rounded_by_opening_the_dialog():
        # Tk reformats a variable when a spinbox is created on it; a value saved
        # with two decimals (typed, or from the entry this control replaced)
        # must come back as it was saved, or its preset stops naming itself.
        s = on(adhesive_ring_trim_mm=0.25)
        w = options(s, presets={"Fine": config.settings_to_sizing_preset(s)})
        try:
            assert w.ring_clearance_spin.get() == "0.25", w.ring_clearance_spin.get()
            assert w._capture_form_to_dict()["adhesive_ring_trim_mm"] == 0.25
            assert w._detect_active_preset() == "Fine" and not w._is_dirty()
            w._apply_dict_to_form(config.normalize_sizing_preset(dict(config.settings_to_sizing_preset(s),
                                                                      adhesive_ring_trim_mm=-1.35)))
            assert w.ring_clearance_spin.get() == "-1.35" and w.ring_trim_var.get() == -1.35
            box_guard.GUARD.expect("Revert to Program Defaults", answer=True)
            w.revert_to_defaults()                 # askyesno is answered Yes in this suite
            assert w.ring_clearance_spin.get() == "0.0" and w.ring_enabled_var.get() is False
        finally:
            w.top.destroy()
    check("GUI: a stored clearance with two decimals is shown and kept as saved (preset still detected)",
          stored_finer_value_is_not_rounded_by_opening_the_dialog)

    def typed_out_of_range_or_not_a_number_does_not_crash_apply():
        # Out of range: held to the range, by Apply and by leaving the field.
        for typed, stored in (("7.3", 5.0), ("-12", -5.0), ("5.0", 5.0), ("-5", -5.0), ("0.25", 0.25), (" 1.5 ", 1.5)):
            s = on()
            w = options(s)
            type_clearance(w, typed)
            assert w._capture_form_to_dict()["adhesive_ring_trim_mm"] == stored, (typed, stored)
            w.save_options()
            assert s["adhesive_ring_trim_mm"] == stored, (typed, s["adhesive_ring_trim_mm"])
        w = options(on())
        try:
            type_clearance(w, "7.3")
            w._clamp_ring_clearance()              # what leaving the field does
            assert w.ring_clearance_spin.get() == "5.0", w.ring_clearance_spin.get()
            type_clearance(w, "1.25")
            w._clamp_ring_clearance()
            assert w.ring_clearance_spin.get() == "1.25", "a value in range is left as typed"
        finally:
            w.top.destroy()
        # Not a number: Apply says so, stores nothing, and the dialog stays open.
        for typed in ("abc", "", "1,5", "--", "0.1.2"):
            s = on(adhesive_ring_trim_mm=0.4)
            w = options(s)
            try:
                type_clearance(w, typed)
                boxes.clear()
                w.on_apply()
                assert w.top.winfo_exists(), f"{typed!r}: Apply closed the dialog"
                assert s["adhesive_ring_trim_mm"] == 0.4, f"{typed!r}: something was stored"
                assert [b[1] for b in boxes] == ["Invalid value"], f"{typed!r}: {boxes}"
                w._clamp_ring_clearance()          # leaving the field must not raise either
                assert w._is_dirty()
                type_clearance(w, "0.4")           # fixed: clean again
                assert not w._is_dirty()
            finally:
                if w.top.winfo_exists():
                    w.top.destroy()
    check("GUI: a typed Wrap clearance out of range is held to the range; not a number is refused without a crash",
          typed_out_of_range_or_not_a_number_does_not_crash_apply)

    # ---- appears only when enabled -------------------------------------
    def material_checkbox_only_when_enabled():
        set_feature(False)
        assert RING not in app.material_vars and RING not in app.material_checkboxes
        set_feature(True)
        assert RING in app.material_vars and app.material_vars[RING].get() is False
        cb = app.material_checkboxes[RING]
        assert cb.cget("text") == "Adhesive rings" and cb.winfo_manager() == "grid"
        assert cb.master is app.material_checkboxes["felt"].master, "beside the other materials"
        app.material_vars[RING].set(True)
        set_feature(False)
        assert RING not in app.material_vars and RING not in app.material_checkboxes
        assert not cb.winfo_exists()
        assert [m for m, v in app.material_vars.items() if v.get() and m == RING] == []
        set_feature(True)
        assert app.material_vars[RING].get() is False, "comes back unticked"
    check("GUI: the Adhesive rings material checkbox exists only while enabled", material_checkbox_only_when_enabled)

    def layer_colors_row_only_when_enabled():
        for enabled in (False, True):
            s = base_settings(adhesive_ring_enabled=enabled)
            w = LayerColorWindow(root, s, lambda: None)
            w.top.withdraw()
            try:
                assert ("adhesive_ring_cut" in w.color_vars) == enabled, f"enabled={enabled}"
                if enabled:
                    assert w.color_vars["adhesive_ring_cut"].get().startswith("07")
                    w.color_vars["adhesive_ring_cut"].set("11 - Dark Green")
                    w.save_colors()
                    assert s["layer_colors"]["adhesive_ring_cut"] == "#00A000"
                else:
                    w.save_colors()
                    assert s["layer_colors"]["adhesive_ring_cut"] == RING_COLOR, "hidden row must keep its color"
            finally:
                try:
                    w.top.destroy()
                except tk.TclError:
                    pass
    check("GUI: Layer Colors lists the ring layer only while enabled", layer_colors_row_only_when_enabled)

    def gcode_section_only_when_enabled():
        seen = []
        real = main_mod.GcodeSettingsWindow

        def recorder(parent, settings, save_cb, **kw):
            seen.append([k for k, _label in kw.get("materials") or []])
        main_mod.GcodeSettingsWindow = recorder
        try:
            set_feature(False)
            app.open_gcode_settings_window()
            set_feature(True)
            app.open_gcode_settings_window()
        finally:
            main_mod.GcodeSettingsWindow = real
        assert seen[0] == ["felt", "card", "leather"], seen[0]
        assert seen[1] == ["felt", "card", "leather", RING], seen[1]
    check("GUI: G-code Settings gets a ring section only while enabled", gcode_section_only_when_enabled)

    def gcode_section_is_cut_only_with_presets():
        s = on()
        presets = config.settings_to_gcode_presets(s)
        saved = []
        w = GcodeSettingsWindow(root, s, lambda _s: saved.append(copy.deepcopy(_s["gcode_settings"])),
                                materials=[("felt", "Felt"), (RING, "Adhesive rings")],
                                gcode_presets=presets, gcode_presets_save_callback=lambda: None)
        w.top.withdraw()
        try:
            assert set(w.vars[RING]) == {"cut", "kerf_width", "air_assist_cut"}, set(w.vars[RING])
            assert "engraving" in w.vars["felt"] and "hole" in w.vars["felt"]
            assert w.active_preset_name[RING] == "Default", "the bootstrapped Default is detected"
            assert not w._dirty_materials()
            snap = w._capture_material_to_dict(RING)
            assert set(snap) == set(config.GCODE_CUT_ONLY_KEYS), snap
            w.vars[RING]["cut"]["speed"].set(2222)
            w.vars[RING]["kerf_width"].set(0.25)
            assert w._dirty_materials() == [RING], "editing tape must not dirty felt"
            edited = w._capture_material_to_dict(RING)
            w._apply_dict_to_material(RING, presets[RING]["Default"])
            assert w._capture_material_to_dict(RING) == snap
            w._apply_dict_to_material(RING, edited)
            w._reset_defaults()
            assert w._capture_material_to_dict(RING) == config.DEFAULT_SETTINGS["gcode_settings"][RING]
            w._apply_dict_to_material(RING, edited)
            w._on_save()
            assert saved, "save callback not called"
            assert saved[-1][RING] == edited and set(saved[-1][RING]) == set(config.GCODE_CUT_ONLY_KEYS)
            assert set(saved[-1]["felt"]) == set(config.GCODE_PRESET_KEYS), "felt keeps its full block"
        finally:
            try:
                w.top.destroy()
            except tk.TclError:
                pass
    check("GUI: the ring's G-code section has cut, kerf and air only, and its presets work",
          gcode_section_is_cut_only_with_presets)

    def framing_power_row_only_when_enabled():
        app.settings["laser_framing_power_by_material"][RING] = 17
        for enabled in (False, True):
            set_feature(enabled)
            before = set(root.winfo_children())
            app._on_machine_framing_power()
            dlgs = [c for c in root.winfo_children() if c not in before and hasattr(c, "vars_by_mat")]
            assert len(dlgs) == 1, "framing power dialog not found"
            dlg = dlgs[0]
            try:
                assert (RING in dlg.vars_by_mat) == enabled, f"enabled={enabled}: {list(dlg.vars_by_mat)}"
                assert "felt" in dlg.vars_by_mat
                if enabled:
                    assert dlg.vars_by_mat[RING].get() == "1.7"
                    dlg.vars_by_mat[RING].set("2.3")
                ok = [b for fr in dlg.winfo_children() for b in fr.winfo_children()
                      if isinstance(b, tk.Button) and b.cget("text") == "OK"]
                ok[0].invoke()
            finally:
                try:
                    dlg.destroy()
                except tk.TclError:
                    pass
            want = 23 if enabled else 17
            assert app.settings["laser_framing_power_by_material"][RING] == want, \
                f"enabled={enabled}: stored {app.settings['laser_framing_power_by_material'].get(RING)}"
    check("GUI: Framing Power has a ring row only while enabled, and keeps the hidden value",
          framing_power_row_only_when_enabled)

    # ---- previews -------------------------------------------------------
    def pad_preview_shows_the_ring_and_number():
        for enabled in (False, True):
            w = options(base_settings(adhesive_ring_enabled=enabled))
            pv = PadPreviewWindow(w)
            pv.withdraw()
            pv._cancel_poll()
            try:
                for layout in ("layered", "side_by_side"):
                    pv.layout_var.set(layout)
                    for pad, inner, n_strokes in ((10.0, 1, 3), (7.0, 0, 1), (20.0, 1, 2)):
                        pv.preview_size_var.set(pad)
                        pv._render()
                        rings = pv.canvas.find_withtag("ring")
                        nums = pv.canvas.find_withtag("ring_number")
                        if not enabled:
                            assert not rings and not nums, "ring drawn while the feature is off"
                            continue
                        assert len(rings) == 1 + inner, f"{layout} {pad}: {len(rings)} ring ovals"
                        assert len(nums) == n_strokes, f"{layout} {pad}: {len(nums)} number strokes"
                assert bool(pv._mat_checkbuttons[RING].winfo_manager()) == enabled
                # The checkbox follows the form live.
                w.ring_enabled_var.set(not enabled)
                pv._render()
                assert bool(pv._mat_checkbuttons[RING].winfo_manager()) == (not enabled)
                assert bool(pv.canvas.find_withtag("ring")) == (not enabled)
            finally:
                pv.destroy()
                w.top.destroy()
    check("GUI: the pad preview shows the ring and its number only while enabled", pad_preview_shows_the_ring_and_number)

    def pad_preview_number_is_where_it_will_be_cut():
        w = options(on())
        pv = PadPreviewWindow(w)
        pv.withdraw()
        pv._cancel_poll()
        try:
            for mat in pv.MATERIALS:
                pv.show_vars[mat].set(mat == RING)
            for pad, under in ((10.0, True), (20.0, False)):
                pv.preview_size_var.set(pad)
                pv._render()
                ring_box = pv.canvas.bbox(pv.canvas.find_withtag("ring")[0])
                num_box = pv.canvas.bbox("ring_number")
                if under:
                    assert num_box[1] >= ring_box[3] - 2, f"{pad}: number should be under the ring"
                else:
                    assert (ring_box[0] < num_box[0] and num_box[2] < ring_box[2]
                            and ring_box[1] < num_box[1] and num_box[3] < ring_box[3]), f"{pad}: number should be in the hole"
        finally:
            pv.destroy()
            w.top.destroy()
    check("GUI: the pad preview puts the number where it will really go", pad_preview_number_is_where_it_will_be_cut)

    def nesting_preview_draws_the_real_ring():
        s = on()
        placed, _z, _fp, _ft = nest(s)
        got = {}

        def probe(with_settings):
            # The window blocks in wait_window() inside its constructor, so the
            # inspection is scheduled first and runs inside that wait: Tk fires
            # no timer before the constructor reaches it, so the window always
            # exists by then, however slow the machine. Whatever happens in
            # go(), the window is closed, or the constructor would never return.
            tries = []

            def go():
                wins = [c for c in root.winfo_children() if isinstance(c, NestingPreviewWindow)]
                if not wins and len(tries) < 400:
                    tries.append(1)          # not built yet (should not happen): look again, don't guess a delay
                    root.after(25, go)
                    return
                try:
                    win = wins[0]
                    cv = win._canvas
                    cv.winfo_width = lambda: 640
                    cv.winfo_height = lambda: 480
                    win._draw()
                    got[with_settings] = {
                        "rings": len(cv.find_withtag("ring")),
                        "inner": len(cv.find_withtag("ring_inner")),
                        "numbers": len(cv.find_withtag("ring_number")),
                        "ovals": sum(1 for i in cv.find_all() if cv.type(i) == "oval"),
                        "widest": max(cv.bbox(i)[2] - cv.bbox(i)[0] for i in cv.find_all() if cv.type(i) == "oval"),
                    }
                except Exception as e:      # reported by the asserts below, not swallowed
                    got[with_settings] = {"error": repr(e)}
                finally:
                    for w in wins:
                        if w.winfo_exists():
                            w._on_adjust()
            root.after(0, go)
            NestingPreviewWindow(root, {RING: placed}, W, H, settings=s if with_settings else None)
        app.settings["seen_preview_tutorial"] = True
        probe(True)
        probe(False)
        holes = sum(1 for p, *_r in placed if se.adhesive_ring_geometry(p, s)["inner_r"] > 0)
        real, fallback = got[True], got[False]
        assert "error" not in real and "error" not in fallback, (real, fallback)
        assert real["rings"] == len(placed) + holes and real["inner"] == holes, real
        assert real["numbers"] >= len(placed), real
        assert fallback["rings"] == 0 and fallback["ovals"] == len(placed), fallback
    check("GUI: the nesting preview draws the real ring outline and number", nesting_preview_draws_the_real_ring)

    # ---- Generate buttons, Frame & Cut, job history ---------------------
    def fill_form(pads_text, w=W, h=H):
        app.settings['units'] = 'mm'
        app.settings['show_engraving_warning'] = False
        app.settings['gcode_cut_grouping'] = 'layer'
        app.preview_var.set(False)
        app.custom_polygon = None
        app.custom_polygon_outline = None
        app._custom_polygon_lb_machine = None
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", pads_text)
        app.width_entry.delete(0, tk.END)
        app.width_entry.insert(0, str(w))
        app.height_entry.delete(0, tk.END)
        app.height_entry.insert(0, str(h))
        app.filename_entry.delete(0, tk.END)
        app.filename_entry.insert(0, "rings")
        for m, v in app.material_vars.items():
            v.set(m == RING)
        root.update()

    PAD_TEXT = "7.0 x 3\n12.5 x 2\n20.0 x 2"

    class FakeSender:
        def __init__(self, port=None):
            self.port = port

        def connect(self):
            pass

        def unlock(self):
            pass

        def disconnect(self):
            pass

    runs = []
    previews = []

    class FakeRun:
        script = []      # how each dialog in turn ends; "complete" once it runs out

        def __init__(self, parent, sender, lines, **kw):
            runs.append({"lines": list(lines), "kw": kw})
            self._final_reason = FakeRun.script.pop(0) if FakeRun.script else "complete"

    class FakePreview:
        def __init__(self, parent, placements, w_mm, h_mm, **kw):
            previews.append({"placements": placements, "kw": kw})
            self.result = "save"

    def frame_and_cut():
        runs.clear()
        previews.clear()
        boxes.clear()
        keep = (ui_dialogs.FalconRunDialog, main_mod.NestingPreviewWindow, falcon_sender.FalconSender)
        ui_dialogs.FalconRunDialog = FakeRun
        main_mod.NestingPreviewWindow = FakePreview
        falcon_sender.FalconSender = FakeSender
        app.falcon_port = "TEST"
        app.live_camera_var.set(False)
        try:
            app.on_frame_and_cut()
        finally:
            ui_dialogs.FalconRunDialog, main_mod.NestingPreviewWindow, falcon_sender.FalconSender = keep
            app.falcon_port = None
        errs = [b for b in boxes if b[0] == "showerror"]
        assert not errs, errs
        return runs, previews

    def frame_and_cut_streams_what_the_file_export_writes():
        set_feature(True)
        app.settings["laser_framing_power_by_material"][RING] = 23
        app.settings["laser_framing_power_by_material"]["felt"] = 10
        app.scrap_mode_var.set(False)
        fill_form(PAD_TEXT)
        # Rings are cut from the sheet (or the traced shape); the paper-size option is card's alone.
        assert app._get_material_dimensions(RING, 120.0, 90.0, (215.9, 279.4)) == (120.0, 90.0, None)
        assert app._get_material_dimensions("card", 120.0, 90.0, (215.9, 279.4)) == (215.9, 279.4, None)
        for old in glob.glob(os.path.join(out_dir, "*")):
            os.remove(old)
        boxes.clear()
        app.on_generate_gcode()
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], boxes
        files = glob.glob(os.path.join(out_dir, "*.gcode"))
        assert [os.path.basename(f) for f in files] == ["rings_adhesive_ring.gcode"], files
        exported = open(files[0], encoding="utf-8").read().splitlines()
        assert any(ln.strip() == "; Ring outer cut" for ln in exported)

        got_runs, got_previews = frame_and_cut()
        assert len(got_previews) == 1, "the nesting preview comes first"
        assert list(got_previews[0]["placements"]) == [RING]
        assert got_previews[0]["kw"].get("settings") is app.settings, "preview needs the settings to draw rings"
        assert len(got_runs) == 3, f"jog, framing, cut expected; got {len(got_runs)}"
        jog, framing, cut = got_runs
        assert jog["lines"] == [] and jog["kw"].get("remind_lid_on_advance") is True
        assert framing["kw"].get("loop") is True
        assert framing["lines"][0] == "G92 X0.000 Y0.000"
        m4 = [ln for ln in framing["lines"] if ln.startswith("M4") or ln.startswith("M3")]
        assert m4 and all(ln == "M4 S23" for ln in m4), f"framing must use the ring's power with M4: {m4}"
        assert cut["lines"][0] == "G92 X0.000 Y0.000"
        assert cut["lines"][1:] == exported, "Frame & Cut streamed something other than the exported file"
        kept = open(os.path.join(_PROFILE, "last_frame_cut.gcode"), encoding="utf-8").read().splitlines()
        assert kept == cut["lines"]
    check("GUI: Frame & Cut takes rings as its one material and streams the exported G-code, framing at the ring's power",
          frame_and_cut_streams_what_the_file_export_writes)

    def frame_and_cut_refuses_two_materials_and_exact_size_rule_stands():
        set_feature(True)
        app.scrap_mode_var.set(False)
        fill_form(PAD_TEXT)
        app.material_vars["felt"].set(True)
        runs.clear()
        previews.clear()
        boxes.clear()
        keep = (ui_dialogs.FalconRunDialog, main_mod.NestingPreviewWindow, falcon_sender.FalconSender)
        ui_dialogs.FalconRunDialog, main_mod.NestingPreviewWindow, falcon_sender.FalconSender = FakeRun, FakePreview, FakeSender
        app.falcon_port = "TEST"
        try:
            app.on_frame_and_cut()
        finally:
            ui_dialogs.FalconRunDialog, main_mod.NestingPreviewWindow, falcon_sender.FalconSender = keep
            app.falcon_port = None
            app.material_vars["felt"].set(False)
        assert not runs and any(b[1] == "Pick One Material" for b in boxes), boxes
    check("GUI: Frame & Cut still wants exactly one material", frame_and_cut_refuses_two_materials_and_exact_size_rule_stands)

    def frame_and_cut_on_a_camera_polygon_flips_with_the_outline():
        set_feature(True)
        app.settings["laser_framing_power_by_material"][RING] = 23
        app.scrap_mode_var.set(False)
        fill_form(PAD_TEXT)
        outline = [(0.0, 0.0), (100.0, 0.0), (100.0, 80.0), (0.0, 80.0)]
        inset = [(5.0, 5.0), (95.0, 5.0), (95.0, 75.0), (5.0, 75.0)]
        app.custom_polygon = inset
        app.custom_polygon_outline = outline
        try:
            got_runs, got_previews = frame_and_cut()
        finally:
            app.custom_polygon = None
            app.custom_polygon_outline = None
        assert len(got_runs) == 3
        _jog, framing, cut = got_runs
        placed = got_previews[0]["placements"][RING]
        assert got_previews[0]["kw"].get("polygon") == inset
        want_frame = ge.generate_polygon_framing_gcode(outline, power_s=23,
                                                       feed=app.settings.get("laser_framing_feed", 2000))
        assert framing["lines"][1:] == want_frame, "framing should trace the outline at the ring's power"
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "want.gcode")
            ge.generate_gcode_from_placed(placed, RING, W, H, path, 0, app.settings,
                                          polygon=inset, flip_height_mm=80.0)
            want = open(path, encoding="utf-8").read().splitlines()
            ge.generate_gcode_from_placed(placed, RING, W, H, path, 0, app.settings, polygon=inset)
            unflipped = open(path, encoding="utf-8").read().splitlines()
        assert cut["lines"][1:] == want, "the cut must be flipped with the outline's height"
        assert want != unflipped, "fixture: the inset flip should differ from the outline flip"
        # Every ring cut lies inside the framed outline (work frame, Y-up).
        for kind, strokes in gcode_sections("\n".join(cut["lines"])):
            for st in strokes:
                for x, y in st:
                    assert 0 <= x <= 100 and 0 <= y <= 80, f"{kind} leaves the framed scrap: {(x, y)}"
    check("GUI: Frame & Cut on a camera polygon flips the rings with the outline and frames at the ring's power",
          frame_and_cut_on_a_camera_polygon_flips_with_the_outline)

    def frame_and_cut_in_scrap_mode_and_job_history():
        set_feature(True)
        continues = []
        app._show_scrap_continue_dialog = lambda *a: continues.append(a)
        app._open_remaining_pads_window = lambda: None
        app._update_remaining_pads_window = lambda: None
        app._clear_scrap_session()
        fill_form("12.5 x 30\n20.0 x 6", w=70, h=60)
        app.scrap_mode_var.set(True)
        try:
            got_runs, got_previews = frame_and_cut()
            assert len(got_runs) == 3
            placed = got_previews[0]["placements"][RING]
            assert 0 < len(placed) < 36
            sess = app.scrap_session
            assert sess["active"] and sess["material"] == RING and sess["scrap_count"] == 1
            assert sum(sess["done"].values()) == len(placed)
            assert sum(p["qty"] for p in sess["remaining_pads"]) == 36 - len(placed)
            assert len(continues) == 1
            assert str(app.material_checkboxes[RING].cget("state")) == "disabled", "materials lock for the session"
            with tempfile.TemporaryDirectory() as d:
                path = os.path.join(d, "want.gcode")
                ge.generate_gcode_from_placed(placed, RING, 70, 60, path, 0, app.settings)
                want = open(path, encoding="utf-8").read().splitlines()
            assert got_runs[2]["lines"][1:] == want
            jobs = config.load_job_history()
            assert jobs and jobs[0]["output"] == "laser" and jobs[0]["materials"] == [RING], jobs[0]
            assert jobs[0]["scrap_num"] == 1 and jobs[0]["placed_count"] == len(placed)
            assert JobHistoryWindow._materials_label(jobs[0]) == "Adhesive rings"
            # Switching the feature off mid-session ends the session with the material.
            set_feature(False)
            assert not app.scrap_session["active"] and RING not in app.material_vars
            assert not app.scrap_mode_var.get()
        finally:
            app._clear_scrap_session()
            app.scrap_mode_var.set(False)
            app._update_scrap_status_display()
        # Reloading the job ticks the material again once the feature is back on.
        set_feature(True)
        for v in app.material_vars.values():
            v.set(False)
        app._load_job_into_form(config.load_job_history()[0])
        assert [m for m, v in app.material_vars.items() if v.get()] == [RING]
    check("GUI: Frame & Cut in scrap mode, and job history records and reloads the material",
          frame_and_cut_in_scrap_mode_and_job_history)

    def frame_and_cut_scrap_mode_on_a_camera_polygon():
        # A camera-captured scrap: the outline as traced, and the same shape
        # inset for nesting. One scrap per click; a stopped cut consumes nothing.
        set_feature(True)
        continues = []
        app._show_scrap_continue_dialog = lambda *a: continues.append(a)
        app._open_remaining_pads_window = lambda: None
        app._update_remaining_pads_window = lambda: None
        app._clear_scrap_session()
        app.settings["laser_framing_power_by_material"][RING] = 23
        fill_form("12.5 x 30\n20.0 x 6", w=300, h=300)
        outline = [(0.0, 0.0), (90.0, 0.0), (90.0, 70.0), (0.0, 70.0)]
        inset = [(4.0, 4.0), (86.0, 4.0), (86.0, 66.0), (4.0, 66.0)]
        app.custom_polygon, app.custom_polygon_outline = inset, outline
        app.scrap_mode_var.set(True)
        feed = app.settings.get("laser_framing_feed", 2000)

        def engine_cut(placed):
            with tempfile.TemporaryDirectory() as d:
                path = os.path.join(d, "want.gcode")
                ge.generate_gcode_from_placed(placed, RING, 300, 300, path, 0, app.settings,
                                              polygon=inset, flip_height_mm=70.0)
                return open(path, encoding="utf-8").read().splitlines()
        try:
            # 1) The user stops the cut: the scrap is not counted.
            FakeRun.script = ["complete", "complete", "stopped"]
            got_runs, got_previews = frame_and_cut()
            assert len(got_runs) == 3
            sess = app.scrap_session
            assert sess["active"] and sess["material"] == RING and sess["scrap_count"] == 0
            assert not sess["done"] and not sess.get("pending"), (sess["done"], sess.get("pending"))
            assert sum(p["qty"] for p in sess["remaining_pads"]) == 36 and not continues
            newest = config.load_job_history()[0]
            assert newest["status"] == "stopped" and newest["materials"] == [RING], newest

            # 2) The same scrap again, cut to the end.
            got_runs, got_previews = frame_and_cut()
            _jog, framing, cut = got_runs
            placed1 = got_previews[0]["placements"][RING]
            assert got_previews[0]["kw"].get("polygon") == inset
            assert framing["lines"] == ["G92 X0.000 Y0.000"] + ge.generate_polygon_framing_gcode(
                outline, power_s=23, feed=feed), "framing: the outline, at the ring's power"
            assert cut["lines"] == ["G92 X0.000 Y0.000"] + engine_cut(placed1), \
                "the cut is the engine's output flipped with the outline's height"
            for kind, strokes in gcode_sections("\n".join(cut["lines"])):
                for st in strokes:
                    for x, y in st:
                        assert 4 - 0.3 <= x <= 86 + 0.3 and 4 - 0.3 <= y <= 66 + 0.3, \
                            f"{kind} leaves the inset scrap: {(x, y)}"
            assert sess is app.scrap_session and app.scrap_session["scrap_count"] == 1
            assert sum(app.scrap_session["done"].values()) == len(placed1) and len(continues) == 1

            # 3) The next scrap takes what is left, not the list from the top.
            got_runs, got_previews = frame_and_cut()
            placed2 = got_previews[0]["placements"][RING]
            assert got_runs[2]["lines"][1:] == engine_cut(placed2)
            assert app.scrap_session["scrap_count"] == 2
            assert sum(app.scrap_session["done"].values()) == len(placed1) + len(placed2)
            left = sum(p["qty"] for p in app.scrap_session["remaining_pads"])
            assert left == 36 - len(placed1) - len(placed2), left
            done_20 = app.scrap_session["done"].get(20.0, 0)
            assert done_20 <= 6, f"more 20.0 rings cut ({done_20}) than were asked for"
        finally:
            FakeRun.script = []
            app.custom_polygon = None
            app.custom_polygon_outline = None
            app._clear_scrap_session()
            app.scrap_mode_var.set(False)
            app._update_scrap_status_display()
    check("GUI: Frame & Cut in scrap mode on a camera polygon: stopped cut consumes nothing, "
          "then two scraps in a row", frame_and_cut_scrap_mode_on_a_camera_polygon)

    def svg_button_and_scrap_file_export():
        set_feature(True)
        app._clear_scrap_session()
        app.scrap_mode_var.set(False)
        fill_form(PAD_TEXT)
        for old in glob.glob(os.path.join(out_dir, "*")):
            os.remove(old)
        boxes.clear()
        app.on_generate_svg()
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], boxes
        svgs = glob.glob(os.path.join(out_dir, "*.svg"))
        assert [os.path.basename(f) for f in svgs] == ["rings_adhesive_ring.svg"], svgs
        text = open(svgs[0], encoding="utf-8").read()
        assert all(c == app.settings["layer_colors"]["adhesive_ring_cut"] for _t, c in svg_strokes(text))
        # Zones on + 'max' on a zoned size is refused for card, but a rings-only job has no zones to clash with.
        app.settings["zone_labels_enabled"] = True
        try:
            fill_form("10.0 x max\n20.0 x 2", w=60, h=50)
            boxes.clear()
            app.on_generate_svg()
            root.update()
            assert not [b for b in boxes if b[0] == "showerror"], boxes
        finally:
            app.settings["zone_labels_enabled"] = False
        # Scrap-mode file export.
        app._show_scrap_continue_dialog = lambda *a: None
        app._open_remaining_pads_window = lambda: None
        app._update_remaining_pads_window = lambda: None
        fill_form("12.5 x 30", w=60, h=50)
        app.scrap_mode_var.set(True)
        try:
            boxes.clear()
            app.on_generate_gcode()
            root.update()
            assert not [b for b in boxes if b[0] == "showerror"], boxes
            assert os.path.exists(os.path.join(out_dir, "rings_adhesive_ring_scrap1.gcode"))
        finally:
            app._clear_scrap_session()
            app.scrap_mode_var.set(False)
            app._update_scrap_status_display()
    check("GUI: Generate SVG, the zones/max rule, and scrap-mode file export work for rings",
          svg_button_and_scrap_file_export)


def main():
    print("Adhesive rings")
    print("=" * 60)
    for name, fn in sorted(list(globals().items())):
        if name.startswith('test_') and callable(fn):
            check(name[5:].replace('_', ' '), fn)

    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"        (GUI cases skipped: no display: {e})")
        root = None
    if root is not None:
        root.withdraw()
        try:
            gui_tests(root)
        finally:
            try:
                root.destroy()
            except tk.TclError:
                pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == '__main__':
    sys.exit(main())
