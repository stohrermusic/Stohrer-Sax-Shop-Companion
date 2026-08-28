"""
Tests for labeled zones (svg_engine + gcode_engine).

Small pads are hard to tell apart once they're off the laser: a 7.0 and a
7.5 disc look identical, and the per-disc number is either dropped by the
font gate or unreadable. Leather can't be marked in the middle at all —
that's the sealing surface — so a zone puts the size on the WASTE instead:
a compact grid of one size, a rectangle drawn round it, and the size
engraved on it.

Groups are nested as units, like oversized pads: on a rectangular sheet
they shelf-pack into a band along the bottom, and on a traced scrap they
tuck in wherever they fit.

What these pin down:
  - zones are opt-in and change nothing when off (regression safety)
  - only in-range, fixed-quantity pads get zoned
  - grids read the way you'd lay pads out by hand: 6 as 3x2, 9 as 3x3,
    8 as an exact 4x2 rather than 3x3-with-a-hole
  - no disc escapes its group, no group overlaps another, and groups stay
    a visible distance apart (touching groups defeat the whole point)
  - groups and their engraved boundaries land on material, not off the
    scrap edge
  - a zoned size is never cut WITHOUT its group — it's left unplaced and
    reported instead, because an unlabeled pile is the failure this
    feature exists to prevent
  - SVG and G-code describe the same rectangle and label (the Y-flip)

No GUI — pure engine calls. Runs headless.

Run:
    python tools/test_zone_labels.py
"""
import math
import os
import re
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import config  # noqa: E402
import svg_engine as se  # noqa: E402
import gcode_engine as ge  # noqa: E402

results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def base_settings(enabled=True, lo=7.0, hi=12.5):
    s = dict(config.DEFAULT_SETTINGS)
    s['zone_labels_enabled'] = enabled
    s['zone_label_min_size'] = lo
    s['zone_label_max_size'] = hi
    s['edge_bias'] = 'center'
    return s


PADS = [
    {'size': 24.0, 'qty': 3},
    {'size': 18.0, 'qty': 4},
    {'size': 12.0, 'qty': 8},
    {'size': 10.0, 'qty': 8},
    {'size': 7.5, 'qty': 10},
    {'size': 7.0, 'qty': 10},
]
W = H = 200.0

# A traced hide offcut, irregular and concave — the shape Matt actually
# works from. Plus the honest job: ~10 each of a few neighbouring octave
# sizes on one piece.
SCRAP = [(6, 42), (34, 14), (72, 6), (112, 10), (146, 26), (158, 58),
         (150, 86), (116, 104), (74, 110), (32, 96), (12, 72)]
BAND_PADS = [{'size': 9.0, 'qty': 10}, {'size': 8.0, 'qty': 10},
             {'size': 7.0, 'qty': 10}]

# A second offcut, for the 'x max' leftover tests: 254 x 162mm (10 x 6.4in).
# Nothing here is ever bigger than 14 x 14in and is usually smaller.
OFFCUT = [(8, 60), (70, 18), (160, 10), (240, 26), (262, 80), (250, 140),
          (170, 172), (80, 166), (16, 130)]
# Fill size for the 'x max' tests. Must be OUTSIDE the zone range so it
# stays ungrouped. 18.0 is a real pad; the 4.0 this used to be is a 1.2mm
# card disc that nobody cuts, and it made these two tests take 70 and 88
# seconds by placing ~2900 of them.
MAX_FILL = 18.0


def _overlaps(placed):
    n = 0
    for i in range(len(placed)):
        for j in range(i + 1, len(placed)):
            _, x1, y1, r1 = placed[i]
            _, x2, y2, r2 = placed[j]
            if math.hypot(x1 - x2, y1 - y2) < r1 + r2 - 1e-6:
                n += 1
    return n


def _counts(placed):
    out = {}
    for size, _, _, _ in placed:
        out[size] = out.get(size, 0) + 1
    return out


# --------------------------------------------------------------------------
# Opt-in / regression safety
# --------------------------------------------------------------------------

def test_disabled_matches_plain_nest():
    """Zones off must reproduce the old nester exactly, placement for
    placement — this feature must not perturb existing jobs."""
    s = base_settings(enabled=False)
    old, _, _ = se._nest_discs(PADS, 'card', W, H, s)
    new, zones, _, _ = se.nest_with_zones(PADS, 'card', W, H, s)
    assert zones == [], f"expected no zones when disabled, got {len(zones)}"
    assert len(old) == len(new), f"count drift: {len(old)} vs {len(new)}"
    for a, b in zip(old, new):
        assert a == b, f"placement drift: {a} != {b}"


def test_polygon_produces_nested_groups():
    """On a traced scrap each zoned size becomes one group — a compact grid
    with a rectangle round it — and the groups are nested into the shape as
    units, rather than each claiming a full-width horizontal band."""
    s = base_settings()
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert placed, "polygon nesting should still place discs"
    assert zones, "polygon sheets should produce groups"
    for z in zones:
        assert z['shape'] == 'rect', f"expected a group rectangle, got {z['shape']}"
        for key in ('x', 'y', 'w', 'h', 'cols', 'rows'):
            assert key in z, f"group missing {key}"


def test_polygon_zones_off_when_disabled():
    s = base_settings(enabled=False)
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert zones == [], "disabled must produce no groups on a polygon either"
    assert placed


def test_groups_do_not_overlap_on_a_polygon():
    """Groups are nested as units, so nothing may collide — otherwise two
    sizes end up sharing space and the whole point is lost."""
    s = base_settings()
    _, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for i in range(len(zones)):
        for j in range(i + 1, len(zones)):
            a, b = zones[i], zones[j]
            overlap = (a['x'] < b['x'] + b['w'] and b['x'] < a['x'] + a['w']
                       and a['y'] < b['y'] + b['h'] and b['y'] < a['y'] + a['h'])
            assert not overlap, f"groups {a['label']} and {b['label']} overlap"


def test_groups_are_separated_by_a_visible_gap():
    """A 7.0 and a 7.5 disc look identical, so touching groups would be as
    bad as no grouping at all."""
    s = base_settings()
    gap = config.DEFAULT_SETTINGS['zone_group_gap_mm']
    _, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for i in range(len(zones)):
        for j in range(i + 1, len(zones)):
            a, b = zones[i], zones[j]
            dx = max(b['x'] - (a['x'] + a['w']), a['x'] - (b['x'] + b['w']), 0)
            dy = max(b['y'] - (a['y'] + a['h']), a['y'] - (b['y'] + b['h']), 0)
            assert max(dx, dy) >= gap - 0.51, (
                f"{a['label']} and {b['label']} are only "
                f"{max(dx, dy):.1f}mm apart")


def test_both_sheet_types_space_boxes_the_same():
    """Same divergence class as the grid-shape rule: the gap was a 2.0
    constant on rectangular sheets and a 6.0 setting on traced polygons,
    so the same job spaced its boxes differently by sheet type."""
    s = base_settings()
    gutter, border, font, gap = se._group_metrics(s)
    specs, _ = se.plan_zone_specs(BAND_PADS, 'felt', s)
    rect_zones, _, _ = se._shelf_pack_zones(specs, 300.0, 200.0, group_gap=gap)
    _, poly_zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s,
                                             polygon=SCRAP)
    assert rect_zones and poly_zones, "need boxes on both paths to compare"

    by_label_rect = {z['label']: z for z in rect_zones}
    for pz in poly_zones:
        rz = by_label_rect.get(pz['label'])
        assert rz is not None, f"{pz['label']} missing from the rect layout"
        assert abs(rz['w'] - pz['w']) < 1e-6 and abs(rz['h'] - pz['h']) < 1e-6, (
            f"{pz['label']}: box is {rz['w']:.2f}x{rz['h']:.2f} on a rect "
            f"sheet but {pz['w']:.2f}x{pz['h']:.2f} on a polygon")
        assert abs(rz['border'] - pz['border']) < 1e-6, (
            f"{pz['label']}: border differs by sheet type")


def test_group_gap_is_not_a_moat():
    """The 6.0mm gap was sized to separate full-width BANDS. Once each size
    became a box with its own border, that much space between boxes was
    just wasted material — and on leather it cost whole groups their place
    on the scrap."""
    s = base_settings()
    _gutter, border, _font, gap = se._group_metrics(s)
    assert gap <= border * 2, (
        f"gap {gap}mm dwarfs the {border}mm each box already carries "
        f"inside its own border")
    assert gap > 0, "boxes must not touch — two abutting lines read as one"

    # Tighter packing must never cost placements.
    wide = dict(s)
    wide['zone_group_gap_mm'] = 6.0
    _, tight_zones, tight_placed, total = se.nest_with_zones(
        BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    _, wide_zones, wide_placed, _ = se.nest_with_zones(
        BAND_PADS, 'felt', 0, 0, wide, polygon=SCRAP)
    assert tight_placed >= wide_placed, (
        f"tight gap placed {tight_placed}/{total} vs {wide_placed} wide")
    assert len(tight_zones) >= len(wide_zones)


def test_polygon_group_discs_stay_inside_their_rectangle():
    s = base_settings()
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for z in zones:
        mine = [(cx, cy, r) for sz, cx, cy, r in placed if sz == z['size']]
        assert len(mine) == z['qty']
        for cx, cy, r in mine:
            assert z['x'] <= cx - r + 1e-9 and cx + r <= z['x'] + z['w'] + 1e-9, \
                f"{z['label']}: disc escapes its group horizontally"
            assert z['y'] <= cy - r + 1e-9 and cy + r <= z['y'] + z['h'] + 1e-9, \
                f"{z['label']}: disc escapes its group vertically"


def test_polygon_groups_and_their_boundaries_land_on_material():
    """Both the pads and the engraved rectangle have to be on the scrap —
    a boundary drawn off the edge marks nothing."""
    s = base_settings()
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for size, cx, cy, r in placed:
        assert se._circle_fits_in_polygon(cx, cy, r, SCRAP, 0.0), \
            f"{size}mm disc at ({cx:.1f},{cy:.1f}) is not on the scrap"
    for z in zones:
        for px, py in ((z['x'], z['y']), (z['x'] + z['w'], z['y']),
                       (z['x'], z['y'] + z['h']),
                       (z['x'] + z['w'], z['y'] + z['h'])):
            assert se._point_in_polygon(px, py, SCRAP), \
                f"{z['label']}: group corner ({px:.1f},{py:.1f}) is off the scrap"


def test_grid_shapes_match_how_you_would_lay_them_out():
    """Six reads as 3x2 and nine as 3x3. Eight prefers an exact 4x2 over a
    3x3 with a hole in it; a prime like seven takes 4x2 with one gap rather
    than a row of seven."""
    expected = {2: (2, 1), 4: (2, 2), 6: (3, 2), 7: (4, 2),
                8: (4, 2), 9: (3, 3), 12: (4, 3), 16: (4, 4)}
    for qty, want in expected.items():
        got = se.zone_grid_candidates(qty)[0]
        assert got == want, f"{qty} pads: expected {want}, got {got}"


def test_grid_candidates_offer_flatter_fallbacks():
    """An awkward scrap must be able to degrade to a flatter grid instead
    of the size being refused outright."""
    for qty in (6, 9, 12):
        cands = se.zone_grid_candidates(qty)
        assert len(cands) > 1, f"{qty}: no fallback shapes"
        assert len(set(cands)) == len(cands), f"{qty}: duplicate shapes"
        rows = [r for _c, r in cands]
        assert min(rows) == 1, f"{qty}: no single-row fallback available"


def test_clip_polygon_y():
    square = [(0, 0), (10, 0), (10, 10), (0, 10)]
    mid = se._clip_polygon_y(square, 2, 8)
    ys = [p[1] for p in mid]
    assert abs(min(ys) - 2) < 1e-9 and abs(max(ys) - 8) < 1e-9, f"bad clip: {mid}"
    assert se._clip_polygon_y(square, 20, 30) == [], "clip outside should be empty"
    concave = [(0, 0), (10, 0), (10, 10), (5, 4), (0, 10)]
    got = se._clip_polygon_y(concave, 0, 5)
    assert len(got) >= 3, "concave clip collapsed"


def test_no_zoned_size_is_ever_cut_unlabeled():
    """The failure mode that matters. If a zoned size can't get a group it
    must NOT be quietly nested alongside the labeled ones — that produces
    exactly the unlabeled mixed pile zones exist to prevent, while the
    caller sees a full placed count and reports success."""
    s = base_settings()
    lo, hi = s['zone_label_min_size'], s['zone_label_max_size']
    tiny = [(0, 0), (70, 0), (70, 55), (0, 55)]
    pads = [{'size': 12.0, 'qty': 6}, {'size': 8.5, 'qty': 9},
            {'size': 7.5, 'qty': 6}, {'size': 7.0, 'qty': 6}]
    placed, zones, fixed_placed, fixed_total = se.nest_with_zones(
        pads, 'leather', 0, 0, s, polygon=tiny)

    grouped = {z['size'] for z in zones}
    cut = {size for size, _cx, _cy, _r in placed}
    unlabeled = [sz for sz in cut if lo <= sz <= hi and sz not in grouped]
    assert not unlabeled, f"zoned sizes cut without a group: {unlabeled}"
    assert fixed_placed == len(placed)
    if fixed_placed < fixed_total:
        assert se.can_all_pads_fit(pads, 'leather', 0, 0, s, polygon=tiny) is False


def test_group_label_sits_in_its_own_rectangle():
    s = base_settings()
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for z in zones:
        label_bottom = z['y'] + z['border'] + z['font']
        top_disc = min(cy - r for sz, _cx, cy, r in placed if sz == z['size'])
        assert top_disc >= label_bottom - 1e-9, \
            f"{z['label']}: discs run into the label strip"


def test_polygon_svg_and_gcode_groups_agree():
    """Same Y-flip contract as the rectangular-sheet path."""
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert zones, "no groups to compare"
    flip = max(p[1] for p in SCRAP)

    with tempfile.TemporaryDirectory() as d:
        gpath = os.path.join(d, 'g.gcode')
        ge.generate_gcode_from_placed(placed, 'felt', 200, 200, gpath, 0, s,
                                      polygon=SCRAP, zones=zones)
        body = open(gpath, encoding='utf-8').read()

    pts = []
    for ln in body.splitlines():
        if not (ln.startswith('G0') or ln.startswith('G1')):
            continue
        xm = re.search(r'X(-?\d+(?:\.\d+)?)', ln)
        ym = re.search(r'Y(-?\d+(?:\.\d+)?)', ln)
        if xm and ym:
            pts.append((float(xm.group(1)), float(ym.group(1))))
    assert pts, "no coordinates in the G-code"

    for z in zones:
        want = [(z['x'], flip - (z['y'] + z['h'])),
                (z['x'] + z['w'], flip - (z['y'] + z['h'])),
                (z['x'] + z['w'], flip - z['y']),
                (z['x'], flip - z['y'])]
        for wx, wy in want:
            hit = any(abs(px - wx) < 0.01 and abs(py - wy) < 0.01 for px, py in pts)
            assert hit, (f"group {z['label']}: corner ({wx:.2f},{wy:.2f}) "
                         f"missing from G-code — Y-flip mismatch")


def test_polygon_svg_has_group_rectangles():
    s = base_settings()
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'g.svg')
        placed, zones = se.nest_pads_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
        se.generate_svg_from_placed(placed, 'felt', 200, 200, path, 0, s,
                                    polygon=SCRAP, zones=zones)
        body = open(path, encoding='utf-8').read()
    assert '<rect' in body, "no group rectangle in the SVG"
    for z in zones:
        assert f">{z['label']}<" in body, f"group label {z['label']} missing"


def test_max_fill_reaches_past_the_groups():
    """An ungrouped 'x max' has to fill the whole leftover of a scrap.
    While the layout was stacked full-width bands the leftover was the
    strip UNDER the lowest group, so a max fill placed nothing at all
    beside them — most of the piece went to waste.

    'Beside' is the sharp signal, not 'above': groups are placed
    biggest-first scanning from the top, so they claim the topmost
    material and there is structurally nothing above them for a pad of
    real size. The old fixture only got pads up there because a 4.0mm
    pad is a 1.2mm disc that fits in slivers — and no one makes those.
    """
    s = base_settings()
    pads = [{'size': 12.0, 'qty': 6}, {'size': 8.5, 'qty': 6},
            {'size': MAX_FILL, 'qty': 'max'}]
    placed, zones, _, _ = se.nest_with_zones(pads, 'card', 0, 0, s, polygon=OFFCUT)
    assert zones, "expected groups"
    g_top = min(z['y'] for z in zones)
    g_bot = max(z['y'] + z['h'] for z in zones)
    ys = [cy for sz, _cx, cy, _r in placed if sz == MAX_FILL]
    assert ys, "max fill placed nothing"
    assert sum(1 for y in ys if g_top <= y <= g_bot) > 0, \
        "nothing filled beside the groups — leftovers clipped to below them?"

    # The whole outline must stay available, not just the part the groups
    # didn't want: compare against the same fill with zones switched off.
    off = base_settings(enabled=False)
    bare, _, _ = se._nest_discs([{'size': MAX_FILL, 'qty': 'max'}], 'card',
                                0, 0, off, 1.0, OFFCUT)
    assert len(ys) >= 0.85 * len(bare), (
        f"zoned max fill placed {len(ys)} vs {len(bare)} ungrouped — the "
        f"groups cost more room than they occupy")


def test_max_fill_stays_out_of_group_boxes():
    """Loose pads must not land inside a labeled group's rectangle — a
    stray disc in another size's box is the confusion this prevents.
    Seeding the group's discs alone isn't enough; the whole footprint has
    to be reserved or pads settle into the gaps between them. (Verified
    to still bite: stubbing the rectangle reservation out puts 6 of these
    pads inside a box.)"""
    s = base_settings()
    pads = [{'size': 12.0, 'qty': 6}, {'size': 8.5, 'qty': 6},
            {'size': MAX_FILL, 'qty': 'max'}]
    placed, zones, _, _ = se.nest_with_zones(pads, 'card', 0, 0, s, polygon=OFFCUT)
    grouped = {z['size'] for z in zones}
    for size, cx, cy, r in placed:
        if size in grouped:
            continue
        for z in zones:
            dx = max(z['x'] - cx, 0, cx - (z['x'] + z['w']))
            dy = max(z['y'] - cy, 0, cy - (z['y'] + z['h']))
            assert math.hypot(dx, dy) - r >= -1e-6, (
                f"{size}mm pad at ({cx:.1f},{cy:.1f}) intrudes into group "
                f"{z['label']}")


def test_preplaced_seeds_never_leak_into_results():
    """The reservation circles are space markers, not pads. If they came
    back in `placed` they'd be cut."""
    s = base_settings()
    scrap = [(0, 0), (200, 0), (200, 150), (0, 150)]
    pads = [{'size': 12.0, 'qty': 6}, {'size': MAX_FILL, 'qty': 'max'}]
    placed, _zones, _, _ = se.nest_with_zones(pads, 'card', 0, 0, s, polygon=scrap)
    for entry in placed:
        assert not isinstance(entry[0], str), f"reservation seed leaked: {entry}"


def test_max_qty_never_zoned():
    """'max' pads have no fixed count, so they can't be gridded."""
    s = base_settings()
    pads = [{'size': 10.0, 'qty': 'max'}, {'size': 7.5, 'qty': 6}]
    _, zones, _, _ = se.nest_with_zones(pads, 'card', W, H, s)
    assert all(z['size'] != 10.0 for z in zones), "'max' pad was zoned"
    assert any(z['size'] == 7.5 for z in zones), "fixed pad should be zoned"


def test_only_in_range_sizes_zoned():
    s = base_settings(lo=7.0, hi=12.5)
    _, zones, _, _ = se.nest_with_zones(PADS, 'card', W, H, s)
    zoned = sorted(z['size'] for z in zones)
    assert zoned == [7.0, 7.5, 10.0, 12.0], f"unexpected zoned sizes: {zoned}"


def test_range_bounds_inclusive():
    s = base_settings(lo=7.0, hi=12.0)
    pads = [{'size': 7.0, 'qty': 4}, {'size': 12.0, 'qty': 4},
            {'size': 12.5, 'qty': 4}]
    _, zones, _, _ = se.nest_with_zones(pads, 'card', W, H, s)
    zoned = sorted(z['size'] for z in zones)
    assert zoned == [7.0, 12.0], f"bounds not inclusive/exclusive as expected: {zoned}"


# --------------------------------------------------------------------------
# Zone geometry
# --------------------------------------------------------------------------

def test_grid_is_a_block_not_a_strip():
    """No zone may come out as a long thin ribbon — those shelf-pack badly
    and are impossible to count."""
    for mat in ('card', 'felt', 'leather'):
        _, zones, _, _ = se.nest_with_zones(PADS, mat, 300.0, 300.0, base_settings())
        for z in zones:
            assert z['cols'] > 1 and z['rows'] > 1 or z['qty'] <= 2, (
                f"{mat} {z['label']}: degenerate {z['cols']}x{z['rows']} grid")
            aspect = max(z['w'], z['h']) / min(z['w'], z['h'])
            assert aspect < 6.0, f"{mat} {z['label']}: aspect {aspect:.1f} too extreme"


def test_both_sheet_types_pick_the_same_grid():
    """One rule, both paths. These disagreed once — a rectangular sheet
    scored zone area on its own and produced 6 as 2x3, 8 as 2x4, and a
    prime like 7 as a 1x7 column, while a traced scrap gave 3x2, 4x2 and
    4x2. Six pads should read as 3x2 wherever they're cut."""
    s = base_settings(lo=3.0, hi=30.0)
    for qty in (2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 16, 20):
        specs, _ = se.plan_zone_specs([{'size': 10.0, 'qty': qty}], 'card', s)
        rect = (specs[0]['cols'], specs[0]['rows'])
        poly = se.zone_grid_candidates(qty)[0]
        assert rect == poly, (
            f"{qty} pads: rectangular sheet gives {rect}, polygon gives {poly}")


def test_no_quantity_produces_a_single_file_strip():
    """A prime count used to fall through to 1xN on the rectangular path.
    Nothing from 2 to 30 may produce a single row or column."""
    for qty in range(3, 31):
        cols, rows = se.zone_grid_candidates(qty)[0]
        assert cols > 1 and rows > 1, f"{qty} pads -> degenerate {cols}x{rows}"


def test_discs_stay_inside_their_zone():
    for mat in ('card', 'felt', 'leather'):
        _, zones, _, _ = se.nest_with_zones(PADS, mat, 300.0, 300.0, base_settings())
        for z in zones:
            r = z['disc_d'] / 2
            pos = se.zone_disc_positions(z)
            assert len(pos) == z['qty'], (
                f"{mat} {z['label']}: {len(pos)} positions for qty {z['qty']}")
            for cx, cy in pos:
                assert z['x'] <= cx - r + 1e-9 and cx + r <= z['x'] + z['w'] + 1e-9, \
                    f"{mat} {z['label']}: disc escapes horizontally"
                assert z['y'] <= cy - r + 1e-9 and cy + r <= z['y'] + z['h'] + 1e-9, \
                    f"{mat} {z['label']}: disc escapes vertically"


def test_label_never_overlaps_the_discs():
    """The label sits in a reserved strip at the top of the zone; the grid
    starts below it."""
    _, zones, _, _ = se.nest_with_zones(PADS, 'card', W, H, base_settings())
    for z in zones:
        label_bottom = z['y'] + z['border'] + z['font']
        top_disc_y = min(cy for _, cy in se.zone_disc_positions(z))
        assert top_disc_y - z['disc_d'] / 2 >= label_bottom - 1e-9, \
            f"{z['label']}: discs run into the label strip"


def test_zone_wide_enough_for_its_label():
    s = base_settings()
    _, zones, _, _ = se.nest_with_zones(PADS, 'card', W, H, s)
    for z in zones:
        text_w = se._zone_text_width_mm(z['label'], z['font'])
        assert z['w'] >= text_w, f"{z['label']}: zone narrower than its label"


def test_zones_do_not_overlap_each_other():
    for mat in ('card', 'felt', 'leather'):
        _, zones, _, _ = se.nest_with_zones(PADS, mat, 300.0, 300.0, base_settings())
        for i in range(len(zones)):
            for j in range(i + 1, len(zones)):
                a, b = zones[i], zones[j]
                overlap = (a['x'] < b['x'] + b['w'] and b['x'] < a['x'] + a['w']
                           and a['y'] < b['y'] + b['h'] and b['y'] < a['y'] + a['h'])
                assert not overlap, f"{mat}: zones {a['label']} and {b['label']} overlap"


def test_nothing_overlaps_and_nothing_leaves_the_sheet():
    for mat in ('card', 'felt', 'leather'):
        sheet = 200.0 if mat != 'leather' else 300.0
        placed, _, _, _ = se.nest_with_zones(PADS, mat, sheet, sheet, base_settings())
        assert _overlaps(placed) == 0, f"{mat}: discs overlap"
        for _, x, y, r in placed:
            assert x - r >= -1e-9 and y - r >= -1e-9, f"{mat}: disc off the sheet (low)"
            assert x + r <= sheet + 1e-9 and y + r <= sheet + 1e-9, \
                f"{mat}: disc off the sheet (high)"


def test_free_pads_stay_clear_of_the_band():
    """Free pads nest in the rectangle above the zone band. If that height
    were wrong they'd collide with the zones."""
    placed, zones, _, _ = se.nest_with_zones(PADS, 'card', W, H, base_settings())
    band_top = min(z['y'] for z in zones)
    zoned_sizes = {z['size'] for z in zones}
    for size, _, y, r in placed:
        if size not in zoned_sizes:
            assert y + r <= band_top + 1e-9, \
                f"free {size}mm pad at y={y:.2f} intrudes into the band at {band_top:.2f}"


# --------------------------------------------------------------------------
# Nothing gets lost
# --------------------------------------------------------------------------

def test_every_pad_is_accounted_for():
    placed, _, fixed_placed, fixed_total = se.nest_with_zones(
        PADS, 'card', W, H, base_settings())
    counts = _counts(placed)
    for pad in PADS:
        assert counts.get(pad['size'], 0) == pad['qty'], \
            f"{pad['size']}mm: placed {counts.get(pad['size'], 0)} of {pad['qty']}"
    assert fixed_placed == fixed_total == sum(p['qty'] for p in PADS)


def test_wide_zone_degrades_to_a_flatter_grid():
    """A block too wide for the sheet should try narrower shapes before
    giving up, rather than refusing the size outright."""
    s = base_settings()
    pads = [{'size': 12.0, 'qty': 12}]
    # 4x3 (the preferred shape) is ~50mm wide, so 45mm forces a fallback
    # while still leaving room for the narrower 3x4.
    _, zones, _, _ = se.nest_with_zones(pads, 'card', 45.0, 400.0, s)
    assert zones, "a narrower grid should have fitted"
    z = zones[0]
    assert z['w'] <= 45.0, "placed a zone wider than the sheet"
    assert (z['cols'], z['rows']) != se.zone_grid_candidates(12)[0], \
        "expected a fallback shape, not the preferred one"


def test_rect_sheet_never_cuts_a_zoned_size_unlabeled():
    """Same rule as the polygon path. A zone that can't be placed must
    leave its pads uncut and reported, not nested loose beside the
    labeled ones."""
    s = base_settings()
    lo, hi = s['zone_label_min_size'], s['zone_label_max_size']
    pads = [{'size': 12.0, 'qty': 40}]
    narrow = 30.0
    placed, zones, fixed_placed, fixed_total = se.nest_with_zones(
        pads, 'card', narrow, 400.0, s)
    grouped = {z['size'] for z in zones}
    cut = {size for size, _cx, _cy, _r in placed}
    unlabeled = [sz for sz in cut if lo <= sz <= hi and sz not in grouped]
    assert not unlabeled, f"zoned sizes cut without a group: {unlabeled}"
    assert fixed_total == 40, f"dropped pads must still be counted: {fixed_total}"
    if fixed_placed < fixed_total:
        assert se.can_all_pads_fit(pads, 'card', narrow, 400.0, s) is False
    for z in zones:
        assert z['w'] <= narrow, "a zone wider than the sheet was placed anyway"


def test_full_band_sheet_still_counts_free_pads():
    """When the band consumes the sheet exactly (free_h == 0), free pads
    have nowhere to go — they must still be COUNTED in fixed_total so
    can_all_pads_fit reports the shortfall, instead of the pads silently
    vanishing from both tallies and the job claiming success."""
    s = base_settings()
    zoned = {'size': 10.0, 'qty': 4}
    free = {'size': 24.0, 'qty': 2}
    specs, _ = se.plan_zone_specs([zoned], 'card', s)
    gap = float(s.get('zone_group_gap_mm', 1.0))
    # Height sized so the band fits exactly and nothing is left above it.
    h_exact = specs[0]['h'] + gap + se.ZONE_SHEET_MARGIN_MM
    _placed, zones, fixed_placed, fixed_total = se.nest_with_zones(
        [zoned, free], 'card', 200.0, h_exact, s)
    assert zones, "the zone band should still fit on the exact-height sheet"
    assert fixed_total == 6, \
        f"free pads dropped from fixed_total: {fixed_total} (expected 6)"
    assert fixed_placed == 4, \
        f"expected only the zoned pads placed: {fixed_placed} (expected 4)"
    assert se.can_all_pads_fit([zoned, free], 'card', 200.0, h_exact, s) is False


def test_can_all_pads_fit_still_works():
    s = base_settings()
    assert se.can_all_pads_fit(PADS, 'card', W, H, s) is True
    assert se.can_all_pads_fit(PADS, 'card', 30.0, 30.0, s) is False


# --------------------------------------------------------------------------
# Renderers: SVG and G-code must describe the same thing
# --------------------------------------------------------------------------

def test_svg_contains_zone_border_and_label():
    s = base_settings()
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'zones.svg')
        se.generate_svg(PADS, 'card', W, H, path, 0, s)
        body = open(path, encoding='utf-8').read()
    assert '<rect' in body, "no zone border rectangle in the SVG"
    for label in ('7', '7.5', '10', '12'):
        assert f'>{label}<' in body, f"zone label {label!r} missing from the SVG"


def test_gcode_contains_zone_strokes():
    s = base_settings()
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'zones.gcode')
        placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
        ge.generate_gcode_from_placed(placed, 'card', W, H, path, 0, s, zones=zones)
        body = open(path, encoding='utf-8').read()
    assert body.strip(), "empty G-code"
    # Every zone corner should appear as a rapid//cut coordinate somewhere.
    for z in zones:
        xs = f"X{z['x']:.3f}".rstrip('0').rstrip('.')
        assert xs[:6] in body, f"zone {z['label']} left edge missing from G-code"


def test_svg_and_gcode_zone_rectangles_agree():
    """The two engines render independently and have drifted before. The
    zone border must describe the same physical rectangle in both, which
    means the Y-flip has to be right."""
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
    assert zones, "no zones to compare"

    with tempfile.TemporaryDirectory() as d:
        gpath = os.path.join(d, 'z.gcode')
        ge.generate_gcode_from_placed(placed, 'card', W, H, gpath, 0, s, zones=zones)
        lines = open(gpath, encoding='utf-8').read().splitlines()

    # Collect every coordinate the G-code visits. Words are NOT
    # space-separated ("G1 X23.500Y2.000S100F1500"), so parse with a regex
    # rather than splitting on whitespace.
    pts = []
    for ln in lines:
        if not (ln.startswith('G0') or ln.startswith('G1')):
            continue
        xm = re.search(r'X(-?\d+(?:\.\d+)?)', ln)
        ym = re.search(r'Y(-?\d+(?:\.\d+)?)', ln)
        if xm and ym:
            pts.append((float(xm.group(1)), float(ym.group(1))))
    assert pts, "no coordinates in the G-code"

    for z in zones:
        # Expected G-code corners after the Y-flip.
        want = [
            (z['x'], H - (z['y'] + z['h'])),
            (z['x'] + z['w'], H - (z['y'] + z['h'])),
            (z['x'] + z['w'], H - z['y']),
            (z['x'], H - z['y']),
        ]
        for wx, wy in want:
            hit = any(abs(px - wx) < 0.01 and abs(py - wy) < 0.01 for px, py in pts)
            assert hit, (f"zone {z['label']}: corner ({wx:.2f},{wy:.2f}) "
                         f"missing from G-code — Y-flip mismatch")


def test_gcode_zone_border_is_engraved_not_cut():
    """A cut border would drop the zone tile through the bed slats."""
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'z.gcode')
        ge.generate_gcode_from_placed(placed, 'card', W, H, path, 0, s, zones=zones)
        body = open(path, encoding='utf-8').read()
    # Card layers are (engrave, hole, cut) = ('C15', 'C14', 'C01'). The first
    # layer emitted must be the engraving one, since zones go down first.
    first = None
    for ln in body.splitlines():
        for layer in ('C15', 'C14', 'C01'):
            if layer in ln:
                first = first or layer
        if first:
            break
    assert first == 'C15', f"zones should be engraved first, got layer {first}"


def test_line_and_filled_modes_both_render_zones():
    for mode in ('line', 'filled'):
        s = base_settings()
        s.setdefault('gcode_settings', {})
        s['gcode_settings'] = dict(s.get('gcode_settings', {}))
        card = dict(s['gcode_settings'].get('card', {}))
        card['engraving_mode'] = mode
        s['gcode_settings']['card'] = card
        placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, f'z_{mode}.gcode')
            ge.generate_gcode_from_placed(placed, 'card', W, H, path, 0, s, zones=zones)
            body = open(path, encoding='utf-8').read()
        assert body.strip(), f"{mode}: empty G-code"
        assert 'G1' in body, f"{mode}: no cutting moves emitted"


# --------------------------------------------------------------------------
# Settings plumbing
# --------------------------------------------------------------------------

def test_zone_keys_exist_in_defaults():
    for key in ('zone_labels_enabled', 'zone_label_min_size', 'zone_label_max_size',
                'zone_gutter_mm', 'zone_border_mm', 'zone_label_font_mm',
                'zone_edge_margin_mm', 'zone_group_gap_mm'):
        assert key in config.DEFAULT_SETTINGS, f"{key} missing from DEFAULT_SETTINGS"


def test_zone_keys_round_trip_through_sizing_presets():
    for key in ('zone_labels_enabled', 'zone_label_min_size', 'zone_label_max_size'):
        assert key in config.SIZING_PRESET_KEYS, f"{key} not captured by sizing presets"
    s = dict(config.DEFAULT_SETTINGS)
    s['zone_labels_enabled'] = True
    s['zone_label_max_size'] = 9.5
    preset = config.settings_to_sizing_preset(s)
    assert preset['zone_labels_enabled'] is True
    assert preset['zone_label_max_size'] == 9.5


def test_defaults_are_off():
    """Zones cost sheet area, so they must never turn themselves on."""
    assert config.DEFAULT_SETTINGS['zone_labels_enabled'] is False


# --------------------------------------------------------------------------
# Preview window — what the user actually looks at before cutting
# --------------------------------------------------------------------------

def _probe_preview(placements, w_mm, h_mm, polygon, zones):
    """Build a NestingPreviewWindow and report what landed on its canvas.

    The window calls wait_window() in __init__, so the inspection has to be
    scheduled on the parent beforehand and run inside that event loop.
    """
    import tkinter as tk
    from i18n import init_translation
    init_translation('en')
    from ui_dialogs import NestingPreviewWindow

    root = tk.Tk()
    root.geometry('900x700')
    out = {}

    def go():
        found = [x for x in root.winfo_children()
                 if isinstance(x, NestingPreviewWindow)]
        if not found:
            out['err'] = 'no window'
            return
        win = found[0]
        win.update_idletasks()
        win._draw()
        win.update_idletasks()
        cv = win._canvas
        kinds = {}
        for item in cv.find_all():
            k = cv.type(item)
            kinds[k] = kinds.get(k, 0) + 1
        out['kinds'] = kinds
        out['texts'] = [cv.itemcget(i, 'text') for i in cv.find_all()
                        if cv.type(i) == 'text']
        win.destroy()

    root.after(400, go)
    NestingPreviewWindow(root, placements, w_mm, h_mm, polygon=polygon, zones=zones)
    root.destroy()
    return out


def test_preview_draws_polygon_groups():
    try:
        import tkinter as tk
        tk.Tk().destroy()
    except Exception as e:
        print(f"    (skipped — no display: {e})")
        return
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    got = _probe_preview({'felt': placed}, 160, 120, SCRAP, zones)
    assert 'err' not in got, got.get('err')
    # The scrap outline is the one polygon; each group is a rectangle.
    assert got['kinds'].get('polygon', 0) >= 1, \
        f"scrap outline missing from preview: {got['kinds']}"
    assert got['kinds'].get('rectangle', 0) >= len(zones), \
        f"group rectangles missing from preview: {got['kinds']}"
    for z in zones:
        assert z['label'] in got['texts'], f"group label {z['label']} not drawn"


def test_preview_draws_rect_zone_blocks():
    try:
        import tkinter as tk
        tk.Tk().destroy()
    except Exception as e:
        print(f"    (skipped — no display: {e})")
        return
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
    got = _probe_preview({'card': placed}, W, H, None, zones)
    assert 'err' not in got, got.get('err')
    # one rectangle for the sheet, plus one per zone block
    assert got['kinds'].get('rectangle', 0) >= len(zones) + 1, \
        f"zone blocks missing from preview: {got['kinds']}"
    for z in zones:
        assert z['label'] in got['texts'], f"zone label {z['label']} not drawn"


def test_preview_shows_nothing_extra_when_zones_off():
    try:
        import tkinter as tk
        tk.Tk().destroy()
    except Exception as e:
        print(f"    (skipped — no display: {e})")
        return
    s = base_settings(enabled=False)
    placed, zones = se.nest_pads_with_zones(PADS, 'card', W, H, s)
    got = _probe_preview({'card': placed}, W, H, None, zones)
    assert 'err' not in got, got.get('err')
    assert got['kinds'].get('rectangle', 0) == 1, \
        f"expected only the sheet rectangle, got {got['kinds']}"


if __name__ == '__main__':
    print("Labeled zones")
    print("=" * 60)
    for name, fn in sorted(list(globals().items())):
        if name.startswith('test_') and callable(fn):
            check(name[5:].replace('_', ' '), fn)
    print("=" * 60)
    passed, total = sum(results), len(results)
    print(f"{passed}/{total} passed")
    sys.exit(0 if passed == total else 1)
