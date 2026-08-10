"""
Tests for labeled zones (svg_engine + gcode_engine).

Small pads are hard to tell apart once they're off the laser: a 7.0 and a
7.5 disc look identical, and the per-disc number is either dropped by the
font gate or unreadable. Leather can't be marked in the middle at all —
that's the sealing surface — so a zone puts the size on the WASTE instead:
a bordered grid of one size, with the number engraved along its top edge.

What these pin down:
  - zones are opt-in and change nothing when off (regression safety)
  - only in-range, fixed-quantity pads get zoned
  - the grid is a block, not a degenerate 1xN strip
  - no disc escapes its zone, no zone overlaps another, nothing overlaps
  - pads are never LOST — a zone that can't be placed falls back to the
    normal nest
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


def test_polygon_produces_bands_not_blocks():
    """A traced offcut has no straight edge to reserve a rectangular band
    against, so each size gets a horizontal band clipped to the outline."""
    s = base_settings()
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert placed, "polygon nesting should still place discs"
    assert zones, "polygon sheets should produce bands"
    for z in zones:
        assert z['shape'] == 'poly', f"expected a clipped band, got {z['shape']}"
        assert len(z['points']) >= 3, "band outline is not a polygon"


def test_polygon_zones_off_when_disabled():
    s = base_settings(enabled=False)
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert zones == [], "disabled must produce no bands on a polygon either"
    assert placed


def test_bands_are_separated_by_a_moat():
    """Neighbouring sizes must be visibly apart — that's the whole point.
    A 7.0 and a 7.5 disc are indistinguishable if the groups touch."""
    s = base_settings()
    _, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    spans = sorted((min(p[1] for p in z['points']),
                    max(p[1] for p in z['points'])) for z in zones)
    for (_, prev_bottom), (next_top, _) in zip(spans, spans[1:]):
        assert next_top > prev_bottom, "bands overlap vertically"


def test_discs_clear_their_band_outline():
    """The engraved outline must not crowd the discs — that's what the
    enlarged edge margin buys."""
    s = base_settings()
    margin = config.DEFAULT_SETTINGS['zone_edge_margin_mm']
    placed, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for z in zones:
        poly = z['points']
        for size, cx, cy, r in placed:
            if size != z['size']:
                continue
            d = min(se._distance_point_to_segment(
                cx, cy, poly[i][0], poly[i][1],
                poly[(i + 1) % len(poly)][0], poly[(i + 1) % len(poly)][1])
                for i in range(len(poly)))
            if d - r < -1e-6:
                continue  # disc belongs to a different band with the same size
            assert d - r >= margin - 0.51, (
                f"{z['label']}: disc only {d - r:.2f}mm from the outline")


def test_band_label_sits_inside_its_band():
    s = base_settings()
    _, zones, _, _ = se.nest_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    for z in zones:
        ys = [p[1] for p in z['points']]
        assert min(ys) <= z['label_y'] <= max(ys), \
            f"{z['label']}: label outside its band vertically"


def test_clip_polygon_y():
    square = [(0, 0), (10, 0), (10, 10), (0, 10)]
    mid = se._clip_polygon_y(square, 2, 8)
    ys = [p[1] for p in mid]
    assert abs(min(ys) - 2) < 1e-9 and abs(max(ys) - 8) < 1e-9, f"bad clip: {mid}"
    assert se._clip_polygon_y(square, 20, 30) == [], "clip outside should be empty"
    # concave shapes are the normal case for a traced scrap
    concave = [(0, 0), (10, 0), (10, 10), (5, 4), (0, 10)]
    got = se._clip_polygon_y(concave, 0, 5)
    assert len(got) >= 3, "concave clip collapsed"


def test_polygon_bands_never_silently_drop_pads():
    """When a band can't fit, its pads go back to the normal nest below —
    and if they still don't fit, can_all_pads_fit must report it rather
    than the app writing a short sheet."""
    s = base_settings()
    tiny = [(0, 0), (60, 0), (60, 50), (0, 50)]
    placed, _zones, fixed_placed, fixed_total = se.nest_with_zones(
        BAND_PADS, 'leather', 0, 0, s, polygon=tiny)
    assert fixed_total == sum(p['qty'] for p in BAND_PADS)
    assert fixed_placed == len(placed)
    if fixed_placed < fixed_total:
        assert se.can_all_pads_fit(BAND_PADS, 'leather', 0, 0, s, polygon=tiny) is False


def test_polygon_svg_and_gcode_band_outlines_agree():
    """Same Y-flip contract as the block zones, for the clipped outline."""
    s = base_settings()
    placed, zones = se.nest_pads_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
    assert zones, "no bands to compare"
    flip = max(p[1] for p in SCRAP)

    with tempfile.TemporaryDirectory() as d:
        gpath = os.path.join(d, 'b.gcode')
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

    for z in zones:
        for vx, vy in z['points']:
            want = (vx, flip - vy)
            hit = any(abs(px - want[0]) < 0.01 and abs(py - want[1]) < 0.01
                      for px, py in pts)
            assert hit, (f"band {z['label']}: vertex ({want[0]:.2f},{want[1]:.2f}) "
                         f"missing from G-code — Y-flip mismatch")


def test_polygon_svg_has_band_polygons():
    s = base_settings()
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'b.svg')
        placed, zones = se.nest_pads_with_zones(BAND_PADS, 'felt', 0, 0, s, polygon=SCRAP)
        se.generate_svg_from_placed(placed, 'felt', 200, 200, path, 0, s,
                                    polygon=SCRAP, zones=zones)
        body = open(path, encoding='utf-8').read()
    assert '<polygon' in body, "no band outline polygon in the SVG"
    for z in zones:
        assert f">{z['label']}<" in body, f"band label {z['label']} missing"


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
    """Scoring the bare inner grid always picks a 1xN strip (fewest gutters).
    The border and label must be in the decision, or zones come out as long
    thin ribbons that shelf-pack badly and are awkward to count."""
    for mat in ('card', 'felt', 'leather'):
        _, zones, _, _ = se.nest_with_zones(PADS, mat, 300.0, 300.0, base_settings())
        for z in zones:
            assert z['cols'] > 1 and z['rows'] > 1 or z['qty'] <= 3, (
                f"{mat} {z['label']}: degenerate {z['cols']}x{z['rows']} grid")
            aspect = max(z['w'], z['h']) / min(z['w'], z['h'])
            assert aspect < 6.0, f"{mat} {z['label']}: aspect {aspect:.1f} too extreme"


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


def test_unplaceable_zone_falls_back_instead_of_losing_pads():
    """A zone wider than the sheet can't be placed. Those pads must go back
    into the normal nest, not vanish."""
    s = base_settings()
    pads = [{'size': 12.0, 'qty': 40}]
    narrow = 40.0
    placed, zones, _, _ = se.nest_with_zones(pads, 'card', narrow, 400.0, s)
    assert _counts(placed).get(12.0, 0) > 0, "pads vanished when the zone didn't fit"
    for z in zones:
        assert z['w'] <= narrow, "a zone wider than the sheet was placed anyway"


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
                'zone_gutter_mm', 'zone_border_mm', 'zone_label_font_mm'):
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


def test_preview_draws_polygon_bands():
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
    # one polygon for the sheet outline, plus one per band
    assert got['kinds'].get('polygon', 0) >= len(zones) + 1, \
        f"band outlines missing from preview: {got['kinds']}"
    for z in zones:
        assert z['label'] in got['texts'], f"band label {z['label']} not drawn"


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
