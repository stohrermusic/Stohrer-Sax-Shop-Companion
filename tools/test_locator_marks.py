"""Leather locator marks (2026-10-06).

Engraved guides on leather pads showing where the felt sits, for centering
the felt on a pad that has no center hole. Feature-requested; opt-in; lives
behind "Lesser-used settings" in Sizing Rules.

Covers:
  - off (default, or key absent) changes nothing in SVG or G-code
  - leather only, size range inclusive
  - line geometry: inner end on the felt edge, outer end 1 mm inside the cut
    (plain pad) or the dart valleys (darted pad); too-short lines dropped
  - styles (lines / circle / both) and dashes, dashes anchored at the felt edge
  - SVG and G-code draw the same points (Y-flip)
  - the G-code layer runs at the leather engraving settings, before holes and
    cuts, in both grouping modes, and says so in the file
  - preset schema: keys captured, backfilled for older presets
  - the layer color exists and is a LightBurn color
  - GUI (self-skips without a display): the lesser-used section collapses by
    default and opens itself when a setting inside is active; the form round-
    trips the keys; Layer Colors lists the layer; the pad preview draws marks
"""
import copy
import math
import os
import re
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config  # noqa: E402
import svg_engine as se  # noqa: E402
import gcode_engine as ge  # noqa: E402
# ui_dialogs evaluates _() at class-definition time, so install gettext before
# any GUI case imports it (same as test_tooltips).
from i18n import init_translation  # noqa: E402
init_translation("en")

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


def base_settings(**over):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    s.update(over)
    return s


def on(**over):
    return base_settings(locator_marks_enabled=True, **over)


# Real sizes: darted (small) and plain (above dart_threshold 18) leather.
PLACED = [(7.0, 20.0, 20.0, 8.5), (10.0, 45.0, 20.0, 10.0), (16.0, 80.0, 25.0, 13.0),
          (20.0, 120.0, 30.0, 15.0)]


def _placed(settings, material='leather'):
    return [(p, cx, cy, se.get_disc_diameter(p, material, settings) / 2) for p, cx, cy, _ in PLACED]


def _render(settings, material='leather', hole=0):
    with tempfile.TemporaryDirectory() as d:
        sp, gp = os.path.join(d, 'a.svg'), os.path.join(d, 'a.gcode')
        se.generate_svg_from_placed(_placed(settings, material), material, 200, 100, sp, hole, settings)
        ge.generate_gcode_from_placed(_placed(settings, material), material, 200, 100, gp, hole, settings)
        return open(sp, encoding='utf-8').read(), open(gp, encoding='utf-8').read()


def _dart_inner_r(pad, settings):
    sizing = config.get_sizing_for_size(pad, settings)
    dart = config.get_dart_settings_for_size(pad, settings)
    felt_r = (pad - sizing["felt_offset"]) / 2
    return felt_r + se.get_felt_thickness_mm(settings, sizing) + dart.get("overwrap", 0.5)


# ---------------------------------------------------------------- off ----

def test_off_by_default():
    assert config.DEFAULT_SETTINGS["locator_marks_enabled"] is False
    assert config.DEFAULT_SETTINGS["locator_marks_style"] == "lines"


def test_off_or_absent_changes_nothing():
    absent = base_settings()
    for k in list(absent):
        if k.startswith("locator_marks_"):
            del absent[k]
    off = base_settings(locator_marks_enabled=False)
    assert _render(absent) == _render(off), "key absent vs off differ"
    svg_off, g_off = _render(off)
    assert "#00E0E0" not in svg_off and "C06" not in g_off
    assert se.locator_mark_geometry(10.0, 10.0, off) is None


# ------------------------------------------------------ who gets marks ----

def test_leather_only_and_range_inclusive():
    s = on(locator_marks_min_size=7.0, locator_marks_max_size=16.0)
    for pad, want in ((6.9, False), (7.0, True), (12.0, True), (16.0, True), (16.1, False)):
        got = se.locator_mark_geometry(pad, se.get_disc_diameter(pad, 'leather', s) / 2, s) is not None
        assert got == want, f"pad {pad}: marks={got}, wanted {want}"
    for mat in ('felt', 'card', 'exact_size'):
        assert not se.locator_marks_apply(10.0, mat, s)
        assert _render(s, mat) == _render(base_settings(), mat), f"{mat} output changed"


# -------------------------------------------------------- geometry -------

def test_lines_run_from_felt_edge_to_inside_the_cut():
    s = on(darts_enabled=False)
    pad = 10.0
    r = se.get_disc_diameter(pad, 'leather', s) / 2
    felt_r = se.get_disc_diameter(pad, 'felt', s) / 2
    g = se.locator_mark_geometry(pad, r, s)
    assert len(g["segments"]) == 4 and not g["circles"]
    dirs = set()
    for (x0, y0), (x1, y1) in g["segments"]:
        assert math.isclose(math.hypot(x0, y0), felt_r, abs_tol=1e-9), "inner end not on felt edge"
        assert math.isclose(math.hypot(x1, y1), r - se.LOCATOR_EDGE_MARGIN_MM, abs_tol=1e-9), "outer end"
        dirs.add((round(math.copysign(1, x1)) if abs(x1) > 1e-9 else 0,
                  round(math.copysign(1, y1)) if abs(y1) > 1e-9 else 0))
    assert dirs == {(1, 0), (-1, 0), (0, 1), (0, -1)}, dirs


def test_darted_pad_lines_stop_inside_the_dart_valleys():
    s = on()
    pad = 10.0
    assert config.get_dart_settings_for_size(pad, s) is not None, "fixture must be a darted pad"
    r = se.get_disc_diameter(pad, 'leather', s) / 2
    inner_r = _dart_inner_r(pad, s)
    assert inner_r < r
    g = se.locator_mark_geometry(pad, r, s)
    for _, (x1, y1) in g["segments"]:
        assert math.isclose(math.hypot(x1, y1), inner_r - se.LOCATOR_EDGE_MARGIN_MM, abs_tol=1e-9)


def test_too_short_line_is_dropped_not_drawn_as_a_speck():
    # Thin felt + darts: valley is barely outside the felt edge.
    s = on(felt_thickness=1.0, felt_thickness_unit="mm", dart_overwrap=0.5)
    pad = 10.0
    r = se.get_disc_diameter(pad, 'leather', s) / 2
    inner_r = _dart_inner_r(pad, s)
    felt_r = se.get_disc_diameter(pad, 'felt', s) / 2
    assert inner_r - se.LOCATOR_EDGE_MARGIN_MM - felt_r < se.LOCATOR_MIN_LINE_MM, "fixture not short enough"
    assert se.locator_mark_geometry(pad, r, s) is None, "lines style: nothing to draw -> None"
    g = se.locator_mark_geometry(pad, r, on(**{k: s[k] for k in ("felt_thickness", "felt_thickness_unit",
                                                                   "dart_overwrap")},
                                              locator_marks_style="both"))
    assert g["circles"] == [felt_r] and not g["segments"], "both style: circle survives, lines dropped"


def test_styles():
    pad, s = 10.0, on()
    r = se.get_disc_diameter(pad, 'leather', s) / 2
    felt_r = se.get_disc_diameter(pad, 'felt', s) / 2
    g = se.locator_mark_geometry(pad, r, on(locator_marks_style="circle"))
    assert g["circles"] == [felt_r] and g["segments"] == []
    g = se.locator_mark_geometry(pad, r, on(locator_marks_style="both"))
    assert g["circles"] == [felt_r] and len(g["segments"]) == 4
    g = se.locator_mark_geometry(pad, r, on(locator_marks_style="nonsense"))
    assert not g["circles"] and len(g["segments"]) == 4, "unknown style falls back to lines"


def test_dashes_start_on_the_felt_edge():
    pad, s = 16.0, on(locator_marks_dashed=True, locator_marks_style="both")
    r = se.get_disc_diameter(pad, 'leather', s) / 2
    felt_r = se.get_disc_diameter(pad, 'felt', s) / 2
    g = se.locator_mark_geometry(pad, r, s)
    assert len(g["segments"]) > 4 and len(g["segments"]) % 4 == 0, len(g["segments"])
    outer_limit = _dart_inner_r(pad, s) - se.LOCATOR_EDGE_MARGIN_MM
    for ux, uy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
        spans = sorted(math.hypot(*a) for a, b in g["segments"]
                       if (abs(b[0]) > 1e-9 and math.copysign(1, b[0]) == ux and uy == 0)
                       or (abs(b[1]) > 1e-9 and math.copysign(1, b[1]) == uy and ux == 0))
        assert math.isclose(spans[0], felt_r, abs_tol=1e-9), "first dash must start at the felt edge"
    for a, b in g["segments"]:
        assert math.hypot(*b) <= outer_limit + 1e-9
        assert math.hypot(*b) - math.hypot(*a) <= se.LOCATOR_DASH_MM + 1e-9
    arcs = se.locator_circle_points(felt_r, True)
    assert len(arcs) >= 4
    for arc in arcs:
        for x, y in arc:
            assert math.isclose(math.hypot(x, y), felt_r, abs_tol=1e-9)


# ------------------------------------------------------- SVG <-> G-code --

def _gcode_layer_points(body, layer):
    pts, inside = [], False
    for ln in body.splitlines():
        if ln.startswith("; Layer "):
            inside = ln.strip() == f"; Layer {layer}"
            continue
        if inside and (ln.startswith("G0") or ln.startswith("G1")):
            xm = re.search(r'X(-?\d+(?:\.\d+)?)', ln)
            ym = re.search(r'Y(-?\d+(?:\.\d+)?)', ln)
            if xm and ym:
                pts.append((float(xm.group(1)), float(ym.group(1))))
    return pts


def test_svg_and_gcode_draw_the_same_marks():
    s = on(locator_marks_style="both")
    svg, g = _render(s)
    flip = 100.0
    lines = re.findall(r'<line[^>]*stroke="#00E0E0"[^>]*/>', svg)
    circles = re.findall(r'<circle[^>]*stroke="#00E0E0"[^>]*/>', svg)
    marked = [p for p in _placed(s) if se.locator_mark_geometry(p[0], p[3], s)]
    assert len(marked) == 3, marked
    assert len(lines) == 4 * len(marked) and len(circles) == len(marked)
    gpts = _gcode_layer_points(g, 'C06')
    assert gpts, "no C06 locator layer in the G-code"

    def near(x, y):
        return any(abs(px - x) < 0.002 and abs(py - y) < 0.002 for px, py in gpts)

    for ln in lines:
        x1, y1 = (float(re.search(r'x1="([0-9.]+)mm"', ln).group(1)),
                  float(re.search(r'y1="([0-9.]+)mm"', ln).group(1)))
        x2, y2 = (float(re.search(r'x2="([0-9.]+)mm"', ln).group(1)),
                  float(re.search(r'y2="([0-9.]+)mm"', ln).group(1)))
        assert near(x1, flip - y1) and near(x2, flip - y2), f"SVG line {ln} not in G-code"
    for c in circles:
        cx = float(re.search(r'cx="([0-9.]+)mm"', c).group(1))
        cy = float(re.search(r'cy="([0-9.]+)mm"', c).group(1))
        rr = float(re.search(r'\br="([0-9.]+)mm"', c).group(1))
        assert near(cx + rr, flip - cy), "circle start point missing from G-code"


def test_compatibility_mode_svg_is_unitless_for_marks():
    svg, _ = _render(on(locator_marks_style="both", compatibility_mode=True))
    marks = re.findall(r'<(?:line|circle)[^>]*stroke="#00E0E0"[^>]*/>', svg)
    assert marks and all('mm"' not in m for m in marks)


def test_gcode_layer_uses_engraving_settings_before_holes_and_cuts():
    for grouping in ("layer", "pad"):
        s = on(gcode_cut_grouping=grouping)
        s["gcode_settings"]["leather"]["engraving_mode"] = "line"
        s["gcode_settings"]["leather"]["engraving_speed"] = 1234
        s["gcode_settings"]["leather"]["engraving_power"] = 7
        _, g = _render(s, hole=3.0)
        lines = g.splitlines()
        idx = [i for i, ln in enumerate(lines) if ln.strip() == "; Layer C06"]
        assert idx, f"{grouping}: no locator layer"
        head = [ln for ln in lines[:idx[0]] if ln.startswith("; Cut @")][-1]
        assert "1234 mm/min, 7% power" in head, head
        assert "Locator marks: engraving settings" in g, "no note about where the power comes from"
        first_cut = min(i for i, ln in enumerate(lines) if ln.strip() in ("; Layer C03", "; Layer C02"))
        assert idx[0] < first_cut, f"{grouping}: marks must go down before any hole or cut"
        if grouping == "pad":
            assert len(idx) == 3, f"per-pad grouping: one locator layer per marked disc, got {len(idx)}"


# ------------------------------------------------------------ presets ----

def test_preset_schema_captures_and_backfills():
    keys = ("locator_marks_enabled", "locator_marks_min_size", "locator_marks_max_size",
            "locator_marks_style", "locator_marks_dashed")
    for k in keys:
        assert k in config.SIZING_PRESET_KEYS, k
    preset = config.settings_to_sizing_preset(on(locator_marks_style="both"))
    assert preset["locator_marks_enabled"] is True and preset["locator_marks_style"] == "both"
    old = {k: v for k, v in preset.items() if not k.startswith("locator_marks_")}
    fixed = config.normalize_sizing_preset(old)
    for k in keys:
        assert fixed[k] == config.DEFAULT_SETTINGS[k], f"{k} not backfilled"


def test_layer_color_is_a_lightburn_color():
    hex_val = config.DEFAULT_SETTINGS["layer_colors"]["leather_locator"]
    assert hex_val in {h for _, h in config.LIGHTBURN_COLORS}
    taken = {v for k, v in config.DEFAULT_SETTINGS["layer_colors"].items()
             if k.startswith("leather_") and k != "leather_locator"}
    assert hex_val not in taken, "locator layer must not share a color with another leather layer"


# ---------------------------------------------------------------- GUI ----

def _tk_root():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"        (no display: {e})")
        return None
    root.withdraw()
    return root


def _options(root, settings):
    from ui_dialogs import OptionsWindow

    class StubApp:
        def open_resonance_window(self):
            pass

    presets = {"Default": config.settings_to_sizing_preset(settings)}
    w = OptionsWindow(root, StubApp(), settings, lambda: None, lambda: None,
                      sizing_presets=presets, sizing_presets_save_callback=lambda: None)
    w.top.withdraw()
    return w


def test_gui_lesser_used_section_and_form_round_trip():
    root = _tk_root()
    if root is None:
        return
    try:
        from i18n import init_translation
        init_translation("en")
        w = _options(root, base_settings())
        assert not w.lesser_open, "collapsed by default when nothing inside is active"
        assert not w.lesser_body.winfo_manager()
        assert str(w.locator_min_entry.cget("state")) == "disabled"
        w._toggle_lesser_used()
        assert w.lesser_open and w.lesser_body.winfo_manager()
        w.locator_enabled_var.set(True)
        w._toggle_locator_fields()
        assert str(w.locator_min_entry.cget("state")) == "normal"
        w.locator_style_var.set("both")
        w.locator_dashed_var.set(True)
        snap = w._capture_form_to_dict()
        assert snap["locator_marks_enabled"] and snap["locator_marks_style"] == "both" and snap["locator_marks_dashed"]
        w2 = _options(root, base_settings())
        w2._apply_dict_to_form(snap)
        assert w2._capture_form_to_dict() == snap
        w.top.destroy()
        w2.top.destroy()

        w3 = _options(root, on())
        assert w3.lesser_open, "opens itself when locator marks are on"
        w3.top.destroy()
        w4 = _options(root, base_settings(compatibility_mode=True))
        assert w4.lesser_open, "opens itself when compatibility mode is on"
        w4.top.destroy()
    finally:
        root.destroy()


def test_gui_layer_colors_lists_the_locator_layer():
    root = _tk_root()
    if root is None:
        return
    try:
        from ui_dialogs import LayerColorWindow
        s = base_settings()
        w = LayerColorWindow(root, s, lambda: None)
        w.top.withdraw()
        assert "leather_locator" in w.color_vars
        assert w.color_vars["leather_locator"].get().startswith("06")
        w.top.destroy()
    finally:
        root.destroy()


def test_gui_pad_preview_draws_marks():
    root = _tk_root()
    if root is None:
        return
    try:
        from ui_dialogs import PadPreviewWindow
        for enabled, want in ((False, 0), (True, 4)):
            w = _options(root, base_settings(locator_marks_enabled=enabled, locator_marks_style="lines"))
            pv = PadPreviewWindow(w)
            pv.withdraw()
            pv._cancel_poll()
            pv.preview_size_var.set(10.0)
            pv.canvas.configure(width=400, height=400)
            pv.update_idletasks()
            pv._render()
            n = len(pv.canvas.find_withtag("locator"))
            assert n == want, f"enabled={enabled}: {n} locator items, wanted {want}"
            pv.destroy()
            w.top.destroy()
    finally:
        root.destroy()


if __name__ == '__main__':
    print("Leather locator marks")
    print("=" * 60)
    for name, fn in sorted(list(globals().items())):
        if name.startswith('test_') and callable(fn):
            check(name[5:].replace('_', ' '), fn)
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    sys.exit(0 if all(results) else 1)
