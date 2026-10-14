"""Tonehole crescent: geometry, SVG path, STL solid, and the Tooling section.

The crescent covers a chosen percentage of a tonehole's open area. The
outer edge is the tonehole wall, the inner edge an arc whose curvature a
slider sets (75 = the tonehole's own curve). Three outputs share one
geometry: the live preview outline, the SVG cut path (true arcs), and the
extruded STL. These tests pin that they agree:

  - the covered area is the requested percentage, on every curve setting
  - slider 75 is the standard crescent; flatter is wider, rounder wraps
    further, and the horns never close into a ring
  - the SVG arc flags draw the same shape (checked with an independent
    SVG-spec arc converter, not the engine's own outline)
  - the STL is watertight, wound outward, and its volume is area x thickness
  - out-of-range input is refused
  - the Tooling section previews, validates and persists its fields
    (the UI cases self-skip without a display)

Generate SVG / Generate STL through the real buttons are covered by
test_tooling_e2e.
"""
import math
import os
import struct
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-crescent-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

from config import DEFAULT_SETTINGS  # noqa: E402
from svg_engine import (  # noqa: E402
    CRESCENT_CURVE_STANDARD, CRESCENT_FLATTEST, CRESCENT_MAX_PERCENT,
    CRESCENT_SVG_MARGIN_MM, _circle_overlap_area, crescent_inner_curvature,
    generate_tonehole_crescent_stl, generate_tonehole_crescent_svg,
    tonehole_crescent_bounds, tonehole_crescent_geometry, tonehole_crescent_outline,
    tonehole_crescent_solid, tonehole_crescent_svg_path,
)

results = []

DIAMETERS = (8.0, 25.0, 42.5, 60.0)
PERCENTS = (0.5, 2.0, 5.0, 10.0, 20.0, 35.0, 50.0, 75.0)
CURVES = (0, 10, 23, 50, 75, 88, 100)


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


def grid():
    for d in DIAMETERS:
        for p in PERCENTS:
            for c in CURVES:
                yield d, p, c, tonehole_crescent_geometry(d, p, c)


def polygon_area(pts):
    s = 0.0
    for (x1, y1), (x2, y2) in zip(pts, pts[1:] + pts[:1]):
        s += x1 * y2 - x2 * y1
    return abs(s) / 2.0


def closed_outline(geom, segments):
    outer, inner = tonehole_crescent_outline(geom, segments)
    return outer + inner[-2:0:-1]


# ---------------------------------------------------------------- geometry

def t_area_exact():
    for d, p, c, g in grid():
        R = g['radius']
        covered = math.pi * R * R - _circle_overlap_area(R, g['inner_radius'], g['offset'])
        want = p / 100.0 * math.pi * R * R
        assert abs(covered / want - 1) < 1e-9, (d, p, c, covered / want)


def t_outline_area():
    # The sampled outline is what the preview and STL use; it must cover the
    # same area the solver aimed for, to well under a percent.
    for d, p, c, g in grid():
        got = polygon_area(closed_outline(g, 512))
        want = p / 100.0 * math.pi * g['radius'] ** 2
        assert abs(got / want - 1) < 1e-3, (d, p, c, got / want)


def t_standard_is_tonehole_curve():
    for p in PERCENTS:
        g = tonehole_crescent_geometry(42.5, p, CRESCENT_CURVE_STANDARD)
        assert abs(g['inner_radius'] - g['radius']) < 1e-12, (p, g['inner_radius'])
    assert abs(crescent_inner_curvature(0, 0.1) - CRESCENT_FLATTEST) < 1e-12
    assert crescent_inner_curvature(-50, 0.1) == crescent_inner_curvature(0, 0.1)
    assert crescent_inner_curvature(500, 0.1) == crescent_inner_curvature(100, 0.1)


def t_curve_is_monotonic():
    # Flatter = bigger inner radius, a wider crescent, horns higher up the
    # wall. Rounder = the reverse. Strict at every step of the slider.
    for p in PERCENTS:
        prev = None
        for c in range(0, 101, 5):
            g = tonehole_crescent_geometry(30.0, p, c)
            if prev:
                assert g['inner_radius'] < prev['inner_radius'], (p, c)
                assert g['width'] < prev['width'], (p, c, g['width'], prev['width'])
                assert g['tip_y'] > prev['tip_y'], (p, c)
            prev = g


def t_horns_never_close():
    # The inner circle reaching the size of the opening it leaves is where the
    # crescent becomes a closed ring. The roundest setting stays clear of it.
    for d, p, c, g in grid():
        limit = g['radius'] * math.sqrt(1 - p / 100.0)
        assert g['inner_radius'] > limit * (1 + 1e-6), (d, p, c)
        assert g['tip_x'] > 0 and g['tip_y'] < g['radius'], (d, p, c)


def t_shape_on_mouthpiece_side():
    for d, p, c, g in grid():
        R, Rc, off = g['radius'], g['inner_radius'], g['offset']
        outer, inner = tonehole_crescent_outline(g, 64)
        assert outer[0] == inner[0] and outer[-1] == inner[-1], "edges must share the horns"
        for x, y in outer:
            assert abs(math.hypot(x, y) - R) < 1e-9 and y <= g['tip_y'] + 1e-9
        for x, y in inner:
            assert abs(math.hypot(x, y - off) - Rc) < 1e-9
            assert math.hypot(x, y) <= R + 1e-9, "inner edge must stay inside the tonehole"
        assert min(y for _x, y in outer) < -R + 1e-3, "the crescent's back must be the top of the tonehole"


def t_bounds():
    for d, p, c, g in grid():
        x0, y0, x1, y1 = tonehole_crescent_bounds(g)
        pts = closed_outline(g, 512)
        xs, ys = [x for x, _y in pts], [y for _x, y in pts]
        tol = 1e-3 * g['radius']
        assert x0 - 1e-9 <= min(xs) < x0 + tol and x1 - tol < max(xs) <= x1 + 1e-9, (d, p, c)
        assert y0 - 1e-9 <= min(ys) < y0 + tol and y1 - tol < max(ys) <= y1 + 1e-9, (d, p, c)


def t_rejects_bad_input():
    for d, p in ((0, 10), (-5, 10), (250, 10), (40, 0), (40, -1), (40, CRESCENT_MAX_PERCENT + 0.1),
                 ("abc", 10), (40, "")):
        try:
            tonehole_crescent_geometry(d, p, 75)
        except ValueError:
            continue
        raise AssertionError(f"accepted d={d!r} p={p!r}")
    g = tonehole_crescent_geometry(40, 10, 75)
    for t in (0, -1, 51):
        try:
            tonehole_crescent_solid(g, t)
        except ValueError:
            continue
        raise AssertionError(f"accepted thickness {t}")


# ---------------------------------------------------------------- SVG path

def _arc_points(x1, y1, rx, ry, large, sweep, x2, y2, n=400):
    """Sample an SVG elliptical arc (no rotation) via the spec's
    endpoint-to-centre conversion (SVG 1.1 appendix F.6.5)."""
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    lam = (dx * dx) / (rx * rx) + (dy * dy) / (ry * ry)
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx * rx * ry * ry - rx * rx * dy * dy - ry * ry * dx * dx
    den = rx * rx * dy * dy + ry * ry * dx * dx
    co = math.sqrt(max(0.0, num / den)) * (-1 if large == sweep else 1)
    cxp, cyp = co * rx * dy / ry, -co * ry * dx / rx
    cx, cy = cxp + (x1 + x2) / 2, cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a
    t1 = ang(1, 0, (dx - cxp) / rx, (dy - cyp) / ry)
    dt = ang((dx - cxp) / rx, (dy - cyp) / ry, (-dx - cxp) / rx, (-dy - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    return [(cx + rx * math.cos(t1 + dt * i / n), cy + ry * math.sin(t1 + dt * i / n)) for i in range(n + 1)]


def _path_points(d):
    tok = d.replace(",", " ").split()
    assert tok[0] == "M" and tok[3] == "A" and tok[11] == "A" and tok[-1] == "Z", d
    x, y = float(tok[1]), float(tok[2])
    pts = []
    for i in (4, 12):
        rx, ry, _rot, large, sweep, x2, y2 = (float(v) for v in tok[i:i + 7])
        pts += _arc_points(x, y, rx, ry, int(large), int(sweep), x2, y2)[:-1]
        x, y = x2, y2
    return pts


def t_svg_path_matches():
    flags = Counter()
    for d, p, c, g in grid():
        cx, cy = 100.0, 100.0
        path = tonehole_crescent_svg_path(g, cx, cy)
        tok = path.replace(",", " ").split()
        flags[("outer", tok[7])] += 1
        flags[("inner", tok[15])] += 1
        pts = _path_points(path)
        want = p / 100.0 * math.pi * g['radius'] ** 2
        got = polygon_area(pts)
        assert abs(got / want - 1) < 2e-3, (d, p, c, got / want, path)
        assert abs(min(y for _x, y in pts) - (cy - g['radius'])) < 1e-3 * g['radius'], \
            "the path's crescent must be on the mouthpiece (top) side"
    # Both large-arc settings must have been exercised on both arcs, or this
    # test proved nothing about the flag logic.
    for arc in ("outer", "inner"):
        assert flags[(arc, "0")] and flags[(arc, "1")], flags


def t_svg_file():
    out = tempfile.mkdtemp(prefix="ssc-crescent-svg-")
    for compat in (False, True):
        settings = dict(DEFAULT_SETTINGS, compatibility_mode=compat)
        path = os.path.join(out, f"c_{compat}.svg")
        g = generate_tonehole_crescent_svg(42.5, 12.0, 60, path, settings)
        root = ET.parse(path).getroot()
        ns = {"s": "http://www.w3.org/2000/svg"}
        x0, y0, x1, y1 = tonehole_crescent_bounds(g)
        m = CRESCENT_SVG_MARGIN_MM
        assert abs(float(root.get("width").rstrip("m")) - (x1 - x0 + 2 * m)) < 2e-3
        assert abs(float(root.get("height").rstrip("m")) - (y1 - y0 + 2 * m)) < 2e-3
        vb = [float(v) for v in root.get("viewBox").split()]
        assert vb[0] == 0 and vb[1] == 0 and abs(vb[2] - float(root.get("width")[:-2])) < 1e-9
        paths = root.findall(".//s:path", ns)
        assert len(paths) == 1 and len(root.findall(".//s:circle", ns)) == 0, "exactly one cut path"
        assert paths[0].get("stroke") == DEFAULT_SETTINGS["layer_colors"]["die_outer_cut"]
        assert paths[0].get("fill") == "none"
        assert (paths[0].get("stroke-width") == "0.1") == compat
        pts = _path_points(paths[0].get("d"))
        assert min(x for x, _y in pts) > m - 1e-3 and max(x for x, _y in pts) < vb[2] - m + 1e-3
        assert min(y for _x, y in pts) > m - 1e-3 and max(y for _x, y in pts) < vb[3] - m + 1e-3


# ---------------------------------------------------------------- STL

def read_stl(path):
    data = open(path, "rb").read()
    assert not data[:5].lower().startswith(b"solid"), "binary STL header must not start with 'solid'"
    n = struct.unpack("<I", data[80:84])[0]
    assert len(data) == 84 + 50 * n, (len(data), n)
    tris = []
    for i in range(n):
        v = struct.unpack("<12fH", data[84 + 50 * i: 84 + 50 * (i + 1)])
        tris.append(((v[0], v[1], v[2]), (v[3], v[4], v[5]), (v[6], v[7], v[8]), (v[9], v[10], v[11])))
    return tris


def _volume(tris):
    return sum((a[0] * (b[1] * c[2] - b[2] * c[1]) - a[1] * (b[0] * c[2] - b[2] * c[0])
                + a[2] * (b[0] * c[1] - b[1] * c[0])) / 6.0 for a, b, c in tris)


def _watertight(tris):
    edges = Counter()
    for t in tris:
        for i in range(3):
            edges[(t[i], t[(i + 1) % 3])] += 1
    return all(k == 1 and edges.get((v, u)) == 1 for (u, v), k in edges.items())


def t_solid_watertight_and_sized():
    for d, p, c, g in grid():
        tris = tonehole_crescent_solid(g, 2.5, 48)
        assert _watertight(tris), (d, p, c)
        cap = [t for t in tris if all(v[2] == 2.5 for v in t)]
        signed = [((b[0] - a[0]) * (cc[1] - a[1]) - (cc[0] - a[0]) * (b[1] - a[1])) / 2 for a, b, cc in cap]
        assert min(signed) > 0, f"cap triangle wound the wrong way or folded over: {(d, p, c)}"
        area = sum(signed)
        want = p / 100.0 * math.pi * g['radius'] ** 2
        assert abs(area / want - 1) < 0.01, (d, p, c, area / want)
        assert abs(_volume(tris) / (2.5 * area) - 1) < 1e-9, "volume must be area x thickness, outward"


def t_stl_file():
    out = tempfile.mkdtemp(prefix="ssc-crescent-stl-")
    path = os.path.join(out, "c.stl")
    g = generate_tonehole_crescent_stl(42.5, 12.0, 60, 3.0, path)
    facets = read_stl(path)
    tris = [f[1:] for f in facets]
    assert _watertight(tris), "STL as written (float32) must still be watertight"
    xs = [v[0] for t in tris for v in t]
    ys = [v[1] for t in tris for v in t]
    zs = [v[2] for t in tris for v in t]
    x0, y0, x1, y1 = tonehole_crescent_bounds(g)
    assert abs(min(zs)) < 1e-6 and abs(max(zs) - 3.0) < 1e-5, "flat on the bed, 3 mm tall"
    assert abs(max(ys) - g['radius']) < 1e-4, "mouthpiece side is +y"
    assert abs(min(ys) + g['tip_y']) < 1e-4
    assert abs(min(xs) - x0) < 0.01 * g['radius'] and abs(max(xs) - x1) < 0.01 * g['radius']
    for n, a, b, c in facets:
        ux, uy, uz = b[0] - a[0], b[1] - a[1], b[2] - a[2]
        vx, vy, vz = c[0] - a[0], c[1] - a[1], c[2] - a[2]
        cx, cy, cz = uy * vz - uz * vy, uz * vx - ux * vz, ux * vy - uy * vx
        ln = math.sqrt(cx * cx + cy * cy + cz * cz)
        assert ln > 0, "degenerate facet"
        assert (n[0] * cx + n[1] * cy + n[2] * cz) / ln > 0.99, "stored normal disagrees with winding"
    assert _volume(tris) > 0


# ---------------------------------------------------------------- UI

def ui_checks():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  skip  UI cases: no display available ({e})")
        return
    root.withdraw()
    import main as main_mod
    app = main_mod.PadSVGGeneratorApp(root)
    app.notebook.select(app.tooling_tab_frame)
    root.update()

    def find_button(widget, text):
        for child in widget.winfo_children():
            if isinstance(child, tk.Button) and child.cget("text") == text:
                return child
            hit = find_button(child, text)
            if hit:
                return hit
        return None

    def polygons():
        return [i for i in app.crescent_canvas.find_all() if app.crescent_canvas.type(i) == "polygon"]

    def section_registered():
        assert "tonehole_crescent" in app._tooling_sections
        app._show_tooling_section("tonehole_crescent")
        root.update()
        # The root is withdrawn, so nothing maps; packed is what "shown" means here.
        assert app._tooling_sections["tonehole_crescent"].winfo_manager() == "pack"
    check("UI: Tonehole Crescent section opens from its selector button", section_registered)

    def preview_follows_input():
        app.crescent_tonehole_var.set("38")
        app.crescent_percent_var.set("12")
        root.update()
        assert len(polygons()) == 1, "one crescent drawn"
        text = app.crescent_readout_var.get()
        assert "88%" in text and "mm" in text, text
        before = app.crescent_canvas.coords(polygons()[0])
        app.crescent_curve_var.set(10)
        app._update_crescent_preview()
        assert app.crescent_canvas.coords(polygons()[0]) != before, "slider must reshape the preview"
    check("UI: preview and readout follow the fields and the slider", preview_follows_input)

    def match_button():
        app.crescent_curve_var.set(20)
        btn = find_button(app._tooling_sections["tonehole_crescent"], "Match tonehole curve")
        assert btn is not None, "Match tonehole curve button"
        btn.invoke()
        assert app.crescent_curve_var.get() == CRESCENT_CURVE_STANDARD
        r = tonehole_crescent_geometry(38, 12, app.crescent_curve_var.get())['inner_radius']
        assert f"{r:.1f}" in app.crescent_readout_var.get()
    check("UI: Match tonehole curve snaps the slider to the standard crescent", match_button)

    def bad_input_shows_message():
        for d, p in (("", "10"), ("abc", "10"), ("40", "0"), ("40", "90"), ("0", "10")):
            app.crescent_tonehole_var.set(d)
            app.crescent_percent_var.set(p)
            root.update()
            assert not polygons(), f"nothing drawn for {d!r}/{p!r}"
            assert app.crescent_readout_var.get(), f"no message for {d!r}/{p!r}"
        app.crescent_tonehole_var.set("40")
        app.crescent_percent_var.set("10")
        root.update()
        assert polygons()
    check("UI: invalid fields clear the preview and say why", bad_input_shows_message)

    def persisted():
        app.crescent_tonehole_var.set("44.5")
        app.crescent_percent_var.set("7.5")
        app.crescent_curve_var.set(62.3)
        app.crescent_thickness_var.set("1.6")
        app._update_tooling_settings()
        t = app.settings["tooling_settings"]
        assert (t["crescent_tonehole_d"], t["crescent_percent"], t["crescent_curve"], t["crescent_thickness"]) \
            == ("44.5", "7.5", 62.3, "1.6"), t
        for key in ("crescent_tonehole_d", "crescent_percent", "crescent_curve", "crescent_thickness"):
            assert key in DEFAULT_SETTINGS["tooling_settings"], key
    check("UI: crescent fields persist into tooling_settings (and have defaults)", persisted)

    try:
        root.destroy()
    except tk.TclError:
        pass


def main():
    print("Tonehole crescent")
    print("=" * 60)
    check("covered area equals the requested percentage (analytic)", t_area_exact)
    check("sampled outline covers the same area", t_outline_area)
    check("slider 75 = inner edge matches the tonehole's curve", t_standard_is_tonehole_curve)
    check("flatter is wider, rounder wraps further, strictly", t_curve_is_monotonic)
    check("the horns never close into a ring", t_horns_never_close)
    check("crescent sits on the mouthpiece side, edges on their circles", t_shape_on_mouthpiece_side)
    check("bounds are tight around the outline", t_bounds)
    check("out-of-range input is refused", t_rejects_bad_input)
    check("SVG arc flags draw the same crescent (independent arc math)", t_svg_path_matches)
    check("SVG file: one cut path, sized to the crescent, both output modes", t_svg_file)
    check("STL solid: watertight, wound outward, volume = area x thickness", t_solid_watertight_and_sized)
    check("STL file: binary, float32-watertight, flat on the bed, normals agree", t_stl_file)
    ui_checks()
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
