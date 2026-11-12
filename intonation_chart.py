"""The intonation chart: one geometry, three renderers.

`chart_primitives()` turns the data of `intonation.overlay(runs)` into a list of
drawing primitives in pixel coordinates (lines, dashed lines, polylines, text with
an anchor and an angle, markers, a background). That list is the chart's look in
one place: a title line above the plot, the legend above the plot in columns with
each run's colour and marker, every written note along the bottom with its label
turned 90 degrees, low notes on the left (or the right, a preference), cents up
the side with a light grid line every 10 cents and a solid zero line, a dashed
vertical line where the register changes, one line per run joining its notes'
medians with a marker at each and a thin bar for the spread of a note's takes. The
y axis is cents AS PLAYED and nothing else (Matt, 2026-11-10). The look is the
comparison figure Matt sent as what he wants.

Three consumers draw the same list: `draw_on_canvas` (the Analyzer tab's
tk.Canvas), `render_png` (Pillow; absent on a source checkout without it,
`PNG_AVAILABLE`) and `render_pdf` (a one-page PDF 1.4 written here, no
dependency). The PNG and the PDF are therefore what the screen shows, not a screen
grab. No tkinter is imported here; the canvas renderer takes the widget as an
argument. History: docs/decisions/toner.md.
"""
import math
import zlib

try:
    from PIL import Image, ImageDraw, ImageFont
    PNG_AVAILABLE = True
except ImportError:          # a source checkout without Pillow; the frozen builds carry it
    Image = ImageDraw = ImageFont = None
    PNG_AVAILABLE = False

# One colour and one marker per run, in order (the Okabe-Ito palette, with a
# marker each, so no line is told apart by colour alone).
LINE_STYLES = (
    ("#0072B2", "circle"), ("#D55E00", "square"), ("#009E73", "triangle"), ("#CC79A7", "diamond"),
    ("#E69F00", "triangle_down"), ("#56B4E9", "circle"), ("#000000", "square"),
)
MARKERS = ("circle", "square", "triangle", "diamond", "triangle_down")
BACKGROUND = "#FFFFFF"
TEXT = "#000000"
GRID = "#EEEEEE"
ZERO_LINE = "#999999"
REGISTER_LINE = "#CCCCCC"
AXIS = "#444444"
FONT_PX = 11          # the labels; the title is a little larger
LEGEND_COLUMNS = 4
MARKER_PX = 5         # half-size
Y_STEP = 10           # the grid, in cents
LINE_WIDTH = 2        # a run's line
HIGHLIGHT_WIDTH = 4   # the highlighted run's line (Matt, 2026-11-11: click a spot, the line is highlighted)
DIM = "#D0D0D0"       # every other run while one is highlighted: lines, markers, spread bars, legend
HIT_TOLERANCE_PX = 6  # how near a click must come to a run's line, marker or legend entry


def y_range(values):
    """The y axis in cents: every value and 0 inside it, with a 5-cent margin, on 10-cent
    steps; -10..10 when there is nothing."""
    vals = [v for v in values if v is not None] + [0.0]
    lo, hi = min(vals), max(vals)
    lo = Y_STEP * math.floor((lo - 5) / Y_STEP)
    hi = Y_STEP * math.ceil((hi + 5) / Y_STEP)
    if hi - lo < 2 * Y_STEP:
        lo, hi = min(lo, -Y_STEP), max(hi, Y_STEP)
    return lo, hi


def style_for(index):
    """The colour and marker of the run at `index` (cycling past the palette)."""
    return LINE_STYLES[index % len(LINE_STYLES)]


TITLE_SEPARATOR = " \u00b7 "     # between the runs' names in the chart's own title


def runs_title(lines, width, size):
    """The chart's own title line: one run's name, cut to what the width holds (0.6 px a
    character at `size`) with three dots; "" with no run, and "" with two or more (Matt,
    2026-11-11: "no title, just list the presets involved": the legend, which lists them,
    is the whole heading and sits where the title would). Until then the names were
    joined with TITLE_SEPARATOR above the legend that listed them again."""
    names = [str(line.get("name", "")) for line in lines]
    names = [n for n in names if n]
    if len(names) != 1:
        return ""
    text = names[0]
    room = max(8, int((width - 24) / (0.6 * size)))
    if len(text) > room:
        text = text[:room - 3].rstrip() + "..."
    return text


LABEL_PITCH = 1.25      # the room a turned note name needs along the axis, in font heights


def chart_primitives(overlay, width, height, *, title="", low_left=True, font_px=FONT_PX, footnote="",
                     highlight=None):
    """The chart for `overlay` (intonation.overlay(runs): notes low to high, registers,
    lines of points) as primitives in a width x height pixel frame. `title` is the line
    above the legend: a string as given, "" for none, None for the runs' own
    (`runs_title`: one run's name, or the names of several joined). `footnote` is one
    plain line under the notes (what was left out, and why). `highlight` is the index of
    the run to bring to the front, or None: that run is drawn last, in its own colour with
    a heavier line, and every other run (lines, markers, spread bars, legend entry) in
    DIM. Every run primitive carries its run index as "run", which is what run_at reads;
    the same primitives go to the canvas, the PNG and the PDF, so an export shows the
    highlight as the screen does."""
    notes = list(overlay.get("notes", []))
    registers = overlay.get("registers", {})
    lines = overlay.get("lines", [])
    if title is None:
        title = runs_title(lines, width, int(font_px * 1.2))
    if highlight is not None and not (isinstance(highlight, int) and 0 <= highlight < len(lines)):
        highlight = None

    def run_color(i, color):
        return DIM if highlight is not None and i != highlight else color
    prims = [{"kind": "background", "color": BACKGROUND, "width": width, "height": height}]
    # ---- the frame ------------------------------------------------------------------------
    title_h = int(font_px * 1.4) + 6 if title else 0
    rows = math.ceil(len(lines) / LEGEND_COLUMNS) if lines else 0
    legend_h = rows * (font_px + 8) + (6 if rows else 0)
    label_h = max((len(n) for n in notes), default=2) * font_px * 0.62 + 10     # the notes turned 90 degrees
    foot_h = font_px + 8 if footnote else 0
    left = int(font_px * 4.2)
    right = 12
    top = 8 + title_h + legend_h
    bottom = int(label_h) + 6 + foot_h
    plot_x0, plot_x1 = left, max(left + 1, width - right)
    plot_y0, plot_y1 = top, max(top + 1, height - bottom)
    plot_w, plot_h = plot_x1 - plot_x0, plot_y1 - plot_y0
    if title:
        prims.append({"kind": "text", "x": width / 2, "y": 8, "text": title, "color": TEXT,
                      "size": int(font_px * 1.2), "anchor": "n", "angle": 0, "bold": True})
    if footnote:
        prims.append({"kind": "text", "x": plot_x0, "y": height - 4, "text": footnote, "color": AXIS,
                      "size": font_px, "anchor": "sw", "angle": 0, "bold": False})
    # ---- the legend, above the plot, in columns ---------------------------------------------
    if lines:
        cols = min(LEGEND_COLUMNS, len(lines))
        col_w = plot_w / cols
        for i, line in enumerate(lines):
            color, marker = style_for(i)
            color = run_color(i, color)
            r, c = divmod(i, cols)
            x = plot_x0 + c * col_w + 4
            y = 8 + title_h + r * (font_px + 8) + (font_px + 8) / 2
            prims.append({"kind": "line", "x1": x, "y1": y, "x2": x + 22, "y2": y, "color": color, "width": 2,
                          "dash": None, "run": i})
            prims.append({"kind": "marker", "shape": marker, "x": x + 11, "y": y, "size": MARKER_PX, "color": color,
                          "run": i})
            prims.append({"kind": "text", "x": x + 28, "y": y, "text": str(line.get("name", "")),
                          "color": run_color(i, TEXT), "size": font_px, "anchor": "w", "angle": 0, "bold": False,
                          "run": i})
    # ---- the axes ------------------------------------------------------------------------
    cents = [p.cents for line in lines for p in line["points"].values()]
    lows = [p.low for line in lines for p in line["points"].values()]
    highs = [p.high for line in lines for p in line["points"].values()]
    y_lo, y_hi = y_range(cents + lows + highs)

    def y_of(c):
        return plot_y0 + (y_hi - c) / (y_hi - y_lo) * plot_h

    n = len(notes)
    order = notes if low_left else list(reversed(notes))
    x_of_note = {}
    for i, note in enumerate(order):
        x_of_note[note] = plot_x0 + (i + 0.5) * (plot_w / n) if n else plot_x0
    # the grid every 10 cents, the zero line solid, the y labels
    c = y_lo
    while c <= y_hi + 1e-9:
        y = y_of(c)
        prims.append({"kind": "line", "x1": plot_x0, "y1": y, "x2": plot_x1, "y2": y,
                      "color": ZERO_LINE if c == 0 else GRID, "width": 1, "dash": None})
        prims.append({"kind": "text", "x": plot_x0 - 4, "y": y, "text": f"{c:+.0f}" if c else "0", "color": AXIS,
                      "size": font_px, "anchor": "e", "angle": 0, "bold": False})
        c += Y_STEP
    prims.append({"kind": "text", "x": 4, "y": (plot_y0 + plot_y1) / 2, "text": "cents", "color": AXIS,
                  "size": font_px, "anchor": "w", "angle": 90, "bold": False})
    # the plot's box
    prims.append({"kind": "line", "x1": plot_x0, "y1": plot_y1, "x2": plot_x1, "y2": plot_y1, "color": AXIS,
                  "width": 1, "dash": None})
    prims.append({"kind": "line", "x1": plot_x0, "y1": plot_y0, "x2": plot_x0, "y2": plot_y1, "color": AXIS,
                  "width": 1, "dash": None})
    # the register change: a dashed line between the notes where it happens
    for i in range(1, n):
        if registers.get(order[i]) != registers.get(order[i - 1]):
            x = (x_of_note[order[i]] + x_of_note[order[i - 1]]) / 2
            prims.append({"kind": "line", "x1": x, "y1": plot_y0, "x2": x, "y2": plot_y1, "color": REGISTER_LINE,
                          "width": 1, "dash": (4, 3)})
    # the notes along the bottom, turned 90 degrees, their end at the axis. In a narrow
    # frame (the chart beside the spectrum or the meter, 2026-11-12) there is not room for
    # every name: every second, third... is written, from the low end, and every point is
    # still drawn.
    step = max(1, math.ceil(font_px * LABEL_PITCH / (plot_w / n))) if n else 1
    labelled = set(notes[::step])
    for note in order:
        if note not in labelled:
            continue
        prims.append({"kind": "text", "x": x_of_note[note], "y": plot_y1 + 4, "text": note, "color": TEXT,
                      "size": font_px, "anchor": "e", "angle": 90, "bold": False})
    # ---- the runs: the highlighted one last, so it lies on top ----------------------------
    draw_order = [i for i in range(len(lines)) if i != highlight] + ([highlight] if highlight is not None else [])
    for i in draw_order:
        line = lines[i]
        color, marker = style_for(i)
        color = run_color(i, color)
        pts = sorted(((x_of_note[nm], p) for nm, p in line["points"].items() if nm in x_of_note), key=lambda t: t[0])
        for x, p in pts:
            if p.takes >= 2 and p.high > p.low:
                prims.append({"kind": "line", "x1": x, "y1": y_of(p.low), "x2": x, "y2": y_of(p.high), "color": color,
                              "width": 1, "dash": None, "run": i})
        if len(pts) >= 2:
            prims.append({"kind": "polyline", "points": [(x, y_of(p.cents)) for x, p in pts], "color": color,
                          "width": HIGHLIGHT_WIDTH if i == highlight else LINE_WIDTH, "run": i})
        for x, p in pts:
            prims.append({"kind": "marker", "shape": marker, "x": x, "y": y_of(p.cents), "size": MARKER_PX,
                          "color": color, "run": i})
    return prims


def _segment_distance(x, y, x1, y1, x2, y2):
    """The distance from (x, y) to the segment (x1, y1)-(x2, y2)."""
    dx, dy = x2 - x1, y2 - y1
    if dx == 0 and dy == 0:
        return math.hypot(x - x1, y - y1)
    t = max(0.0, min(1.0, ((x - x1) * dx + (y - y1) * dy) / (dx * dx + dy * dy)))
    return math.hypot(x - (x1 + t * dx), y - (y1 + t * dy))


def run_at(prims, x, y, tolerance=HIT_TOLERANCE_PX):
    """The index of the run whose line, marker, spread bar or legend entry (its line,
    marker or name) lies nearest to the pixel (x, y), within `tolerance`; None when
    nothing is that near, so a click there counts as empty chart area. Reads the "run"
    the primitives carry: what the chart shows is what can be hit."""
    best_run, best_d = None, float(tolerance)
    for p in prims:
        run = p.get("run")
        if run is None:
            continue
        k = p["kind"]
        if k == "polyline":
            pts = p["points"]
            d = min((_segment_distance(x, y, a[0], a[1], b[0], b[1]) for a, b in zip(pts, pts[1:])), default=None)
            if d is None:
                continue
        elif k == "line":
            d = _segment_distance(x, y, p["x1"], p["y1"], p["x2"], p["y2"])
        elif k == "marker":
            d = max(0.0, math.hypot(x - p["x"], y - p["y"]) - p["size"])
        elif k == "text":
            # a legend name, anchored west: its box as the renderers size text, 0.6 px a character
            x0, x1 = p["x"], p["x"] + len(p["text"]) * 0.6 * p["size"]
            y0, y1 = p["y"] - p["size"] / 2, p["y"] + p["size"] / 2
            d = max(0.0, x0 - x, x - x1, y0 - y, y - y1)
        else:
            continue
        if d <= best_d and (best_run is None or d < best_d):
            best_run, best_d = run, d
    return best_run


def marker_polygon(shape, x, y, size):
    """The corners of a marker, or None for a circle."""
    s = size
    if shape == "square":
        return [(x - s, y - s), (x + s, y - s), (x + s, y + s), (x - s, y + s)]
    if shape == "triangle":
        return [(x, y - s * 1.2), (x + s * 1.15, y + s * 0.8), (x - s * 1.15, y + s * 0.8)]
    if shape == "triangle_down":
        return [(x, y + s * 1.2), (x + s * 1.15, y - s * 0.8), (x - s * 1.15, y - s * 0.8)]
    if shape == "diamond":
        return [(x, y - s * 1.3), (x + s * 1.3, y), (x, y + s * 1.3), (x - s * 1.3, y)]
    return None


# ---- the tk.Canvas renderer -----------------------------------------------------------

def draw_on_canvas(canvas, prims, tag="chart"):
    """Draw the primitives on a tk.Canvas (replacing what carried `tag`)."""
    canvas.delete(tag)
    for p in prims:
        k = p["kind"]
        if k == "background":
            canvas.create_rectangle(0, 0, p["width"], p["height"], fill=p["color"], outline="", tags=tag)
        elif k == "line":
            kw = {"fill": p["color"], "width": p["width"], "tags": tag}
            if p.get("dash"):
                kw["dash"] = p["dash"]
            canvas.create_line(p["x1"], p["y1"], p["x2"], p["y2"], **kw)
        elif k == "polyline":
            flat = [c for pt in p["points"] for c in pt]
            canvas.create_line(*flat, fill=p["color"], width=p["width"], tags=tag)
        elif k == "marker":
            poly = marker_polygon(p["shape"], p["x"], p["y"], p["size"])
            if poly is None:
                s = p["size"]
                canvas.create_oval(p["x"] - s, p["y"] - s, p["x"] + s, p["y"] + s, fill=p["color"],
                                   outline=p["color"], tags=tag)
            else:
                canvas.create_polygon(*[c for pt in poly for c in pt], fill=p["color"], outline=p["color"], tags=tag)
        elif k == "text":
            weight = "bold" if p.get("bold") else "normal"
            canvas.create_text(p["x"], p["y"], text=p["text"], fill=p["color"], anchor=p["anchor"],
                               angle=p.get("angle", 0), font=("Helvetica", -int(p["size"]), weight), tags=tag)


# ---- the PNG renderer (Pillow) --------------------------------------------------------

def _font(size, bold=False):
    try:
        return ImageFont.load_default(size=size)
    except TypeError:                 # a Pillow before 10.1: one size only
        return ImageFont.load_default()


def _anchor_offset(anchor, w, h):
    """Where a w x h box's top-left sits relative to its anchor point, for any of the
    canvas anchors (n, s, e, w, their corners nw/ne/sw/se, center)."""
    a = anchor or "center"
    if a == "center":
        return -w / 2, -h / 2
    dx = 0 if "w" in a else (-w if "e" in a else -w / 2)
    dy = 0 if "n" in a else (-h if "s" in a else -h / 2)
    return dx, dy


def render_png(prims, width, height, scale=1):
    """The chart as PNG bytes, `scale` times the pixel size (2 for a sharper file)."""
    if not PNG_AVAILABLE:
        raise RuntimeError("Pillow is not available: no PNG export")
    img = Image.new("RGB", (int(width * scale), int(height * scale)), BACKGROUND)
    draw = ImageDraw.Draw(img)
    S = scale
    for p in prims:
        k = p["kind"]
        if k == "background":
            draw.rectangle([0, 0, p["width"] * S, p["height"] * S], fill=p["color"])
        elif k == "line":
            if p.get("dash"):
                _dashed(draw, p["x1"] * S, p["y1"] * S, p["x2"] * S, p["y2"] * S, p["color"], p["width"] * S,
                        [d * S for d in p["dash"]])
            else:
                draw.line([(p["x1"] * S, p["y1"] * S), (p["x2"] * S, p["y2"] * S)], fill=p["color"],
                          width=max(1, int(round(p["width"] * S))))
        elif k == "polyline":
            draw.line([(x * S, y * S) for x, y in p["points"]], fill=p["color"],
                      width=max(1, int(round(p["width"] * S))), joint="curve")
        elif k == "marker":
            poly = marker_polygon(p["shape"], p["x"] * S, p["y"] * S, p["size"] * S)
            if poly is None:
                s = p["size"] * S
                draw.ellipse([p["x"] * S - s, p["y"] * S - s, p["x"] * S + s, p["y"] * S + s], fill=p["color"])
            else:
                draw.polygon(poly, fill=p["color"])
        elif k == "text":
            font = _font(int(p["size"] * S), p.get("bold"))
            text = p["text"]
            box = draw.textbbox((0, 0), text, font=font)
            w, h = box[2] - box[0], box[3] - box[1]
            if p.get("angle", 0) == 90:
                layer = Image.new("RGBA", (max(1, w + 2), max(1, h + 2)), (0, 0, 0, 0))
                ImageDraw.Draw(layer).text((-box[0] + 1, -box[1] + 1), text, font=font, fill=p["color"])
                layer = layer.rotate(90, expand=True)
                # anchor "e": the text's end (now its top) at the point; "w": its start (bottom) there
                lw, lh = layer.size
                x, y = p["x"] * S, p["y"] * S
                if p["anchor"] == "e":
                    img.paste(layer, (int(x - lw / 2), int(y)), layer)
                elif p["anchor"] == "w":
                    img.paste(layer, (int(x - lw / 2), int(y - lh)), layer)
                else:
                    img.paste(layer, (int(x - lw / 2), int(y - lh / 2)), layer)
            else:
                dx, dy = _anchor_offset(p["anchor"], w, h)
                draw.text((p["x"] * S + dx - box[0], p["y"] * S + dy - box[1]), text, font=font, fill=p["color"])
    import io
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def _dashed(draw, x1, y1, x2, y2, color, width, dash):
    length = math.hypot(x2 - x1, y2 - y1)
    if length == 0:
        return
    ux, uy = (x2 - x1) / length, (y2 - y1) / length
    on, off = dash[0], dash[1] if len(dash) > 1 else dash[0]
    t = 0.0
    while t < length:
        t2 = min(length, t + on)
        draw.line([(x1 + ux * t, y1 + uy * t), (x1 + ux * t2, y1 + uy * t2)], fill=color,
                  width=max(1, int(round(width))))
        t = t2 + off


# ---- the PDF renderer (one page, written here) ------------------------------------------

# Helvetica's advance widths for the printable ASCII characters, in 1/1000 em (the standard
# AFM metrics), so the anchors land where the canvas puts them.
_HELVETICA_WIDTHS = [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556,
    556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667,
    556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556,
    556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722,
    500, 500, 500, 334, 260, 334, 584,
]


def _text_width(text, size):
    total = 0
    for ch in text:
        code = ord(ch)
        total += _HELVETICA_WIDTHS[code - 32] if 32 <= code < 127 else 556
    return total * size / 1000.0


def _pdf_string(text):
    raw = text.encode("cp1252", "replace")
    return b"(" + raw.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)") + b")"


def _rgb(color):
    color = color.lstrip("#")
    r, g, b = (int(color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return f"{r:.3f} {g:.3f} {b:.3f}"


def render_pdf(prims, width, height, title="Intonation chart"):
    """The chart as a one-page PDF 1.4 (bytes), one point per pixel."""
    H = float(height)
    ops = []

    def y(v):
        return H - v

    for p in prims:
        k = p["kind"]
        if k == "background":
            ops.append(f"{_rgb(p['color'])} rg 0 0 {p['width']} {p['height']} re f")
        elif k == "line":
            dash = f"[{p['dash'][0]} {p['dash'][1]}] 0 d" if p.get("dash") else "[] 0 d"
            ops.append(f"q {_rgb(p['color'])} RG {p['width']} w {dash} {p['x1']:.2f} {y(p['y1']):.2f} m "
                       f"{p['x2']:.2f} {y(p['y2']):.2f} l S Q")
        elif k == "polyline":
            pts = p["points"]
            path = " ".join(f"{x:.2f} {y(yy):.2f} {'m' if i == 0 else 'l'}" for i, (x, yy) in enumerate(pts))
            ops.append(f"q {_rgb(p['color'])} RG {p['width']} w 1 j 1 J {path} S Q")
        elif k == "marker":
            poly = marker_polygon(p["shape"], p["x"], p["y"], p["size"])
            if poly is None:
                ops.append(_circle_ops(p["x"], y(p["y"]), p["size"], p["color"]))
            else:
                path = " ".join(f"{x:.2f} {y(yy):.2f} {'m' if i == 0 else 'l'}" for i, (x, yy) in enumerate(poly))
                ops.append(f"q {_rgb(p['color'])} rg {path} h f Q")
        elif k == "text":
            size = p["size"]
            w = _text_width(p["text"], size)
            h = size * 0.72                     # the cap height, near enough for anchoring
            angle = p.get("angle", 0)
            if angle == 90:
                # the text runs upward; its baseline is a vertical line to the right of the anchor point
                if p["anchor"] == "e":
                    bx, by = p["x"] + h / 2, p["y"] + w          # its end at the point: it starts w below
                elif p["anchor"] == "w":
                    bx, by = p["x"] + h / 2, p["y"]
                else:
                    bx, by = p["x"] + h / 2, p["y"] + w / 2
                ops.append(f"BT /F1 {size} Tf {_rgb(p['color'])} rg 0 1 -1 0 {bx:.2f} {y(by):.2f} Tm "
                           + _pdf_string(p["text"]).decode("latin-1") + " Tj ET")
            else:
                dx, dy = _anchor_offset(p["anchor"], w, h)
                bx, by = p["x"] + dx, p["y"] + dy + h           # the baseline
                ops.append(f"BT /F1 {size} Tf {_rgb(p['color'])} rg {bx:.2f} {y(by):.2f} Td "
                           + _pdf_string(p["text"]).decode("latin-1") + " Tj ET")
    content = "\n".join(ops).encode("latin-1", "replace")
    stream = zlib.compress(content)
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Contents 4 0 R "
         f"/Resources << /Font << /F1 5 0 R >> >> >>").encode(),
        b"<< /Length " + str(len(stream)).encode() + b" /Filter /FlateDecode >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        b"<< /Title " + _pdf_string(title) + b" /Producer (Stohrer Sax Shop Companion) >>",
    ]
    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for i, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R /Info 6 0 R >>\nstartxref\n{xref}\n%%EOF\n").encode()
    return bytes(out)


def _circle_ops(cx, cy, r, color):
    k = 0.5523 * r
    return (f"q {_rgb(color)} rg {cx + r:.2f} {cy:.2f} m "
            f"{cx + r:.2f} {cy + k:.2f} {cx + k:.2f} {cy + r:.2f} {cx:.2f} {cy + r:.2f} c "
            f"{cx - k:.2f} {cy + r:.2f} {cx - r:.2f} {cy + k:.2f} {cx - r:.2f} {cy:.2f} c "
            f"{cx - r:.2f} {cy - k:.2f} {cx - k:.2f} {cy - r:.2f} {cx:.2f} {cy - r:.2f} c "
            f"{cx + k:.2f} {cy - r:.2f} {cx + r:.2f} {cy - k:.2f} {cx + r:.2f} {cy:.2f} c f Q")


__all__ = ["BACKGROUND", "LINE_STYLES", "MARKERS", "PNG_AVAILABLE", "Y_STEP", "chart_primitives", "draw_on_canvas",
           "marker_polygon", "render_pdf", "render_png", "style_for", "y_range"]
