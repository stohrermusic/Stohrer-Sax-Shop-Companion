"""intonation_chart.py: one geometry, three renderers.

The primitives for known runs (where each note sits, the dashed register line, low on
the left and flipped, one colour and marker per run, the legend, the title), and the
PNG and the PDF drawn from the same primitives (the PNG opens at the expected size
with non-background pixels where a line is; the PDF parses as one page, every object
at its xref offset, and carries the title). Headless; the canvas renderer is driven
by the Analyzer suite. Run:
    python tools/test_intonation_chart.py
"""
import io
import os
import re
import sys
import tempfile
import traceback
import zlib

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Never the user's real settings folder: an isolated profile, set before any
# project module is imported (tools/test_config_isolation.py checks this).
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-intonation-chart-")

import intonation as I  # noqa: E402
import intonation_chart as C  # noqa: E402
from toner_engine import _shift_note  # noqa: E402

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


def _a_narrow_chart_thins_its_names():
    """Beside the spectrum or the meter the chart is a third of the window: not every note
    name fits. Names are thinned from the low end; every point and marker is still drawn."""
    def names(prims, notes):
        return [p["text"] for p in prims if p["kind"] == "text" and p.get("angle") == 90 and p["text"] in notes]
    notes = OV["notes"]
    wide = C.chart_primitives(OV, 900, 420)
    assert names(wide, notes) == notes, "a wide chart names every note"
    narrow = C.chart_primitives(OV, 300, 300)
    got = names(narrow, notes)
    assert 2 <= len(got) < len(notes), (len(got), len(notes))
    assert got[0] == notes[0], "the thinning starts at the lowest note"
    step = notes.index(got[1]) - notes.index(got[0])
    assert got == notes[::step], (got, step)
    # no two names closer than a name needs, and nothing else thinned
    xs = sorted(p["x"] for p in narrow if p["kind"] == "text" and p.get("angle") == 90 and p["text"] in notes)
    assert all(b - a >= C.FONT_PX * C.LABEL_PITCH - 1e-6 for a, b in zip(xs, xs[1:])), xs

    def markers(prims):
        return sum(1 for p in prims if p.get("run") is not None and p["kind"] not in ("line", "text"))
    assert markers(narrow) == markers(wide) > 0, (markers(narrow), markers(wide))
    assert C.run_at(narrow, *next((p["x"], p["y"]) for p in narrow if p.get("run") == 0 and "x" in p and p["kind"] != "text")) == 0


def hz(midi, a4=440.0):
    return a4 * 2 ** ((midi - 69) / 12)


def cap(concert_midi, cents, frames=30):
    return {"note": _shift_note("C4", concert_midi - 60), "fundamental_freq": hz(concert_midi) * 2 ** (cents / 1200),
            "harmonics_db": [0.0, -3.0, -6.0], "n_frames": frames, "method": "free"}


# Two tenor runs: A over concert G#2..E4 (written Bb3..F#5, both registers), B over part of it.
RUN_A = I.run_from_captures([{"captures": [cap(44 + k, -6 + k) for k in range(0, 19)]
                                           + [cap(44, 4)]}], "Tenor", name="Mark VI, Link 7*")   # Bb3 twice: 2 takes
RUN_B = I.run_from_captures([{"captures": [cap(46 + k, 3 - k) for k in range(0, 8)]}], "Tenor", name="Conn 10M")
OV = I.overlay([RUN_A, RUN_B])

check("a narrow chart thins its note names from the low end and still draws every point",
      _a_narrow_chart_thins_its_names)
W, H = 800, 400
PRIMS = C.chart_primitives(OV, W, H, title="Two horns")


def of(kind, prims=None):
    return [p for p in (prims or PRIMS) if p["kind"] == kind]


def notes_x(prims):
    """{note: x} from the rotated labels along the bottom."""
    return {p["text"]: p["x"] for p in of("text", prims) if p.get("angle") == 90 and p["text"] != "cents"}


print("Intonation chart")
print("=" * 60)


def the_frame_and_the_axis():
    assert OV["notes"][0] == "Bb3" and OV["notes"][-1] == "E5" and len(OV["notes"]) == 19, "OV[\"notes\"][0] == \"Bb3\" and OV[\"notes\"][-1] == \"E5\" and len(OV[\"notes\"]) == 19"
    assert PRIMS[0]["kind"] == "background" and PRIMS[0]["width"] == W and PRIMS[0]["height"] == H, "PRIMS[0][\"kind\"] == \"background\" and PRIMS[0][\"width\"] == W and PRIMS[0][\"height\"] == H"
    xs = notes_x(PRIMS)
    assert set(xs) == set(OV["notes"]), set(xs) ^ set(OV["notes"])
    order = sorted(xs, key=xs.get)
    assert order == OV["notes"], "low notes are on the left, in order"
    gaps = [xs[b] - xs[a] for a, b in zip(order, order[1:])]
    assert max(gaps) - min(gaps) < 1e-6, "the notes are evenly spaced"
    labels = [p for p in of("text") if p.get("angle") == 90 and p["text"] != "cents"]
    assert all(p["anchor"] == "e" for p in labels), "the labels hang from the axis, turned 90 degrees"
    y_axis = [p for p in of("text") if p.get("angle") == 0 and p["anchor"] == "e"]
    ticks = [p["text"] for p in y_axis]
    lo, hi = C.y_range([p.cents for ln in OV["lines"] for p in ln["points"].values()])
    assert (lo, hi) == (-20, 20), "(lo, hi) == (-20, 20)"                                   # cents -6..+12 and 0, a 5-cent margin, 10-cent steps
    assert ticks == [(f"{c:+.0f}" if c else "0") for c in range(lo, hi + 1, 10)], ticks
    grid = [p for p in of("line") if p["color"] in (C.GRID, C.ZERO_LINE)]
    assert len(grid) == 5 and sum(1 for p in grid if p["color"] == C.ZERO_LINE) == 1, "a line every 10 cents, the zero line solid"
    assert any(p["text"] == "cents" and p["angle"] == 90 for p in of("text")), "any(p[\"text\"] == \"cents\" and p[\"angle\"] == 90 for p in of(\"text\"))"
    assert any(p["text"] == "Two horns" and p["bold"] for p in of("text")), "the title line"
check("the frame: the notes along the bottom low to high, evenly spaced, turned 90 degrees; a grid every 10 cents "
      "with the zero line solid; the y labels; the title", the_frame_and_the_axis)


def the_register_change_is_dashed_between_c_sharp5_and_d5():
    dashed = [p for p in of("line") if p.get("dash")]
    assert len(dashed) == 1 and dashed[0]["color"] == C.REGISTER_LINE, dashed
    xs = notes_x(PRIMS)
    x = dashed[0]["x1"]
    assert xs["C#5"] < x < xs["D5"] and abs(x - (xs["C#5"] + xs["D5"]) / 2) < 1e-6, (x, xs["C#5"], xs["D5"])
    assert dashed[0]["x2"] == x, "vertical"
    # a run within one register has no dashed line
    one = C.chart_primitives(I.overlay([RUN_B]), W, H)
    assert not [p for p in of("line", one) if p.get("dash")], "not [p for p in of(\"line\", one) if p.get(\"dash\")]"
check("the register change is one dashed vertical line, halfway between written C#5 and D5",
      the_register_change_is_dashed_between_c_sharp5_and_d5)


def flipped_puts_the_low_notes_on_the_right():
    prims = C.chart_primitives(OV, W, H, low_left=False)
    xs = notes_x(prims)
    assert sorted(xs, key=xs.get) == list(reversed(OV["notes"])), "sorted(xs, key=xs.get) == list(reversed(OV[\"notes\"]))"
    assert xs["Bb3"] > xs["E5"], "xs[\"Bb3\"] > xs[\"E5\"]"
    dashed = [p for p in of("line", prims) if p.get("dash")]
    assert xs["D5"] < dashed[0]["x1"] < xs["C#5"], "xs[\"D5\"] < dashed[0][\"x1\"] < xs[\"C#5\"]"
check("the preference flips it: low notes on the right, the register line between the same two notes",
      flipped_puts_the_low_notes_on_the_right)


def one_colour_and_marker_per_run():
    markers = of("marker")
    polys = of("polyline")
    assert len(polys) == 2 and polys[0]["color"] == C.LINE_STYLES[0][0] and polys[1]["color"] == C.LINE_STYLES[1][0], "len(polys) == 2 and polys[0][\"color\"] == C.LINE_STYLES[0][0] and polys[1][\"color\"] == C.LINE_STYLES[1][0]"
    a = [m for m in markers if m["color"] == C.LINE_STYLES[0][0]]
    b = [m for m in markers if m["color"] == C.LINE_STYLES[1][0]]
    assert {m["shape"] for m in a} == {"circle"} and {m["shape"] for m in b} == {"square"}, "{m[\"shape\"] for m in a} == {\"circle\"} and {m[\"shape\"] for m in b} == {\"square\"}"
    assert len(a) == 19 + 1 and len(b) == 8 + 1, (len(a), len(b))        # the notes plus the legend's marker
    assert len(polys[0]["points"]) == 19 and len(polys[1]["points"]) == 8, "len(polys[0][\"points\"]) == 19 and len(polys[1][\"points\"]) == 8"
    xs = notes_x(PRIMS)
    assert [round(x) for x, _y in polys[1]["points"]] == [round(xs[n]) for n in RUN_B.notes], "B's points sit on its notes"
    # the spread bar: Bb3 has two takes (4 and -6 cents): a vertical bar in A's colour, none for single takes
    bars = [p for p in of("line") if p["color"] == C.LINE_STYLES[0][0] and p["x1"] == p["x2"] and not p.get("dash")]
    assert len(bars) == 1 and abs(bars[0]["x1"] - xs["Bb3"]) < 1e-6, bars
    assert not [p for p in of("line") if p["color"] == C.LINE_STYLES[1][0] and p["x1"] == p["x2"] and p["y1"] != p["y2"]], "not [p for p in of(\"line\") if p[\"color\"] == C.LINE_STYLES[1][0] and p[\"x1\"] == p[\"x2\"] and p[\"y1\"] != p[\"y2\"]]"
    # seven styles, then they cycle; five marker shapes
    assert len(C.LINE_STYLES) == 7 and C.style_for(7) == C.LINE_STYLES[0] and set(C.MARKERS) >= {s for _c, s in C.LINE_STYLES}, "len(C.LINE_STYLES) == 7 and C.style_for(7) == C.LINE_STYLES[0] and set(C.MARKERS) >= {s for _c, s in C.LINE..."
check("one colour and one marker per run, the medians joined, a bar where a note has several takes",
      one_colour_and_marker_per_run)


def the_legend_names_each_run_above_the_plot():
    names = [p["text"] for p in of("text") if p["anchor"] == "w" and p["angle"] == 0]
    assert names == ["Mark VI, Link 7*", "Conn 10M"], names
    legend_y = [p["y"] for p in of("text") if p["text"] in names]
    plot_top = min(p["y1"] for p in of("line") if p["color"] in (C.GRID, C.ZERO_LINE))
    assert max(legend_y) < plot_top, "the legend sits above the plot"
    title_y = [p["y"] for p in of("text") if p["text"] == "Two horns"][0]
    assert title_y < min(legend_y), "the title is above the legend"
    # six runs: two rows of up to four columns
    six = I.overlay([RUN_A] * 6)
    for i, ln in enumerate(six["lines"]):
        ln["name"] = f"run {i}"
    prims = C.chart_primitives(six, W, H)
    ys = sorted({round(p["y"]) for p in of("text", prims) if p["text"].startswith("run ")})
    assert len(ys) == 2, ys
check("the legend names each run with its marker, above the plot and below the title, in columns",
      the_legend_names_each_run_above_the_plot)


def the_y_range_rule():
    assert C.y_range([]) == (-10, 10), "C.y_range([]) == (-10, 10)"
    assert C.y_range([3.0, -4.0]) == (-10, 10), "C.y_range([3.0, -4.0]) == (-10, 10)"
    assert C.y_range([12.0]) == (-10, 20), "C.y_range([12.0]) == (-10, 20)"
    assert C.y_range([-37.0, 2.0]) == (-50, 10), "C.y_range([-37.0, 2.0]) == (-50, 10)"
    assert C.y_range([55.0, 60.0]) == (-10, 70), "C.y_range([55.0, 60.0]) == (-10, 70)"      # 0 is always inside
check("the y range: every value and 0 inside, a 5-cent margin, 10-cent steps, never narrower than -10..10",
      the_y_range_rule)


def the_png_is_the_same_picture():
    assert C.PNG_AVAILABLE, "Pillow is in the lock"
    from PIL import Image
    data = C.render_png(PRIMS, W, H, scale=2)
    img = Image.open(io.BytesIO(data))
    assert img.format == "PNG" and img.size == (2 * W, 2 * H), (img.format, img.size)
    px = img.convert("RGB").load()
    # the zero line is drawn where the primitives say, in its colour
    zero = [p for p in of("line") if p["color"] == C.ZERO_LINE][0]
    x, y = int((zero["x1"] + zero["x2"]) / 2) * 2, int(round(zero["y1"] * 2))
    assert px[x, y] != (255, 255, 255), "the zero line's pixels are not background"
    # a marker of run A is filled with A's colour
    m = [p for p in of("marker") if p["color"] == C.LINE_STYLES[0][0]][1]
    assert px[int(m["x"] * 2), int(m["y"] * 2)] == (0, 114, 178), px[int(m["x"] * 2), int(m["y"] * 2)]
    # and a corner is background
    assert px[2, 2] == (255, 255, 255), "px[2, 2] == (255, 255, 255)"
check("the PNG opens at the expected size and has the zero line and a marker where the primitives put them",
      the_png_is_the_same_picture)


def the_pdf_is_one_page_with_the_title():
    data = C.render_pdf(PRIMS, W, H, title="Two horns")
    assert data.startswith(b"%PDF-1.4") and data.rstrip().endswith(b"%%EOF"), "data.startswith(b\"%PDF-1.4\") and data.rstrip().endswith(b\"%%EOF\")"
    start = int(re.search(rb"startxref\n(\d+)", data).group(1))
    assert data[start:start + 4] == b"xref", "data[start:start + 4] == b\"xref\""
    rows = re.findall(rb"(\d{10}) (\d{5}) ([nf])", data[start:])
    offsets = [int(r[0]) for r in rows if r[2] == b"n"]
    for i, off in enumerate(offsets, start=1):
        assert data[off:].startswith(f"{i} 0 obj".encode()), (i, data[off:off + 12])
    assert data.count(b"/Type /Page ") == 1 and b"/Count 1" in data, "data.count(b\"/Type /Page \") == 1 and b\"/Count 1\" in data"
    assert b"/Title (Two horns)" in data and b"/BaseFont /Helvetica" in data, "b\"/Title (Two horns)\" in data and b\"/BaseFont /Helvetica\" in data"
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", data, re.S).group(1)).decode("latin-1")
    assert "(Two horns) Tj" in content and "(Mark VI, Link 7*) Tj" in content, "\"(Two horns) Tj\" in content and \"(Mark VI, Link 7*) Tj\" in content"
    assert content.count("0 1 -1 0") == 19 + 1, "every note label and the axis label are turned 90 degrees"
    assert "[4 3] 0 d" in content, "the register line is dashed"
    assert "re f" in content and content.count(" c f Q") >= 20, "a background and the circle markers"
    # a parenthesis in a name is escaped
    ov = I.overlay([RUN_A])
    ov["lines"][0]["name"] = "Link (7*)"
    data = C.render_pdf(C.chart_primitives(ov, W, H), W, H)
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", data, re.S).group(1)).decode("latin-1")
    assert "(Link \\(7*\\)) Tj" in content, "\"(Link \\\\(7*\\\\)) Tj\" in content"
check("the PDF is one page, every object at its xref offset, Helvetica, the title and the run names as text, "
      "the notes turned 90 degrees, the register line dashed", the_pdf_is_one_page_with_the_title)


def a_single_run_is_titled_with_its_name_and_a_comparison_has_no_title():
    """Matt, 2026-11-11: "no title, just list the presets involved". One run's chart is titled with its
    name (cut to the width with three dots); two or more carry no title line and the legend, which lists
    them, sits where the title would; the canvas, the PNG and the PDF agree since the primitives are the
    same. `title=""` is none and a given string is drawn as given. (Between round 10 and this decision
    the names were joined above a legend that listed them again.)"""
    size = int(C.FONT_PX * 1.2)

    def title_of(prims):
        return [p for p in of("text", prims) if p["y"] == 8 and p["anchor"] == "n"]

    def legend_y(prims):
        return min(p["y1"] for p in of("line", prims) if p["dash"] is None and p["x2"] - p["x1"] == 22)

    one = C.chart_primitives(I.overlay([RUN_B]), W, H, title=None)
    t = title_of(one)
    assert len(t) == 1 and t[0]["text"] == "Conn 10M" and t[0]["bold"] and t[0]["size"] == size, t
    two = C.chart_primitives(OV, W, H, title=None)
    assert not title_of(two), f"a comparison drew a title: {title_of(two)}"
    assert legend_y(two) == 8 + (C.FONT_PX + 8) / 2, f"the legend did not move up into the title's place: {legend_y(two)}"
    assert legend_y(C.chart_primitives(OV, W, H, title="Two horns")) > legend_y(two), "a given title no longer pushes the legend down"
    assert not title_of(C.chart_primitives(OV, W, H, title="")), "an explicit empty title drew one"
    assert not title_of(C.chart_primitives(I.overlay([]), W, H, title=None)), "an empty chart has no title"
    many = I.overlay([RUN_A, RUN_B] * 3)
    assert not title_of(C.chart_primitives(many, W, H, title=None)), "six runs drew a title"
    long_one = I.overlay([RUN_A])
    long_one["lines"][0]["name"] = "A long preset name on a tenor, " * 5
    cut = title_of(C.chart_primitives(long_one, W, H, title=None))[0]["text"]
    assert cut.endswith("...") and len(cut) <= (W - 24) / (0.6 * size), cut
    assert C.runs_title(OV["lines"], 4000, size) == "" and C.runs_title(OV["lines"][:1], 4000, size) == "Mark VI, Link 7*", \
        "runs_title's rule"
    # the renderers follow: the comparison's PDF has the names once each (the legend) and no joined title;
    # the single run's PDF has its name twice (title and legend)
    pdf = C.render_pdf(two, W, H)
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", pdf, re.S).group(1)).decode("latin-1")
    assert "\xb7" not in content and content.count("(Mark VI, Link 7*) Tj") == 1 and content.count("(Conn 10M) Tj") == 1, \
        "the comparison's PDF carries a title"
    pdf1 = C.render_pdf(one, W, H)
    content1 = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", pdf1, re.S).group(1)).decode("latin-1")
    assert content1.count("(Conn 10M) Tj") == 2, "the single run's PDF lacks its title"
    from PIL import Image
    import io
    band = Image.open(io.BytesIO(C.render_png(two, W, H))).convert("L").crop((0, 8, W, 8 + C.FONT_PX + 8))
    assert min(band.getdata()) < 128, "no ink where the legend should now sit in the PNG"
check("a single-run chart is titled with its name (cut to the width); a comparison carries no title and its legend "
      "moves up into the title's place; a given title is drawn as given; the PDF and PNG follow",
      a_single_run_is_titled_with_its_name_and_a_comparison_has_no_title)


def a_highlighted_run_in_the_primitives_and_the_hit_test():
    """Matt, 2026-11-11: "click a spot in the comparison chart and have that line highlighted". In the
    primitives: the highlighted run is drawn last, heavier, in its own colour; every other run's line,
    markers, spread bars and legend entry are DIM; without a highlight the picture is as before, only
    every run primitive carries its run index. run_at reads the run from the primitives: a line
    segment, a marker, a legend entry, within HIT_TOLERANCE_PX; beyond it, or on empty area, None."""
    def runs(prims):
        return [p["run"] for p in of("polyline", prims)]

    def strip(prims):
        return [{k: v for k, v in p.items() if k != "run"} for p in prims]
    plain = C.chart_primitives(OV, W, H, title="Two horns")
    assert runs(plain) == [0, 1] and all(p["width"] == C.LINE_WIDTH for p in of("polyline", plain)), runs(plain)
    assert all("run" in p for p in plain if p["kind"] in ("polyline", "marker")), "a run primitive without its index"
    assert strip(plain) == strip(PRIMS), "no highlight must give the same picture as before"
    hi = C.chart_primitives(OV, W, H, title="Two horns", highlight=0)
    assert runs(hi) == [1, 0], f"the highlighted run is not drawn last: {runs(hi)}"
    polys = {p["run"]: p for p in of("polyline", hi)}
    assert polys[0]["width"] == C.HIGHLIGHT_WIDTH and polys[0]["color"] == C.LINE_STYLES[0][0], polys[0]
    assert polys[1]["width"] == C.LINE_WIDTH and polys[1]["color"] == C.DIM, polys[1]
    assert {p["color"] for p in of("marker", hi) if p["run"] == 1} == {C.DIM}, "the other run's markers are not dim"
    assert {p["color"] for p in of("marker", hi) if p["run"] == 0} == {C.LINE_STYLES[0][0]}, "the highlighted run's markers changed"
    bars = [p for p in of("line", hi) if p.get("run") == 0 and p["x1"] == p["x2"]]
    assert len(bars) == 1 and bars[0]["color"] == C.LINE_STYLES[0][0], bars
    legend = {p["run"]: p for p in of("text", hi) if p.get("run") is not None}
    assert legend[0]["color"] == C.TEXT and legend[1]["color"] == C.DIM, {k: v["color"] for k, v in legend.items()}
    legend_lines = {p["run"]: p for p in of("line", hi) if p.get("run") is not None and p["y1"] == p["y2"]}
    assert legend_lines[1]["color"] == C.DIM and legend_lines[0]["color"] == C.LINE_STYLES[0][0], legend_lines
    # the same set of primitives, only reordered and recoloured: nothing added, nothing lost
    assert len(hi) == len(plain), (len(hi), len(plain))
    # the other way round, and nonsense indexes mean none
    assert runs(C.chart_primitives(OV, W, H, highlight=1)) == [0, 1] and of("polyline", C.chart_primitives(OV, W, H, highlight=1))[1]["width"] == C.HIGHLIGHT_WIDTH
    for bad in (5, -1, "0", 1.0, None):
        assert runs(C.chart_primitives(OV, W, H, highlight=bad)) == [0, 1] and C.DIM not in {p["color"] for p in C.chart_primitives(OV, W, H, highlight=bad)}, bad
    # the renderers carry it: the PDF's content has the dim grey, the PNG renders
    pdf = C.render_pdf(hi, W, H)
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", pdf, re.S).group(1)).decode("latin-1")
    assert "0.816 0.816 0.816" in content, "the dim grey is not in the PDF"
    assert len(C.render_png(hi, W, H)) > 200, "the PNG did not render"
    # the hit test, on the plain primitives
    seg = of("polyline", plain)[1]["points"]
    (x1, y1), (x2, y2) = seg[0], seg[1]
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2
    length = ((x2 - x1) ** 2 + (y2 - y1) ** 2) ** 0.5
    nx, ny = -(y2 - y1) / length, (x2 - x1) / length                 # the segment's unit normal
    assert C.run_at(plain, mx, my) == 1, "a point on B's line"
    near, far = C.HIT_TOLERANCE_PX - 1, C.HIT_TOLERANCE_PX + 1.5
    assert C.run_at(plain, mx + nx * near, my + ny * near) == 1, "inside the tolerance"
    assert C.run_at(plain, mx + nx * far, my + ny * far) is None, "outside the tolerance"
    marker = [p for p in of("marker", plain) if p["run"] == 0 and p["y"] > 60][5]
    assert C.run_at(plain, marker["x"], marker["y"]) == 0, "a marker's centre"
    assert C.run_at(plain, marker["x"] + marker["size"] + near, marker["y"]) == 0, "the marker's rim plus a little"
    name = [p for p in of("text", plain) if p.get("run") == 1][0]
    assert C.run_at(plain, name["x"] + 10, name["y"]) == 1, "the legend's name"
    lm = [p for p in of("marker", plain) if p["run"] == 1 and p["y"] < 60][0]
    assert C.run_at(plain, lm["x"], lm["y"]) == 1, "the legend's marker"
    assert C.run_at(plain, W / 2, 10) is None and C.run_at(plain, W - 2, H - 2) is None, "empty area must be None"
    assert C.run_at(C.chart_primitives(I.overlay([]), W, H), 100, 100) is None, "an empty chart has no run"
check("a highlighted run: drawn last, heavier, its own colour, every other run and its legend entry dim, the same "
      "picture otherwise, the PDF and PNG carrying it; the hit test finds a run on its line (within the tolerance), "
      "its markers and its legend entry, and nothing on empty area", a_highlighted_run_in_the_primitives_and_the_hit_test)


def an_empty_chart_still_draws():
    prims = C.chart_primitives(I.overlay([]), W, H)
    assert prims[0]["kind"] == "background" and not of("marker", prims) and not of("polyline", prims), "prims[0][\"kind\"] == \"background\" and not of(\"marker\", prims) and not of(\"polyline\", prims)"
    assert [p["text"] for p in of("text", prims) if p["anchor"] == "e" and p["angle"] == 0] == ["-10", "0", "+10"], "[p[\"text\"] for p in of(\"text\", prims) if p[\"anchor\"] == \"e\" and p[\"angle\"] == 0] == [\"-10\", \"0\", \"+10\"]"
    assert len(C.render_pdf(prims, W, H)) > 200 and len(C.render_png(prims, W, H)) > 200, "len(C.render_pdf(prims, W, H)) > 200 and len(C.render_png(prims, W, H)) > 200"
check("an empty chart still draws its frame, and renders", an_empty_chart_still_draws)


def the_footnote_is_one_line_under_the_notes():
    line = "3 takes left out: their pitch did not match their note."
    with_note = C.chart_primitives(OV, W, H, title="Two horns", footnote=line)
    foot = [p for p in of("text", with_note) if p["text"] == line]
    assert len(foot) == 1 and foot[0]["anchor"] == "sw" and foot[0]["angle"] == 0, foot
    assert foot[0]["y"] == H - 4 and foot[0]["x"] < W / 4, foot[0]
    # the notes move up to make room; nothing else changes shape
    label_y = max(p["y"] for p in of("text", PRIMS) if p.get("angle") == 90)
    label_y_with = max(p["y"] for p in of("text", with_note) if p.get("angle") == 90)
    assert label_y_with < label_y - 10, (label_y_with, label_y)
    assert not [p for p in of("text", PRIMS) if p["anchor"] == "sw"], "a footnote with none asked for"
    assert len(C.render_png(with_note, W, H)) > 200, "len(C.render_png(with_note, W, H)) > 200"
    pdf = C.render_pdf(with_note, W, H)
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", pdf, re.S).group(1)).decode("latin-1")
    assert f"({line}) Tj" in content, "the footnote is not in the PDF"
    # and the PNG draws it where the primitive says: from the plot's left edge rightwards, in the
    # bottom strip (the first PNG put the text centred on that point, so a long line ran off the image)
    if C.PNG_AVAILABLE:
        import io
        from PIL import Image
        img = Image.open(io.BytesIO(C.render_png(with_note, W, H))).convert("L")
        strip = [(x, y) for y in range(H - 16, H) for x in range(W) if img.getpixel((x, y)) < 128]
        assert strip, "no footnote pixels in the bottom strip"
        left = min(x for x, _y in strip)
        assert abs(left - foot[0]["x"]) <= 3, f"the footnote starts at x={left}, the primitive at {foot[0]['x']}"
        assert max(x for x, _y in strip) < W - 4, "the footnote runs off the right edge"
check("the footnote (what was left out) is one plain line under the notes, in every renderer, drawn from the "
      "plot's left edge, only when asked for", the_footnote_is_one_line_under_the_notes)


def the_note_under_a_pixel():
    """Every point marker carries its note, and note_at finds the note whose column a pixel is in:
    the height does not matter, a pixel between columns or off the plot finds none, and the legend's
    markers (which carry no note) are never found."""
    prims = C.chart_primitives(OV, W, H, title="Two horns")
    points = [p for p in of("marker", prims) if "note" in p]
    legend = [p for p in of("marker", prims) if "note" not in p]
    assert points and legend, (len(points), len(legend))
    assert {p["note"] for p in points} == set(OV["notes"]), {p["note"] for p in points}
    for p in points:
        assert C.note_at(prims, p["x"], p["y"]) == p["note"]
        assert C.note_at(prims, p["x"] + 3, 5) == p["note"] == C.note_at(prims, p["x"] - 3, H - 5)
    xs = sorted({p["x"] for p in points})
    gap = min(b - a for a, b in zip(xs, xs[1:]))
    assert gap > 2 * C.NOTE_TOLERANCE_PX + 2, f"the fixture's columns are too close to test a miss: {gap}"
    assert C.note_at(prims, xs[0] + gap / 2, H / 2) is None, "a pixel between two columns found a note"
    assert C.note_at(prims, 1, H / 2) is None and C.note_at([], 10, 10) is None
check("every point carries its note; note_at finds the note of a pixel's column at any height, none between "
      "columns, never a legend marker", the_note_under_a_pixel)


print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
