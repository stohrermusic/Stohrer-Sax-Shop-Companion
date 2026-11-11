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
    assert OV["notes"][0] == "Bb3" and OV["notes"][-1] == "E5" and len(OV["notes"]) == 19
    assert PRIMS[0]["kind"] == "background" and PRIMS[0]["width"] == W and PRIMS[0]["height"] == H
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
    assert (lo, hi) == (-20, 20)                                   # cents -6..+12 and 0, a 5-cent margin, 10-cent steps
    assert ticks == [(f"{c:+.0f}" if c else "0") for c in range(lo, hi + 1, 10)], ticks
    grid = [p for p in of("line") if p["color"] in (C.GRID, C.ZERO_LINE)]
    assert len(grid) == 5 and sum(1 for p in grid if p["color"] == C.ZERO_LINE) == 1, "a line every 10 cents, the zero line solid"
    assert any(p["text"] == "cents" and p["angle"] == 90 for p in of("text"))
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
    assert not [p for p in of("line", one) if p.get("dash")]
check("the register change is one dashed vertical line, halfway between written C#5 and D5",
      the_register_change_is_dashed_between_c_sharp5_and_d5)


def flipped_puts_the_low_notes_on_the_right():
    prims = C.chart_primitives(OV, W, H, low_left=False)
    xs = notes_x(prims)
    assert sorted(xs, key=xs.get) == list(reversed(OV["notes"]))
    assert xs["Bb3"] > xs["E5"]
    dashed = [p for p in of("line", prims) if p.get("dash")]
    assert xs["D5"] < dashed[0]["x1"] < xs["C#5"]
check("the preference flips it: low notes on the right, the register line between the same two notes",
      flipped_puts_the_low_notes_on_the_right)


def one_colour_and_marker_per_run():
    markers = of("marker")
    polys = of("polyline")
    assert len(polys) == 2 and polys[0]["color"] == C.LINE_STYLES[0][0] and polys[1]["color"] == C.LINE_STYLES[1][0]
    a = [m for m in markers if m["color"] == C.LINE_STYLES[0][0]]
    b = [m for m in markers if m["color"] == C.LINE_STYLES[1][0]]
    assert {m["shape"] for m in a} == {"circle"} and {m["shape"] for m in b} == {"square"}
    assert len(a) == 19 + 1 and len(b) == 8 + 1, (len(a), len(b))        # the notes plus the legend's marker
    assert len(polys[0]["points"]) == 19 and len(polys[1]["points"]) == 8
    xs = notes_x(PRIMS)
    assert [round(x) for x, _y in polys[1]["points"]] == [round(xs[n]) for n in RUN_B.notes], "B's points sit on its notes"
    # the spread bar: Bb3 has two takes (4 and -6 cents): a vertical bar in A's colour, none for single takes
    bars = [p for p in of("line") if p["color"] == C.LINE_STYLES[0][0] and p["x1"] == p["x2"] and not p.get("dash")]
    assert len(bars) == 1 and abs(bars[0]["x1"] - xs["Bb3"]) < 1e-6, bars
    assert not [p for p in of("line") if p["color"] == C.LINE_STYLES[1][0] and p["x1"] == p["x2"] and p["y1"] != p["y2"]]
    # seven styles, then they cycle; five marker shapes
    assert len(C.LINE_STYLES) == 7 and C.style_for(7) == C.LINE_STYLES[0] and set(C.MARKERS) >= {s for _c, s in C.LINE_STYLES}
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
    assert C.y_range([]) == (-10, 10)
    assert C.y_range([3.0, -4.0]) == (-10, 10)
    assert C.y_range([12.0]) == (-10, 20)
    assert C.y_range([-37.0, 2.0]) == (-50, 10)
    assert C.y_range([55.0, 60.0]) == (-10, 70)      # 0 is always inside
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
    assert px[2, 2] == (255, 255, 255)
check("the PNG opens at the expected size and has the zero line and a marker where the primitives put them",
      the_png_is_the_same_picture)


def the_pdf_is_one_page_with_the_title():
    data = C.render_pdf(PRIMS, W, H, title="Two horns")
    assert data.startswith(b"%PDF-1.4") and data.rstrip().endswith(b"%%EOF")
    start = int(re.search(rb"startxref\n(\d+)", data).group(1))
    assert data[start:start + 4] == b"xref"
    rows = re.findall(rb"(\d{10}) (\d{5}) ([nf])", data[start:])
    offsets = [int(r[0]) for r in rows if r[2] == b"n"]
    for i, off in enumerate(offsets, start=1):
        assert data[off:].startswith(f"{i} 0 obj".encode()), (i, data[off:off + 12])
    assert data.count(b"/Type /Page ") == 1 and b"/Count 1" in data
    assert b"/Title (Two horns)" in data and b"/BaseFont /Helvetica" in data
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", data, re.S).group(1)).decode("latin-1")
    assert "(Two horns) Tj" in content and "(Mark VI, Link 7*) Tj" in content
    assert content.count("0 1 -1 0") == 19 + 1, "every note label and the axis label are turned 90 degrees"
    assert "[4 3] 0 d" in content, "the register line is dashed"
    assert "re f" in content and content.count(" c f Q") >= 20, "a background and the circle markers"
    # a parenthesis in a name is escaped
    ov = I.overlay([RUN_A])
    ov["lines"][0]["name"] = "Link (7*)"
    data = C.render_pdf(C.chart_primitives(ov, W, H), W, H)
    content = zlib.decompress(re.search(rb"stream\n(.*?)\nendstream", data, re.S).group(1)).decode("latin-1")
    assert "(Link \\(7*\\)) Tj" in content
check("the PDF is one page, every object at its xref offset, Helvetica, the title and the run names as text, "
      "the notes turned 90 degrees, the register line dashed", the_pdf_is_one_page_with_the_title)


def an_empty_chart_still_draws():
    prims = C.chart_primitives(I.overlay([]), W, H)
    assert prims[0]["kind"] == "background" and not of("marker", prims) and not of("polyline", prims)
    assert [p["text"] for p in of("text", prims) if p["anchor"] == "e" and p["angle"] == 0] == ["-10", "0", "+10"]
    assert len(C.render_pdf(prims, W, H)) > 200 and len(C.render_png(prims, W, H)) > 200
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
    assert len(C.render_png(with_note, W, H)) > 200
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


print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
