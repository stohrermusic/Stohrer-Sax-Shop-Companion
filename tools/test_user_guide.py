"""The User Guide after its upgrade (2026-11-09).

Matt: "let's give the user guide a bit of an upgrade- keep it concise, but
have it be a little more of a guide, vs just blocks of description. have a
TOC that is clickable on the right, and have it open to full screen. put
example graphics where appropriate."

Pinned here, on the real window (off screen), the real app and the source:
  - it opens maximized (a normal window filling the screen) and is not
    transient; nothing is scheduled (no timer races a timer)
  - the contents list: ten sections in order, every entry backed by a mark
    in the text whose line holds that entry's title, Tooling's entries named
    exactly as its buttons; a click on any entry puts its heading at the top
    of the view (or, where the end of the text stops it, on screen) and
    selects it; the highlight follows the scroll
  - every figure draws from the app's own geometry (the nest figure places
    every pad it was given, the die figure carries the engravings), in a
    canvas with explicit colours, every label through _(); each figure once
  - search covers the whole guide whatever section it opened at
  - Help > User Guide opens at the current tab's section; the tour's stop
    and the Quick Start Guide's last line open it too
  - the text: shorter than the old guide, every string a plain literal, no
    empty string, the Feature Set tick named as the dialog names it, every
    safety and correctness point still there, "G-code or SVG?" in place
  - in German (a child process): it builds, the figure captions and the
    headings are translated
Isolated profile (SAXSHOP_CONFIG_DIR). GUI checks self-skip without a display.
"""
import ast
import inspect
import json
import os
import re
import subprocess
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-user-guide-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

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


TOOL_BUTTONS = ["Die Inserts", "Die Holders", "Die Organizer", "Kerf Test", "Speed & Power Test",
                "Pad Press Spacers", "Tonehole Crescent"]
SECTION_TITLES = ["Pad Maker", "Getting a job to the laser", "Tooling", "Key Height Library", "Serial Lookup",
                  "Screw Specs", "Tuner", "Analyzer", "Sharing data", "Learn to make pads"]
MUST_SAY = [
    "Is the lid closed?",                               # the lid question, on Start Frame
    "low power is still power",                         # framing is a live laser
    "keep your hands out of the bed",
    "flesh",                                            # locator marks: flesh side up
    "starting point, not tested",                       # adhesive ring defaults
    "Kerf = hole diameter minus disc diameter",         # the Kerf Test's measuring step
    "LightBurn asks for half",
    "Test on the material you will really cut",
    "DESIGNED BY PHIL NOY", "\"NOY\"", "noysaxophonesupplies.com",
    "Direct machine control (GRBL)",                    # the Feature Set tick, as the dialog names it
    "Camera #1 could not be opened; showing camera #2.",
    "ignores those values",                             # G-code or SVG?
    "Only one \"max\" line per sheet",
    "bottom-left corner, largest first",                # dies nest by rows
]
MUST_NOT_SAY = ["Experimental: machine integration", "Show Full Guide"]


# ---------------------------------------------------------------- the source
def guide_strings():
    """Every _() literal in UserGuideWindow, (method name, text)."""
    import ui_dialogs
    src = inspect.getsource(ui_dialogs.UserGuideWindow)
    tree = ast.parse(src)
    out = []
    for fn in ast.walk(tree):
        if not isinstance(fn, ast.FunctionDef):
            continue
        for node in ast.walk(fn):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id == "_":
                assert len(node.args) == 1, f"{fn.name}: _() with {len(node.args)} arguments"
                a = node.args[0]
                assert isinstance(a, ast.Constant) and isinstance(a.value, str), \
                    f"{fn.name}: _() on something that is not a plain string: {ast.unparse(a)}"
                out.append((fn.name, a.value))
    return out


def the_text_is_shorter_and_plain():
    strings = guide_strings()
    words = sum(len(t.split()) for _f, t in strings)
    assert 2500 <= words <= 5000, f"{words} words in the guide (the old guide had 7,513; the rewrite must be shorter, not empty)"
    assert not [t for _f, t in strings if not t.strip()], "an empty string"
    long = [t[:50] for _f, t in strings if len(t.split()) > 60]
    assert not long, f"strings over 60 words (one paragraph, one string, short): {long}"
    # Every section is a method of its own, in the guide's order, and each starts its own contents entry.
    import ui_dialogs
    src = inspect.getsource(ui_dialogs.UserGuideWindow._insert_content)
    order = re.findall(r"self\.(_section_\w+)\(\)", src)
    assert order == ["_section_pad_maker", "_section_laser", "_section_tooling", "_section_key_heights",
                     "_section_serial_lookup", "_section_screw_specs", "_section_tuner", "_section_toner",
                     "_section_sharing", "_section_learn"], order
    for name in order:
        body = inspect.getsource(getattr(ui_dialogs.UserGuideWindow, name))
        assert re.search(r'self\._section\("[a-z_]+", _\(', body), f"{name} does not open with its section heading"


check("the text: shorter than the old guide, every string a plain literal, none empty or over 60 words, "
      "ten sections in order", the_text_is_shorter_and_plain)


def it_says_what_it_must():
    whole = "\n".join(t for _f, t in guide_strings())
    missing = [s for s in MUST_SAY if s not in whole]
    assert not missing, f"the guide no longer says: {missing}"
    stray = [s for s in MUST_NOT_SAY if s in whole]
    assert not stray, f"the guide still says: {stray}"
    # The tick it names is the one the Feature Set dialog has.
    import main as main_mod
    assert 'Direct machine control (GRBL)' in inspect.getsource(main_mod), "the Feature Set tick was renamed"
    # G-code or SVG? sits in the laser section with its reason (test_gcode_file_note reads it too).
    import ui_dialogs
    laser = inspect.getsource(ui_dialogs.UserGuideWindow._section_laser)
    assert '"G-code or SVG?"' in laser and "LightBurn" in laser and "ignores those values" in laser


check("the guide keeps every safety and correctness point, names the Feature Set tick as the dialog does, "
      "and no longer quotes a label that does not exist", it_says_what_it_must)


def figure_labels_go_through_gettext():
    import ui_dialogs
    src = inspect.getsource(ui_dialogs.UserGuideWindow)
    tree = ast.parse(src)
    figs = [n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef) and n.name.startswith("_fig_")]
    assert len(figs) == len(ui_dialogs.UserGuideWindow.FIGURE_NAMES), [f.name for f in figs]
    # Every string with a space in it (a sentence, a caption) is the argument
    # of _(), whichever variable it travels through; the engraved credit is
    # the one literal that is not UI text.
    allowed = {"DESIGNED BY PHIL NOY"}
    bad = []
    for fn in figs:
        translated = {id(n.args[0]) for n in ast.walk(fn)
                      if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == "_" and n.args}
        docstring = fn.body[0].value if fn.body and isinstance(fn.body[0], ast.Expr) \
            and isinstance(fn.body[0].value, ast.Constant) else None
        for node in ast.walk(fn):
            if node is docstring:
                continue
            if isinstance(node, ast.Constant) and isinstance(node.value, str) and " " in node.value:
                if id(node) not in translated and node.value not in allowed:
                    bad.append(f"{fn.name}: {node.value!r}")
    assert not bad, f"figure text not through _(): {bad}"
    # And every figure is placed exactly once in the text.
    placed = re.findall(r'self\._figure\("(\w+)"\)', src)
    assert sorted(placed) == sorted(ui_dialogs.UserGuideWindow.FIGURE_NAMES), placed


check("every figure label goes through _(), and each figure is placed exactly once", figure_labels_go_through_gettext)


# ---------------------------------------------------------------- the window
def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import ui_dialogs
    import main as main_mod

    def open_guide(section=None):
        g = ui_dialogs.UserGuideWindow(root, section=section)
        g.geometry("+-4000+-4000")        # on nobody's screen; keeps its zoomed state
        root.update()
        return g

    def line_of(g, index):
        return int(g.text.index(index).split(".")[0])

    def top_line(g):
        return line_of(g, "@0,0")

    g = open_guide()

    def opens_maximized_not_transient_and_schedules_nothing():
        # The rule in the window manager's own terms, not in pixels of the
        # screen: the window is in the 'zoomed' state (Windows, X11), or,
        # where there is no such state (macOS Aqua), it is as large as Tk
        # says a window may be. `wm maxsize` with no size set by the script
        # is the usable screen: on Aqua the visible frame (less the menu
        # bar and the Dock) less the window frame's decorations. The first
        # version asked for `screen - 100` of height, and CI's Macs, which
        # lose 122 and 138 px to the menu bar, title bar and Dock, failed
        # it with the window filling everything it could (e2293e4, run
        # 42755511492). The evidence is in the message.
        sw, sh = g.winfo_screenwidth(), g.winfo_screenheight()
        w, h = g.winfo_width(), g.winfo_height()
        try:
            state = g.state()
        except tk.TclError:
            state = "?"
        max_w, max_h = (int(v) for v in g.wm_maxsize())
        evidence = (f"state {state!r}, window {w}x{h}, Tk allows {max_w}x{max_h}, screen {sw}x{sh}, "
                    f"{root.tk.call('tk', 'windowingsystem')}")
        assert state == "zoomed" or (w >= 0.9 * max_w and h >= 0.8 * max_h), f"not maximized: {evidence}"
        print(f"    ({evidence})")
        assert g.wm_transient() == "", f"transient to {g.wm_transient()!r}: no maximize button on Windows"
        assert str(g.title()) == "User Guide" and g.opened_at is None
        assert not root.tk.call("after", "info"), "the guide scheduled something"
        assert str(g.text.cget("state")) == "disabled"
        buttons = [w for w in g.winfo_children() for b in [w] if isinstance(b, tk.Frame)]
        assert buttons, "no button row"
    check("opens maximized (or screen-filling on Aqua), not transient, read-only, nothing scheduled",
          opens_maximized_not_transient_and_schedules_nothing)

    def the_contents_list_matches_the_text():
        tops = g.toc.get_children("")
        assert [g.toc.item(i, "text") for i in tops] == SECTION_TITLES, [g.toc.item(i, "text") for i in tops]
        assert list(g._toc_items) == list(g.SECTION_KEYS)
        items = [i for t in tops for i in (t, *g.toc.get_children(t))]
        marks = dict((item, mark) for mark, item in g._marks)
        assert set(items) == set(marks), "an entry with no mark, or a mark with no entry"
        assert [it for _m, it in g._marks] == items, "the marks are not in the contents list's order"
        for item in items:
            mark = marks[item]
            line = g.text.get(f"{mark} linestart", f"{mark} lineend").strip()
            assert line == g.toc.item(item, "text"), f"{g.toc.item(item, 'text')!r} marks the line {line!r}"
        tooling = g._toc_items["tooling"]
        subs = [g.toc.item(i, "text") for i in g.toc.get_children(tooling)]
        assert subs[:7] == TOOL_BUTTONS, f"Tooling's entries are not its buttons: {subs}"
        for key in ("pad_generator", "laser", "tooling", "toner"):
            assert len(g.toc.get_children(g._toc_items[key])) >= 4, key
        # Marks are in reading order, and strictly increasing.
        lines = [line_of(g, m) for m, _i in g._marks]
        assert lines == sorted(lines) and len(set(lines)) == len(lines), lines
    check("the contents list: ten sections in order, every entry backed by a heading on its own line, "
          "Tooling's entries named as its buttons", the_contents_list_matches_the_text)

    def a_click_on_any_entry_goes_there_and_none_is_dead():
        dead = []
        for mark, item in g._marks:
            g._go_to_item(item)
            root.update()
            sel = g.toc.selection()
            at_top = top_line(g) == line_of(g, mark)
            at_bottom = g.text.yview()[1] >= 0.999
            if not (at_top or (at_bottom and g.text.bbox(mark) is not None)):
                dead.append(g.toc.item(item, "text"))
            assert sel and sel[0] == item, f"{g.toc.item(item, 'text')!r}: selected {sel}"
        assert not dead, f"entries that do not scroll to their heading: {dead}"
        # A real click: the mouse on an entry's row.
        g.scroll_to("pad_generator")
        root.update()
        target = g._toc_items["tooling"]
        g.toc.see(target)
        root.update()
        bx, by, bw, bh = g.toc.bbox(target)
        g.toc.event_generate("<ButtonPress-1>", x=bx + 10, y=by + bh // 2)
        g.toc.event_generate("<ButtonRelease-1>", x=bx + 10, y=by + bh // 2)
        root.update()
        assert top_line(g) == line_of(g, "sec_tooling"), "a click on Tooling did not scroll there"
        assert g.toc.selection() == (target,)
    check("a click on any contents entry puts its heading at the top of the view (or on screen at the end), "
          "selects it, and none is dead", a_click_on_any_entry_goes_there_and_none_is_dead)

    def the_highlight_follows_the_scroll():
        for key in g.SECTION_KEYS:
            g.text.yview(f"sec_{key}")
            root.update()
            sel = g.toc.selection()
            # A short last section cannot reach the top: its entry, or the one whose heading is at the top.
            assert sel and sel[0] in (g._toc_items[key], g._heading_at_top()), key
            if g.text.yview()[1] < 0.999:
                assert sel[0] == g._toc_items[key], f"{key}: selected {g.toc.item(sel[0], 'text')!r}"
        # Scrolled a little into a section: still that section (or its first sub-heading once it passed the top).
        g.text.yview("sec_tooling")
        g.text.yview_scroll(2, "units")
        root.update()
        sel = g.toc.selection()[0]
        assert sel == g._toc_items["tooling"] or g.toc.parent(sel) == g._toc_items["tooling"], g.toc.item(sel, "text")
        # The very top (title and introduction): the first section's entry.
        g.text.yview_moveto(0.0)
        root.update()
        assert g.toc.selection() == (g._toc_items["pad_generator"],)
    check("the highlighted entry follows the scroll position", the_highlight_follows_the_scroll)

    def the_figures_draw_from_real_geometry():
        from svg_engine import get_disc_diameter
        from config import DEFAULT_SETTINGS
        names = [n for n, _c in g._figures]
        assert sorted(names) == sorted(g.FIGURE_NAMES) and len(names) == len(set(names)), names
        for name, cv in g._figures:
            assert str(cv.cget("bg")) == g.FIGURE_BG, f"{name}: canvas background {cv.cget('bg')!r}"
            assert int(cv.cget("width")) == g.FIGURE_W
            items = cv.find_all()
            assert len(items) >= 4, f"{name}: {len(items)} items"
            for item in items:
                kind = cv.type(item)
                if kind == "text":
                    assert cv.itemcget(item, "fill"), f"{name}: a text item with no explicit colour"
                    assert cv.itemcget(item, "text").strip(), f"{name}: an empty text item"
                elif kind in ("oval", "rectangle", "polygon"):
                    assert cv.itemcget(item, "outline") or cv.itemcget(item, "fill"), f"{name}: a shape with no colour"
                elif kind == "line":
                    assert cv.itemcget(item, "fill"), f"{name}: a line with no colour"
        figs = dict(g._figures)
        # The nest figure placed every pad it was given (6 + 3 + 2) and drew the zone box.
        nest = figs["nest_zone"]
        ovals = [i for i in nest.find_all() if nest.type(i) == "oval"]
        assert len(ovals) == 11, f"nest figure: {len(ovals)} discs, expected 11"
        assert [i for i in nest.find_all() if nest.type(i) == "rectangle"], "nest figure: no zone box"
        # The pad figure: three discs whose relative sizes are the sizing rules' (leather > felt > card).
        pad = figs["pad_discs"]
        widths = sorted(abs(pad.coords(i)[2] - pad.coords(i)[0]) for i in pad.find_all() if pad.type(i) == "oval")
        d_felt, d_card = get_disc_diameter(14.0, "felt", DEFAULT_SETTINGS), get_disc_diameter(14.0, "card", DEFAULT_SETTINGS)
        assert len(widths) == 2 and abs(widths[1] / widths[0] - d_felt / d_card) < 0.02, widths
        assert [i for i in pad.find_all() if pad.type(i) == "line"], "pad figure: no darted leather outline"
        # The die figure carries the engravings: many stroke lines (the size, NOY, DESIGNED BY PHIL NOY).
        die = figs["die_ring"]
        assert len([i for i in die.find_all() if die.type(i) == "line"]) >= 25, "die figure: engravings missing"
        # The crescent is a filled polygon with many points, inside the tonehole circle.
        cres = figs["crescent"]
        polys = [i for i in cres.find_all() if cres.type(i) == "polygon"]
        assert polys and len(cres.coords(polys[0])) > 60
    check("every figure draws from the app's own geometry, in explicit colours; the nest figure places every pad, "
          "the pad figure keeps the sizing rules' ratios, the die figure carries the engravings",
          the_figures_draw_from_real_geometry)

    def the_reading_column_is_bounded():
        width = g.text.winfo_width()
        rmargin = int(g.text.tag_cget("col", "rmargin"))
        if width > g.TEXT_MAX_PX + 2 * g.TEXT_PADX:
            assert width - 2 * g.TEXT_PADX - rmargin == g.TEXT_MAX_PX, (width, rmargin)
        else:
            assert rmargin == 0, (width, rmargin)
        assert g.text.tag_ranges("col"), "the column tag is not on the text"
    check("the reading column never runs edge to edge on a wide window", the_reading_column_is_bounded)

    def search_covers_the_whole_guide():
        g2 = open_guide("tuner")
        try:
            assert g2.opened_at == "tuner" and top_line(g2) == line_of(g2, "sec_tuner")
            g2._search_var.set("crescent")
            n = g2.text.get("1.0", "end").lower().count("crescent")
            assert n >= 3 and len(g2._search_hits) == n and g2._search_status_var.get() == f"1 of {n}"
            g2._search_var.set("G-code or SVG?")
            assert len(g2._search_hits) == 1
            assert str(g2.text.cget("state")) == "disabled"
        finally:
            g2.destroy()
    check("search covers the whole guide whatever section it opened at", search_covers_the_whole_guide)

    g.destroy()

    # ---- through the app ----
    real_supported = main_mod.machine_supported
    main_mod.machine_supported = lambda: True
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()
    opened = []
    real_cls = main_mod.UserGuideWindow

    class Recorded(ui_dialogs.UserGuideWindow):
        def __init__(self, parent, section=None):
            super().__init__(parent, section=section)
            self.geometry("+-4000+-4000")
            self.cursor_put_in_search = 0
            opened.append(self)

        def focus_search(self):
            self.cursor_put_in_search += 1
            return super().focus_search()
    main_mod.UserGuideWindow = Recorded

    def help_opens_at_the_current_tab():
        try:
            for tab, key in ((app.tooling_tab_frame, "tooling"), (app.key_tab, "key_heights"), (app.pad_tab, "pad_generator")):
                app.notebook.select(tab)
                root.update()
                del opened[:]
                app.open_user_guide()
                root.update()
                assert len(opened) == 1 and opened[0].opened_at == key, (key, [o.opened_at for o in opened])
                w = opened[0]
                assert top_line(w) == line_of(w, f"sec_{key}"), f"{key}: the heading is not at the top"
                assert w.toc.selection() == (w._toc_items[key],)
                w.destroy()
            # The Welcome window's search link: the whole guide, cursor in
            # the search box. What the window did, not what the X server
            # granted: under xvfb there is no window manager, no Tk window
            # ever gets the keyboard focus, and focus_get() is None on every
            # window (Linux CI, e2293e4). Tk itself records where a window's
            # focus goes (focus -lastfor), window manager or not.
            del opened[:]
            w = app.open_user_guide_search()
            root.update()
            assert isinstance(w, Recorded) and w.opened_at is None
            assert w.cursor_put_in_search == 1 and w.focus_lastfor() == w.search_entry, \
                (f"the cursor is not in the search box: focus_search called {w.cursor_put_in_search} time(s), "
                 f"the window's focus goes to {w.focus_lastfor()}")
            w.destroy()
        finally:
            del opened[:]
    check("Help > User Guide opens the whole guide at the current tab's section; the Welcome link at the top "
          "with the cursor in the search box", help_opens_at_the_current_tab)

    def the_quick_start_guides_last_line_opens_the_guide():
        win = app.open_quick_start()
        try:
            root.update()
            assert win.guide_link is not None, "the last line is not a link"
            assert str(win.guide_link.cget("text")) == "The rest is in Help > User Guide."
            assert win.guide_link is win.line_labels[-1]
            assert str(win.guide_link.cget("cursor")) == "hand2"
            del opened[:]
            assert win.guide_link.bind("<Button-1>"), "the last line is not bound to a click"
            win._open_user_guide()              # what the click does
            root.update()
            assert len(opened) == 1 and opened[0].opened_at is None, "the link did not open the whole guide"
            opened[0].destroy()
            # Without a callback (a plain window) the line is a plain label.
            plain = ui_dialogs.QuickStartWindow(root, show_frame_cut=False)
            assert plain.guide_link is None
            plain.destroy()
        finally:
            win.destroy()
            del opened[:]
    check("the Quick Start Guide's last line opens the User Guide", the_quick_start_guides_last_line_opens_the_guide)

    def the_tour_stop_opens_at_the_pad_maker():
        src = inspect.getsource(main_mod.run_tour)
        assert 'UserGuideWindow(root, section="pad_generator")' in src, "the tour's stop changed"
        w = Recorded(root, section="pad_generator")
        root.update()
        assert top_line(w) == line_of(w, "sec_pad_generator")
        w.destroy()
        del opened[:]
    check("the tour's User Guide stop still opens the guide at the Pad Maker", the_tour_stop_opens_at_the_pad_maker)

    main_mod.UserGuideWindow = real_cls
    main_mod.machine_supported = real_supported
    try:
        root.destroy()
    except tk.TclError:
        pass


gui()


# ---------------------------------------------------------------- in German, in a child
CHILD = r'''
import json, os, sys
sys.path.insert(0, sys.argv[1])
from i18n import init_translation
init_translation(sys.argv[2])
import tkinter as tk
import ui_dialogs
root = tk.Tk(); root.withdraw()
g = ui_dialogs.UserGuideWindow(root, section="tooling")
g.geometry("+-4000+-4000")
root.update()
out = {"sections": [g.toc.item(i, "text") for i in g.toc.get_children("")],
       "tooling": [g.toc.item(i, "text") for i in g.toc.get_children(g._toc_items["tooling"])],
       "figures": {n: [cv.itemcget(i, "text") for i in cv.find_all() if cv.type(i) == "text"] for n, cv in g._figures},
       "top_ok": g.text.index("@0,0").split(".")[0] == g.text.index("sec_tooling").split(".")[0],
       "words": len(g.text.get("1.0", "end").split())}
g.destroy(); root.destroy()
print("REPORT " + json.dumps(out))
'''


def in_german():
    folder = tempfile.mkdtemp(prefix="ssc-guide-de-")
    env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
    reports = {}
    for lang in ("en", "de"):
        p = subprocess.run([sys.executable, "-c", CHILD, REPO, lang], capture_output=True, text=True, env=env,
                           timeout=180, cwd=REPO, encoding="utf-8")
        line = [ln for ln in (p.stdout or "").splitlines() if ln.startswith("REPORT ")]
        assert p.returncode == 0 and line, (lang, p.returncode, (p.stdout or "")[-300:], (p.stderr or "")[-800:])
        reports[lang] = json.loads(line[0][7:])
    en, de = reports["en"], reports["de"]
    assert de["top_ok"], "German: the Tooling heading is not at the top"
    assert len(de["sections"]) == 10 and de["sections"] != en["sections"], de["sections"]
    assert de["tooling"][:7] != en["tooling"][:7], "German Tooling entries are still English"
    for name in en["figures"]:
        en_texts, de_texts = en["figures"][name], de["figures"][name]
        assert len(en_texts) == len(de_texts), name
        # Numbers stay numbers; the captions with words in them are translated
        # (a word the two languages share, "Laser", "LightBurn", may stay).
        lettered = [(a, b) for a, b in zip(en_texts, de_texts) if re.search(r"[A-Za-z]{3,}", a)]
        same = [a for a, b in lettered if a == b]
        assert lettered and len(same) <= len(lettered) // 2, f"{name}: untranslated captions {same}"
    assert de["words"] > 1000


check("in German (a child process): the guide builds, opens at the section asked for, and its headings and "
      "figure captions are translated", in_german)

print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
