"""Fixes made before 3.0 for behaviour that was already in v2.9, and one for the suites.

  1. Frame & Cut for card on "Fit card to paper" with a shape still loaded.
     The nest is on the paper (the shape is ignored, as in the file export),
     but the framing pass traced the shape, the G92 came from the shape's
     anchor and the cut was flipped with the shape's height, so the cut
     landed off the sheet. A job that is not nested on the loaded shape now
     takes nothing from it: it streams exactly what it streams with no shape
     loaded, and the cut is the file Generate G-code writes.
  2. Cancelling at the jog step returned without disconnecting the sender,
     so the port stayed open until the object was collected. Every way out
     after a successful connect now disconnects, on both tabs.
  5. The no-laser return came after the sleep lock was taken and before the
     try whose finally releases it. No way out leaves the lock held.
  3. Die inserts nested with the Pad Maker's Edge Bias, whose default starts
     in the middle of the sheet. Dies always nest from the bottom-left
     corner as drawn, one rule for the SVG, the G-code, Frame & Cut, the fit
     check and Scrap Mode, and the Pad Maker's setting is never written to.
  4. Tooling's section buttons were a fixed 22 characters wide and cut off
     longer labels in German and Spanish. A button whose label needs more
     takes the width its label asks for; English is unchanged.

  6. No suite may put a native message box on the screen (tools/box_guard.py).
     A suite stood in for askyesno but not askokcancel, a planted bug made a
     handler ask a question, and the real box sat on screen until it was
     clicked; then the suite passed. The guard replaces all eight
     tkinter.messagebox functions, and a question no check expected fails
     the check that caused it.
  7. The lid question's buttons were a fixed 16 characters wide, which cut
     off the French label that matches its Start Frame button ("Oui,
     démarrer l'encadrement"). Both buttons now take the width the longer
     label needs; English is unchanged.

Mock laser throughout (a fake sender and fake run dialogs that record what
they were handed); nothing here touches a serial port. Nothing is sampled on
a timer. The section buttons are measured in a child process per language,
with the window mapped off screen. Isolated profile (SAXSHOP_CONFIG_DIR).
The die engine checks are headless; everything else self-skips without a
display.
"""
import copy
import json
import logging
import os
import re
import subprocess
import sys
import tempfile
import time
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
CHILD = sys.argv[2] if len(sys.argv) > 2 and sys.argv[1] == "--child" else None
_PROFILE = os.environ["SAXSHOP_CONFIG_DIR"] if CHILD else tempfile.mkdtemp(prefix="ssc-3-0-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
with open(os.path.join(_PROFILE, "app_settings.json"), "w") as _f:
    # The child shows the app in one language; the suite itself runs as
    # someone who has machine integration switched on.
    json.dump({"language": CHILD} if CHILD else {"experimental_machine_menu": True}, _f)

from i18n import init_translation  # noqa: E402
init_translation(CHILD or "en")

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


BIASES = ("center", "off", "n", "ne", "e", "se", "s", "sw", "w", "nw")
SEVEN_DIES = (30.0, 32.0, 34.0, 36.0, 38.0, 42.0, 44.0)     # five 50 mm rings, two 70 mm
G92_ZERO = "G92 X0.000 Y0.000"
FRAMING_S = {"felt": 12, "card": 14, "leather": 27, "acrylic": 19, "basswood": 21, "cork": 23}
FIRST_SECTION = {"en": "Die Inserts", "de": "Stanzeinsätze", "es": "Troqueles", "fr": "Matrices", "it": "Stampi"}


# ---------------------------------------------------------------------------
# Fix 4, measured in a child process: one language per process.
# ---------------------------------------------------------------------------
def child(lang):
    """The Tooling tab's section buttons in `lang`, as JSON: how wide each
    one really is and how wide a plain button with the same label asks to
    be, at the window's own size and again at the app's base 640 px."""
    import tkinter as tk
    import main as main_mod
    root = tk.Tk()
    root.geometry("+-4000+-4000")        # mapped for real sizes, on nobody's screen
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()                        # runs the after_idle tab fit
    app.notebook.select(app.tooling_tab_frame)

    def measure(width=None):
        # Wait for the layout itself, not for a number of update() calls: on
        # a busy machine the tab's first layout can arrive after two of them
        # (seen once as a 1 px button). Bounded, and it says so if it never
        # happens.
        deadline = time.time() + 60
        while True:
            root.update()
            laid_out = all(b.winfo_ismapped() and b.winfo_width() > 1 for b in app._tooling_buttons.values())
            if laid_out and width in (None, root.winfo_width()):
                break
            if time.time() > deadline:
                raise RuntimeError(f"the Tooling tab was never laid out (window {root.winfo_width()} px, wanted {width})")
            time.sleep(0.05)
        root.update()
        buttons = []
        for key, btn in app._tooling_buttons.items():
            text = str(btn.cget("text"))
            plain = tk.Button(btn.master, text=text, font=btn.cget("font"))
            buttons.append({
                "key": key, "text": text, "chars": int(btn.cget("width")),
                "width": btn.winfo_width(), "label_asks": plain.winfo_reqwidth(),
                "x": btn.winfo_rootx() - root.winfo_rootx(),
                "y": btn.winfo_rooty() - root.winfo_rooty(),
            })
            plain.destroy()
        return {"window": root.winfo_width(), "buttons": buttons}

    own = measure()
    root.geometry("640x720")             # the width the app starts from
    narrow = measure(width=640)

    # The lid question, built by the app's own code in this language. The
    # real dialog takes the grab and waits for an answer; this one is the
    # same class with those two steps taken out, under a withdrawn window,
    # so nothing shows and nothing waits.
    import builtins
    import ui_dialogs
    root.withdraw()
    root.update()
    built = []

    class Unanswered(ui_dialogs.LidClosedDialog):
        def grab_set(self):
            built.append(self)

        def wait_window(self, window=None):
            pass

    real = ui_dialogs.LidClosedDialog
    ui_dialogs.LidClosedDialog = Unanswered
    try:
        ui_dialogs.FalconRunDialog._remind_lid_closed(root)
    finally:
        ui_dialogs.LidClosedDialog = real
    lid = []
    for dlg in built:
        for frame in dlg.winfo_children():
            if not isinstance(frame, tk.Frame):
                continue
            for btn in frame.winfo_children():
                text = str(btn.cget("text"))
                plain = tk.Button(frame, text=text, font=btn.cget("font"))
                lid.append({"text": text, "chars": int(btn.cget("width")),
                            "asks": btn.winfo_reqwidth(), "label_asks": plain.winfo_reqwidth()})
                plain.destroy()
        dlg.destroy()
    print("REPORT " + json.dumps({"lang": lang, "own": own, "narrow": narrow, "lid": lid,
                                  "start_frame": builtins._("Start Frame →")}))
    root.destroy()


if CHILD:
    child(CHILD)
    sys.exit(0)


# ---------------------------------------------------------------------------
# Fix 3, engine level (headless).
# ---------------------------------------------------------------------------
def engine_checks():
    import config
    import svg_engine as se

    def settings_with(bias):
        s = copy.deepcopy(config.DEFAULT_SETTINGS)
        s["edge_bias"] = bias
        return s

    seven = [{"size": s, "qty": 1} for s in SEVEN_DIES]
    four = [{"size": s, "qty": 1} for s in SEVEN_DIES[:4]]
    mixed = [{"size": s, "qty": 1} for s in (30.0, 32.0, 42.0)]

    def dies_nest_from_the_corner_under_every_bias():
        nests = {}
        for bias in BIASES:
            s = settings_with(bias)
            placed, remaining, any_placed = se.try_nest_partial(seven, "die_ring", 130.0, 130.0, s)
            assert s["edge_bias"] == bias, f"the nest wrote edge_bias = {s['edge_bias']!r} into the settings"
            nests[bias] = (placed, remaining)
        placed, remaining = nests["sw"]
        for bias in BIASES:
            assert nests[bias] == nests["sw"], \
                f"edge bias {bias!r}: {len(nests[bias][0])} dies placed, {len(placed)} with 'sw': the layouts differ"
        assert len(placed) == 4, f"seven dies on 130 x 130 mm: {len(placed)} placed, 4 fit"
        assert sorted(p["size"] for p in remaining) == [38.0, 42.0, 44.0], remaining
        # Bottom-left as drawn: SVG y runs down, so the first die sits
        # against x = 0 and y = the sheet's height (the nester keeps its
        # 1 mm spacing from the edges).
        _size, cx, cy, r = placed[0]
        left, below = cx - r, 130.0 - (cy + r)
        assert 0 <= left <= 1.0 and 0 <= below <= 1.0, \
            f"the first die is {left} mm from the left edge and {below} mm from the bottom, not in that corner"
    check("dies: the same layout under every Edge Bias, from the bottom-left corner (130 x 130 mm takes 4 of 7)",
          dies_nest_from_the_corner_under_every_bias)

    def every_die_path_nests_that_layout():
        want = se._nest_discs(four, "die_ring", 130.0, 130.0, settings_with("sw"))[0]
        assert len(want) == 4
        with tempfile.TemporaryDirectory() as d:
            svgs = {}
            for bias in BIASES:
                s = settings_with(bias)
                assert se.can_all_pads_fit(four, "die_ring", 130.0, 130.0, s), \
                    f"edge bias {bias!r}: the fit check says four 50 mm dies do not fit 130 x 130 mm"
                assert se.can_all_pads_fit(mixed, "die_ring", 130.0, 130.0, s), \
                    f"edge bias {bias!r}: the fit check says two 50 mm dies and a 70 mm die do not fit 130 x 130 mm"
                assert se._nest_discs(four, "die_ring", 130.0, 130.0, s)[0] == want, bias
                assert se.nest_pads(four, "die_ring", 130.0, 130.0, s) == want, bias
                path = os.path.join(d, f"dies_{bias}.svg")
                assert se.generate_die_svg(four, 130.0, 130.0, path, s) == want, bias
                with open(path, encoding="utf-8") as f:
                    svgs[bias] = f.read()
                assert s["edge_bias"] == bias
            assert len(set(svgs.values())) == 1, "the die SVG differs with the Pad Maker's Edge Bias"
        # The traced-outline nester reads the same rule (dies are never cut
        # on a polygon in the app; one rule all the same).
        poly = [(0.0, 0.0), (140.0, 0.0), (140.0, 140.0), (0.0, 140.0)]
        on_poly = {b: se._nest_discs(four, "die_ring", 0, 0, settings_with(b), polygon=poly)[0] for b in BIASES}
        assert all(v == on_poly["sw"] for v in on_poly.values()), "the polygon nester still reads the Pad Maker's bias for dies"
    check("dies: the fit check, the plain nest, the preview nest and the SVG all use that one layout",
          every_die_path_nests_that_layout)

    def pads_still_follow_the_setting():
        for bias in BIASES:
            s = settings_with(bias)
            assert se.nest_edge_bias("die_ring", s) == "sw" == se.DIE_NEST_EDGE_BIAS
            for material in ("felt", "card", "leather", "exact_size", "adhesive_ring", "acrylic"):
                assert se.nest_edge_bias(material, s) == bias, (material, bias)
        s = settings_with("center")
        del s["edge_bias"]
        assert se.nest_edge_bias("felt", s) == "center", "a settings dict without the key means the default"
        pads = [{"size": 20.0, "qty": 6}, {"size": 35.5, "qty": 3}]
        layouts = {b: se._nest_discs(pads, "felt", 200.0, 150.0, settings_with(b))[0] for b in BIASES}
        assert all(len(v) == 9 for v in layouts.values())
        assert layouts["center"] != layouts["sw"] != layouts["ne"], "fixture: Edge Bias should move a felt nest"
        assert len({json.dumps(v) for v in layouts.values()}) >= 8, \
            "pads no longer follow the Pad Maker's Edge Bias"
        # Rings on a traced outline too.
        poly = [(0.0, 0.0), (200.0, 0.0), (200.0, 150.0), (0.0, 150.0)]
        a = se._nest_discs(pads, "leather", 0, 0, settings_with("center"), polygon=poly)[0]
        b = se._nest_discs(pads, "leather", 0, 0, settings_with("ne"), polygon=poly)[0]
        assert a != b, "fixture: Edge Bias should move a leather nest on a polygon"
    check("pads: every other material still nests with the Pad Maker's Edge Bias", pads_still_follow_the_setting)


# ---------------------------------------------------------------------------
# No native boxes: the guard itself, and the suites that must carry it.
# ---------------------------------------------------------------------------
GUARDED_SUITES = ("test_adhesive_ring_upgrade", "test_cork_material", "test_die_organizer_gcode",
                  "test_gcode_file_note", "test_mac_paths", "test_release_3_0", "test_tonehole_crescent_gcode",
                  "test_tooling_e2e", "test_tooling_frame_cut", "test_welcome")


def raises(fn, *words):
    """fn() must raise an AssertionError that mentions every word."""
    try:
        fn()
    except AssertionError as e:
        missing = [w for w in words if w not in str(e)]
        assert not missing, f"the failure does not mention {missing}: {e}"
        return
    raise AssertionError(f"nothing was raised (expected a failure mentioning {words})")


def box_guard_checks():
    import tkinter.commondialog as cd
    import tkinter.filedialog as fd
    import tkinter.messagebox as mb
    import tkinter.simpledialog as sd
    guard = box_guard.GUARD

    def stand_in(module, name):
        """module.name, but only if it is the guard's. These checks must
        never call the real thing: it would open on the screen and wait
        (it did, twice, when the guard was broken on purpose to test
        them). So: look first, call second."""
        fn = getattr(module, name)
        owner = getattr(fn, "__module__", None)
        assert owner == "box_guard", \
            f"{module.__name__}.{name} is not the guard's stand-in (it is {owner}'s); not calling it"
        return fn

    def all_eight_are_stand_ins():
        assert box_guard.ALL == ("showinfo", "showwarning", "showerror", "askokcancel", "askyesno",
                                 "askyesnocancel", "askquestion", "askretrycancel")
        for name in box_guard.ALL:
            stand_in(mb, name)
        # The function underneath them, and the one underneath every native
        # box and file dialog, which is what would finally open one.
        assert getattr(mb._show, "__self__", None) is guard, "tkinter.messagebox._show can still open a native box"
        stand_in(cd.Dialog, "show")
        for name in ("askopenfilename", "askopenfilenames", "asksaveasfilename", "askdirectory"):
            stand_in(fd, name)
        for name in ("askstring", "askinteger", "askfloat"):
            stand_in(sd, name)
    check("box guard: all eight message-box functions, the layers under them and the native dialogs are stood in for",
          all_eight_are_stand_ins)

    def a_question_nobody_expected_fails_the_check():
        del guard.shown[:], guard.asked[:]
        # show* boxes just record.
        assert stand_in(mb, "showinfo")("Done", "saved") == "ok" and stand_in(mb, "showerror")("Error", "no") == "ok"
        assert guard.shown == [("showinfo", "Done", "saved"), ("showerror", "Error", "no")]
        box_guard.settle()                       # nothing pending: no complaint
        # Each kind of question, unexpected: answered no / cancel, and settle() raises, once.
        for name, declined in (("askokcancel", False), ("askyesno", False), ("askyesnocancel", None),
                               ("askquestion", "no"), ("askretrycancel", False)):
            got = stand_in(mb, name)("Engraving Not Set", "so the crescents will not be numbered")
            assert got is declined or got == declined, f"{name} answered {got!r} with nobody expecting it"
            raises(box_guard.settle, "not expected by this check", name, "Engraving Not Set")
            box_guard.settle()                   # and it is forgotten afterwards
        # A dialog nobody stood in for is cancelled and counted the same way.
        assert stand_in(fd, "asksaveasfilename")(title="Save") == "" and stand_in(sd, "askstring")("Name", "?") is None
        raises(box_guard.settle, "asksaveasfilename", "askstring")
        # A box built underneath the eight never opens either...
        assert stand_in(mb, "_show")("Raw", "box", mb.QUESTION, mb.YESNO) == "no"
        raises(box_guard.settle, "NATIVE")
        # ...nor one built at the very bottom, as a dialog object.
        stand_in(cd.Dialog, "show")
        assert mb.Message(title="Raw", message="box").show() == ""
        raises(box_guard.settle, "NATIVE Message")
        assert fd.SaveAs(title="Save as").show() == ""
        raises(box_guard.settle, "NATIVE SaveAs")
    check("box guard: a question, dialog or raw box nobody expected is declined, and fails the check that caused it",
          a_question_nobody_expected_fails_the_check)

    def an_expected_question_is_answered_as_said():
        guard.expect("Unsaved changes", answer=False)
        guard.expect("Engraving Not Set", answer=True, kind="askokcancel")
        assert stand_in(mb, "askokcancel")("Engraving Not Set", "x") is True
        assert stand_in(mb, "askyesnocancel")("Unsaved changes", "y") is False
        box_guard.settle()
        guard.expect("Proceed?", answer=True)
        assert stand_in(mb, "askquestion")("Proceed?", "z") == "yes"
        box_guard.settle()
        # Another kind of box with that title is not the expected one.
        guard.expect("Engraving Not Set", answer=True, kind="askokcancel")
        assert stand_in(mb, "askyesno")("Engraving Not Set", "x") is False
        raises(box_guard.settle, "not expected by this check", "expected but never asked")
        # An expectation nobody used fails too, and does not leak into the next check.
        guard.expect("Never asked", answer=True)
        raises(box_guard.settle, "expected but never asked", "Never asked")
        box_guard.settle()
        # A check that failed for its own reason reports what was pending, and clears it.
        stand_in(mb, "askyesno")("Stray", "q")
        assert "Stray" in box_guard.reset() and box_guard.reset() == ""
    check("box guard: an expected question gets the answer its check named; a wrong or unused expectation fails",
          an_expected_question_is_answered_as_said)

    def the_suites_that_drive_the_app_carry_the_guard():
        for name in GUARDED_SUITES:
            with open(os.path.join(REPO, "tools", name + ".py"), encoding="utf-8") as f:
                src = f.read()
            head = src.split("\ndef ", 1)[0]
            assert "box_guard.install()" in head, f"{name} does not install the guard before its first function"
            body = src.split("def check(name, fn):", 1)[1].split("\ndef ", 1)[0]
            assert "box_guard.settle()" in body, f"{name}: check() does not fail on a question nobody expected"
            assert "box_guard.reset()" in body, f"{name}: a failed check leaves its questions for the next one"
            for fn in box_guard.ASKS:
                assert f".{fn} = lambda" not in src, (
                    f"{name} answers every {fn} with a lambda again; name the question with guard.expect()")
    check("box guard: the ten suites install it before anything else, settle in check(), and answer nothing blindly",
          the_suites_that_drive_the_app_carry_the_guard)


# ---------------------------------------------------------------------------
# Fix 4, the parent's side.
# ---------------------------------------------------------------------------
def section_button_checks():
    reports = {}

    def measured(lang):
        if lang not in reports:
            reports[lang] = measure_in_child(lang)
        return reports[lang]

    def measure_in_child(lang):
        folder = tempfile.mkdtemp(prefix=f"ssc-3-0-{lang}-")
        env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
        p = subprocess.run([sys.executable, os.path.abspath(__file__), "--child", lang],
                           capture_output=True, text=True, env=env, timeout=300, cwd=REPO, encoding="utf-8")
        line = [ln for ln in (p.stdout or "").splitlines() if ln.startswith("REPORT ")]
        assert p.returncode == 0 and line, (p.returncode, (p.stdout or "")[-300:], (p.stderr or "")[-800:])
        r = json.loads(line[0][7:])
        first = r["own"]["buttons"][0]["text"]
        assert first == FIRST_SECTION[lang], f"fixture: the child should be in {lang}, its first button reads {first!r}"
        return r

    for lang in ("en", "de", "es", "fr", "it"):
        def no_label_is_cut_off(lang=lang):
            r = measured(lang)
            for where in ("own", "narrow"):
                m = r[where]
                assert len(m["buttons"]) == 7
                if where == "narrow":
                    assert m["window"] == 640, f"fixture: the window should be 640 px wide, is {m['window']}"
                for b in m["buttons"]:
                    assert b["width"] > 1, f"fixture, {lang}: {b['text']!r} was measured before it was laid out"
                    assert b["width"] >= b["label_asks"], (
                        f"{lang}, window {m['window']} px: {b['text']!r} is on a {b['width']} px button, "
                        f"a button with that label asks for {b['label_asks']} px")
                    assert b["x"] >= 0 and b["x"] + b["width"] <= m["window"], (
                        f"{lang}, window {m['window']} px: {b['text']!r} runs from {b['x']} to "
                        f"{b['x'] + b['width']} px, past the window's edge")
        check(f"section buttons, {lang}: no label is cut off, at the window's own width and at 640 px",
              no_label_is_cut_off)

    for lang in ("en", "de", "es", "fr", "it"):
        def the_lid_question_fits(lang=lang):
            r = measured(lang)
            lid = r["lid"]
            assert len(lid) == 2, f"{lang}: the lid question has {len(lid)} buttons"
            yes, no = lid
            for b in lid:
                assert b["asks"] >= b["label_asks"], (
                    f"{lang}: {b['text']!r} is on a button that asks for {b['asks']} px, "
                    f"a button with that label asks for {b['label_asks']} px")
            assert yes["asks"] == no["asks"] and yes["chars"] == no["chars"], f"{lang}: the two buttons differ in width"
            assert yes["chars"] >= 16, f"{lang}: narrower than the 16 characters the buttons always had"
            # The yes button says what the Start Frame button it follows says.
            start = r["start_frame"].replace("→", "").strip()
            assert start and start.lower() in yes["text"].lower(), \
                f"{lang}: the lid question's {yes['text']!r} does not name the {start!r} button it follows"
            if lang == "en":
                assert [b["text"] for b in lid] == ["Yes, Start Frame", "No, Cancel"], lid
                assert yes["chars"] == 16, f"English buttons are 16 characters wide, as before: {yes['chars']}"
        check(f"lid question, {lang}: both labels fit their buttons, one width, and Yes names the Start Frame button",
              the_lid_question_fits)

    def english_is_as_it_was():
        b = measured("en")["own"]["buttons"]
        assert [x["text"] for x in b] == ["Die Inserts", "Die Holders", "Die Organizer", "Kerf Test",
                                          "Speed & Power Test", "Pad Press Spacers", "Tonehole Crescent"]
        assert [x["chars"] for x in b] == [22] * 7, f"English buttons are 22 characters wide: {[x['chars'] for x in b]}"
        assert len({x["width"] for x in b}) == 1, "English buttons are all one width"
        rows = [b[0:3], b[3:6], b[6:7]]
        assert all(len({x["y"] for x in row}) == 1 for row in rows), "three buttons to a row"
        ys = [row[0]["y"] for row in rows]
        assert ys[0] < ys[1] < ys[2], f"three rows, top to bottom: {ys}"
        for row in rows:
            assert [x["x"] for x in row] == sorted(x["x"] for x in row)
        assert [x["x"] for x in rows[0]] == [x["x"] for x in rows[1]], "the columns line up"
        assert rows[2][0]["x"] == rows[0][0]["x"]
    check("section buttons, English: the same seven buttons, 22 characters wide, three to a row, in order",
          english_is_as_it_was)


# ---------------------------------------------------------------------------
# Fixes 1, 2, 5 and the die handlers, on the live app against a mock laser.
# ---------------------------------------------------------------------------
def app_checks(root):
    import tkinter as tk
    import falcon_sender
    import gcode_engine as ge
    import main as main_mod
    import sleep_lock
    import svg_engine as se
    import tooling_tab
    import ui_dialogs

    main_mod.machine_supported = lambda: True
    app = main_mod.PadSVGGeneratorApp(root)
    assert app._machine_enabled(), "fixture: machine integration should be on in this profile"
    app.notebook.select(app.tooling_tab_frame)
    root.update()
    app.settings["laser_framing_power_by_material"].update(FRAMING_S)
    feed = app.settings.get("laser_framing_feed", 2000)

    out = tempfile.mkdtemp(prefix="ssc-3-0-out-")
    events, runs, boxes, saved, log_lines, previews = [], [], [], [], [], []
    answer = {"nest": "save", "fs_preview": "save"}

    def fake_save(**kw):
        path = os.path.join(out, kw.get("initialfile"))
        saved.append(path)
        return path

    tooling_tab.filedialog.asksaveasfilename = fake_save
    tooling_tab.filedialog.askdirectory = lambda **kw: out
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))

    class Grab(logging.Handler):
        def emit(self, record):
            msg = record.getMessage()
            if "frame&cut" in msg:
                log_lines.append(msg)
    logging.getLogger().addHandler(Grab())

    class FakeSender:
        def __init__(self, port=None):
            events.append("sender:new")

        def connect(self):
            events.append("sender:connect")

        def unlock(self):
            events.append("sender:unlock")

        def disconnect(self):
            events.append("sender:disconnect")

    class DeadSender(FakeSender):
        def connect(self):
            events.append("sender:connect-failed")
            raise OSError("port busy")

    class NoAnswerSender(FakeSender):
        def unlock(self):
            events.append("sender:unlock")
            raise OSError("no answer")

    class FakeRun:
        script = []      # how each dialog in turn ends; "complete" once it runs out

        def __init__(self, parent, sender, lines, **kw):
            how = FakeRun.script.pop(0) if FakeRun.script else "complete"
            if how == "raise":
                events.append("run:raised")
                raise RuntimeError("the dialog could not be built")
            self._final_reason = how
            runs.append({"lines": list(lines), "kw": dict(kw), "ends": how})
            events.append(f"run:{how}")

    class FakeFSPreview:
        def __init__(self, parent, pieces, w_mm, h_mm, **kw):
            self.result = answer["fs_preview"]

    class FakeNest:
        def __init__(self, parent, placements, w_mm, h_mm, **kw):
            previews.append({"placements": placements, "w": w_mm, "h": h_mm, **kw})
            self.result = answer["nest"]

    app._show_kerf_instructions = lambda name: None
    app._resolve_camera_index = lambda: None
    app._open_remaining_pads_window = lambda: None
    app._update_remaining_pads_window = lambda: None
    app._show_scrap_continue_dialog = lambda *a: None
    app._detect_falcon_async = lambda: events.append("detect-again")

    def machine(fn, script=(), sender=FakeSender, port="TEST"):
        """One click against the mock laser; the run dialogs it opened."""
        del events[:], runs[:], boxes[:], log_lines[:], previews[:]
        FakeRun.script = list(script)
        keep = (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender,
                tooling_tab.SpeedPowerTestPreview, main_mod.NestingPreviewWindow,
                sleep_lock.prevent_sleep, sleep_lock.allow_sleep)
        ui_dialogs.FalconRunDialog = FakeRun
        falcon_sender.FalconSender = sender
        tooling_tab.SpeedPowerTestPreview = FakeFSPreview
        main_mod.NestingPreviewWindow = FakeNest
        sleep_lock.prevent_sleep = lambda *a, **k: events.append("sleep:prevent")
        sleep_lock.allow_sleep = lambda *a, **k: events.append("sleep:allow")
        app.falcon_port = port
        try:
            fn()
            root.update()
        finally:
            (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender,
             tooling_tab.SpeedPowerTestPreview, main_mod.NestingPreviewWindow,
             sleep_lock.prevent_sleep, sleep_lock.allow_sleep) = keep
            FakeRun.script = []
            app.falcon_port = "TEST"
        return [dict(r) for r in runs]

    def errors():
        return [b for b in boxes if b[0] == "showerror"]

    def read_lines(path):
        with open(path) as f:           # the encoding the app reads its own G-code with
            return f.read().splitlines()

    # ---- forms -------------------------------------------------------------
    def pad_reset():
        app.custom_polygon = None
        app.custom_polygon_outline = None
        app._custom_polygon_lb_machine = None
        app._falcon_homed_this_session = False
        app._clear_scrap_session()
        app.scrap_mode_var.set(False)
        app._update_scrap_status_display()
        app.live_camera_var.set(False)
        app.card_paper_var.set(False)
        answer["nest"] = "save"

    def pad_form(pads_text, material, w=200, h=150, name="pads30"):
        app.settings["units"] = "mm"
        app.settings["show_engraving_warning"] = False
        app.settings["gcode_cut_grouping"] = "layer"
        app.preview_var.set(False)
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", pads_text)
        app.width_entry.delete(0, tk.END)
        app.width_entry.insert(0, str(w))
        app.height_entry.delete(0, tk.END)
        app.height_entry.insert(0, str(h))
        app.filename_entry.delete(0, tk.END)
        app.filename_entry.insert(0, name)
        for m, v in app.material_vars.items():
            v.set(m == material)
        root.update()

    def die_form(sizes="30.0, 35.5, 42.0", w="300", h="300", name="dies30"):
        app.die_scrap_var.set(False)
        app._clear_die_scrap_session()
        app.die_units_var.set("mm")
        app.die_width_var.set(w)
        app.die_height_var.set(h)
        app.die_size_entry.delete("1.0", tk.END)
        app.die_size_entry.insert("1.0", sizes)
        app.die_filename_var.set(name)

    def fs_form(preview):
        app.fs_material_var.set("Leather")
        app.fs_show_preview_var.set(preview)
        app.fs_filename_var.set("sp30")
        app.fs_diameter_var.set("20")
        app.fs_inner_diameter_var.set("0")

    # =======================================================================
    # Fix 1: card on the paper size takes nothing from a loaded shape
    # =======================================================================
    CARD_PADS = "20.0 x 4\n35.5 x 2"
    # 160 x 122 mm, tilted, so its anchor vertex is not its bounding box's corner.
    HAND_DRAWN = [(12.0, 0.0), (160.0, 10.0), (150.0, 122.0), (0.0, 110.0)]
    OUTLINE = [(8.0, 0.0), (190.0, 14.0), (182.0, 150.0), (0.0, 138.0)]
    INSET = [(13.0, 5.5), (184.5, 18.7), (177.3, 144.6), (5.4, 133.2)]
    LB_MACHINE = (123.4, 56.7)

    def card_on_paper_click(shape=None, outline=None, lb=None, homed=False, calibrated=False, scrap=False):
        """Frame & Cut for card on Letter paper with whatever shape state is asked for."""
        pad_reset()
        pad_form(CARD_PADS, "card", name="card30")
        app.card_paper_var.set(True)
        app.card_paper_dropdown.set("letter (8.5×11 in)")
        app.custom_polygon, app.custom_polygon_outline = shape, outline
        app._custom_polygon_lb_machine = lb
        app._falcon_homed_this_session = homed
        app.scrap_mode_var.set(scrap)
        keep_cal = app._camera_calibration_present
        if calibrated:
            app._camera_calibration_present = lambda: True
        try:
            got = machine(app.on_frame_and_cut)
            return got, list(log_lines), [dict(p) for p in previews], list(errors())
        finally:
            app._camera_calibration_present = keep_cal

    SHAPE_STATES = [
        ("a hand-drawn shape", dict(shape=HAND_DRAWN)),
        ("a camera outline with an inset, a machine position, homed, calibrated",
         dict(shape=INSET, outline=OUTLINE, lb=LB_MACHINE, homed=True, calibrated=True)),
    ]

    def card_on_paper_ignores_a_loaded_shape():
        try:
            base, base_log, base_prev, errs = card_on_paper_click()
            assert not errs and len(base) == 3, (errs, len(base))
            jog0, framing0, cut0 = base
            paper = (base_prev[0]["w"], base_prev[0]["h"])
            # Letter paper less its margins, and not the 200 x 150 typed in the sheet fields.
            assert 200 < paper[0] <= 215.9 and 260 < paper[1] <= 279.4, f"fixture: the nest is on {paper}"
            assert base_prev[0].get("polygon") is None, "fixture: no shape is loaded for the baseline"
            assert cut0["lines"][0] == G92_ZERO and framing0["lines"][0] == G92_ZERO
            for label, state in SHAPE_STATES:
                # The shape is really there, and would move things if it were used.
                app.custom_polygon, app.custom_polygon_outline = state["shape"], state.get("outline")
                assert app._frame_cut_anchor() != (0.0, 0.0), f"fixture, {label}: the anchor should be off the corner"
                got, log, prev, errs = card_on_paper_click(**state)
                assert not errs and len(got) == 3, (label, errs, len(got))
                jog, framing, cut = got
                assert prev[0].get("polygon") is None and (prev[0]["w"], prev[0]["h"]) == paper, \
                    f"{label}: the nest should be on the paper"
                assert framing["lines"] == framing0["lines"], \
                    f"{label}: the framing pass is not the one a sheet with no shape gets ({framing['lines'][:2]})"
                assert cut["lines"] == cut0["lines"], \
                    f"{label}: the cut is not the one a sheet with no shape gets (starts {cut['lines'][0]!r})"
                assert jog["kw"].get("auto_locate_target") is None, \
                    f"{label}: Try Auto Locate was offered a target on a job that is not on the shape"
                assert len(log) == 1 and "flip sheet" in log[0] and "polygon bbox none" in log[0] \
                    and "G92 X0.000 Y0.000" in log[0], log
                assert log == base_log
        finally:
            pad_reset()
    check("card on paper: with a shape loaded Frame & Cut frames, anchors, flips and cuts exactly as with none",
          card_on_paper_ignores_a_loaded_shape)

    def card_on_paper_in_scrap_mode_ignores_the_shape():
        try:
            base, _log, base_prev, errs = card_on_paper_click(scrap=True)
            assert not errs and len(base) == 3, (errs, len(base))
            for label, state in SHAPE_STATES:
                got, _log, prev, errs = card_on_paper_click(scrap=True, **state)
                assert not errs and len(got) == 3, (label, errs, len(got))
                assert prev[0].get("polygon") is None, f"scrap mode, {label}: the nest should be on the paper"
                assert got[1]["lines"] == base[1]["lines"], f"scrap mode, {label}: the framing pass follows the shape"
                assert got[2]["lines"] == base[2]["lines"], f"scrap mode, {label}: the cut follows the shape"
                assert got[2]["lines"][0] == G92_ZERO
                assert got[0]["kw"].get("auto_locate_target") is None
        finally:
            pad_reset()
    check("card on paper, Scrap Mode: a loaded shape changes nothing about what is streamed",
          card_on_paper_in_scrap_mode_ignores_the_shape)

    def the_card_on_paper_cut_is_the_exported_file():
        try:
            for label, state in SHAPE_STATES:
                got, _log, _prev, errs = card_on_paper_click(**state)
                assert not errs and len(got) == 3, (label, errs)
                cut = got[2]["lines"]
                # The same form through Generate G-code.
                target = os.path.join(out, "card30_card.gcode")
                if os.path.exists(target):
                    os.remove(target)
                machine(app.on_generate_gcode)
                assert not errors(), errors()
                assert os.path.isfile(target), sorted(os.listdir(out))
                assert cut == [G92_ZERO] + read_lines(target), \
                    f"{label}: what Frame & Cut streams is not the file Generate G-code writes"
                xs_ys = [ln for ln in cut[1:] if ln.startswith(("G0", "G1"))]
                ys =[float(v) for ln in xs_ys for v in re.findall(r"Y(-?\d+(?:\.\d+)?)", ln)]
                xs = [float(v) for ln in xs_ys for v in re.findall(r"X(-?\d+(?:\.\d+)?)", ln)]
                paper_w, paper_h = _prev[0]["w"], _prev[0]["h"]
                assert min(xs) >= 0 and min(ys) >= 0 and max(xs) <= paper_w and max(ys) <= paper_h, \
                    f"{label}: the cut leaves the paper: x {min(xs)}..{max(xs)}, y {min(ys)}..{max(ys)}"
        finally:
            pad_reset()
    check("card on paper: the streamed cut is the file Generate G-code writes, and it stays on the paper",
          the_card_on_paper_cut_is_the_exported_file)

    def a_job_on_the_shape_still_uses_the_shape():
        try:
            # Leather on a camera outline: traced, anchored, flipped with the outline, Auto Locate offered.
            pad_reset()
            pad_form("12 x 4\n18.5 x 3", "leather")
            app.custom_polygon, app.custom_polygon_outline = INSET, OUTLINE
            app._custom_polygon_lb_machine = LB_MACHINE
            app._falcon_homed_this_session = True
            keep_cal = app._camera_calibration_present
            app._camera_calibration_present = lambda: True
            try:
                got = machine(app.on_frame_and_cut)
                log = list(log_lines)
            finally:
                app._camera_calibration_present = keep_cal
            assert not errors() and len(got) == 3, (errors(), len(got))
            ax, ay = app._frame_cut_anchor()
            g92 = f"G92 X{ax:.3f} Y{ay:.3f}"
            assert (ax, ay) != (0.0, 0.0)
            assert got[1]["lines"] == [g92] + ge.generate_polygon_framing_gcode(
                OUTLINE, power_s=FRAMING_S["leather"], feed=feed), "the outline is no longer traced"
            assert got[2]["lines"][0] == g92, f"the cut starts {got[2]['lines'][0]!r}, the anchor is {g92!r}"
            assert got[0]["kw"].get("auto_locate_target") == LB_MACHINE
            assert "flip 150.0" in log[0], log
            # Card NOT on the paper size is on the shape like any other material.
            pad_reset()
            pad_form("12 x 4\n18.5 x 3", "card")
            app.custom_polygon = HAND_DRAWN
            got = machine(app.on_frame_and_cut)
            log = list(log_lines)
            assert not errors() and len(got) == 3, (errors(), len(got))
            ax, ay = app._frame_cut_anchor()
            g92 = f"G92 X{ax:.3f} Y{ay:.3f}"
            assert got[1]["lines"] == [g92] + ge.generate_polygon_framing_gcode(
                HAND_DRAWN, power_s=FRAMING_S["card"], feed=feed)
            assert got[2]["lines"][0] == g92 and "flip 122.0" in log[0], (got[2]["lines"][0], log)
        finally:
            pad_reset()
    check("a job that IS on the loaded shape still takes its frame, anchor, flip and Auto Locate target from it",
          a_job_on_the_shape_still_uses_the_shape)

    # =======================================================================
    # Fixes 2 and 5: every way out disconnects and releases the sleep lock
    # =======================================================================
    def pad_ready():
        pad_reset()
        pad_form("20 x 4\n30.5 x 3", "felt")

    def pad_empty():
        pad_reset()
        pad_form("", "felt")

    def pad_adjust():
        pad_ready()
        answer["nest"] = None

    def die_ready():
        die_form()

    def die_empty():
        die_form(sizes="")

    def fs_adjust():
        fs_form(preview=True)
        answer["fs_preview"] = None

    def tidy():
        answer["nest"], answer["fs_preview"] = "save", "save"
        pad_reset()
        die_form()
        fs_form(preview=False)

    # (what happens, set the form up, the button, run-dialog script, sender, port, reaches the laser)
    def ways_out(tab):
        pad = tab == "Pad Maker"
        ready, click = (pad_ready, app.on_frame_and_cut) if pad else (die_ready, app._on_frame_cut_die)
        return [
            ("no laser detected", ready, click, [], FakeSender, None, False),
            ("the form is refused", pad_empty if pad else die_empty, click, [], FakeSender, "TEST", False),
            ("the preview is answered Adjust", pad_adjust if pad else fs_adjust,
             click if pad else app._on_frame_cut_feeds_speeds, [], FakeSender, "TEST", False),
            ("the laser will not connect", ready, click, [], DeadSender, "TEST", False),
            ("the unlock raises", ready, click, [], NoAnswerSender, "TEST", True),
            ("the jog dialog raises", ready, click, ["raise"], FakeSender, "TEST", True),
            ("cancelled at the jog step", ready, click, ["cancelled"], FakeSender, "TEST", True),
            ("the framing dialog raises", ready, click, ["complete", "raise"], FakeSender, "TEST", True),
            ("framing stopped", ready, click, ["complete", "stopped"], FakeSender, "TEST", True),
            ("the cut dialog raises", ready, click, ["complete", "complete", "raise"], FakeSender, "TEST", True),
            ("cut stopped", ready, click, ["complete", "complete", "stopped"], FakeSender, "TEST", True),
            ("cut errored", ready, click, ["complete", "complete", "error"], FakeSender, "TEST", True),
            ("a finished cut", ready, click, [], FakeSender, "TEST", True),
        ]

    for tab in ("Pad Maker", "Tooling"):
        def every_way_out_disconnects(tab=tab):
            try:
                for what, setup, click, script, sender, port, connected in ways_out(tab):
                    setup()
                    machine(click, script=script, sender=sender, port=port)
                    ran = list(events)
                    tidy()
                    if not connected:
                        assert "sender:connect" not in ran, f"{tab}, {what}: fixture, the laser should not be reached: {ran}"
                        continue
                    assert ran.count("sender:connect") == 1, f"{tab}, {what}: {ran}"
                    assert ran.count("sender:disconnect") == 1, \
                        f"{tab}, {what}: the port was opened and {ran.count('sender:disconnect')} disconnects followed: {ran}"
                    after = ran[ran.index("sender:disconnect") + 1:]
                    assert not [e for e in after if e.startswith(("run:", "sender:"))], \
                        f"{tab}, {what}: the laser was used after the disconnect: {ran}"
            finally:
                tidy()
        check(f"{tab}: every way out after the port is open disconnects, once (jog cancel and errors included)",
              every_way_out_disconnects)

        def every_way_out_releases_the_sleep_lock(tab=tab):
            try:
                for what, setup, click, script, sender, port, connected in ways_out(tab):
                    setup()
                    machine(click, script=script, sender=sender, port=port)
                    ran = [e for e in events if e.startswith("sleep:")]
                    tidy()
                    held = 0
                    for e in ran:
                        held += 1 if e == "sleep:prevent" else -1
                        assert 0 <= held <= 1, f"{tab}, {what}: sleep lock calls out of order: {ran}"
                    assert held == 0, f"{tab}, {what}: the sleep lock is still held afterwards: {ran}"
                    if connected:
                        assert ran == ["sleep:prevent", "sleep:allow"], f"{tab}, {what}: {ran}"
            finally:
                tidy()
        check(f"{tab}: no way out leaves the sleep lock held (no laser, refusal, Adjust, cancel, stop, error, done)",
              every_way_out_releases_the_sleep_lock)

    def the_stream_is_untouched_by_the_cleanup():
        # What is streamed on a normal run, and in what order, is as it was.
        try:
            pad_ready()
            got = machine(app.on_frame_and_cut)
            order = [e for e in events if e.startswith(("sender:", "run:", "sleep:"))]
            assert order == ["sleep:prevent", "sender:new", "sender:connect", "sender:unlock", "run:complete",
                             "run:complete", "run:complete", "sender:disconnect", "sleep:allow"], order
            assert got[2]["lines"][0] == G92_ZERO and len(got[2]["lines"]) > 20
        finally:
            tidy()
    check("a finished run: lock, connect, unlock, jog, framing, cut, disconnect, release, in that order",
          the_stream_is_untouched_by_the_cleanup)

    # =======================================================================
    # Fix 3 through the real Die Inserts buttons
    # =======================================================================
    def settings_on_disk():
        path = os.path.join(_PROFILE, "app_settings.json")
        with open(path) as f:
            return json.load(f)

    def die_files_do_not_depend_on_the_pad_makers_edge_bias():
        seen = {}
        keep_bias = app.settings.get("edge_bias", "center")
        try:
            for bias in ("center", "off", "ne", "w", "sw"):
                app._set_edge_bias(bias)
                disk_before = settings_on_disk().get("edge_bias")

                def same_setting(after):
                    assert app.settings["edge_bias"] == bias and app.edge_bias_var.get() == bias, \
                        f"{after} changed the Pad Maker's Edge Bias from {bias!r} to {app.settings['edge_bias']!r}"
                    # Whatever is on disk is what was there, or the user's own setting if a save ran.
                    assert settings_on_disk().get("edge_bias") in (disk_before, bias), \
                        f"{after} saved another Edge Bias ({settings_on_disk().get('edge_bias')!r}) to disk"

                got = {}
                # Standard mode: four 50 mm dies fill a 130 x 130 mm sheet exactly as a 2 x 2 block.
                die_form(sizes="30, 32, 34, 36", w="130", h="130", name="b")
                del boxes[:]
                app._on_generate_die_svg()
                root.update()
                assert not errors(), f"edge bias {bias!r}, Generate SVG: {errors()}"
                with open(os.path.join(out, "b.svg"), encoding="utf-8") as f:
                    got["svg"] = f.read()
                same_setting("Generate SVG")
                del boxes[:]
                app._on_generate_die_gcode()
                root.update()
                assert not errors(), f"edge bias {bias!r}, Generate G-code: {errors()}"
                got["gcode"] = read_lines(os.path.join(out, "b.gcode"))
                same_setting("Generate G-code")
                runs_ = machine(app._on_frame_cut_die)
                assert not errors() and len(runs_) == 3, f"edge bias {bias!r}, Frame & Cut: {errors()}"
                got["framing"], got["cut"] = runs_[1]["lines"], runs_[2]["lines"]
                assert got["cut"] == [G92_ZERO] + got["gcode"], "Frame & Cut does not stream the exported die file"
                same_setting("Frame & Cut")

                # Scrap Mode: seven dies, the first 130 x 130 mm sheet takes four.
                for kind, handler, name in (("scrap svg", app._on_generate_die_svg, "s_scrap1.svg"),
                                            ("scrap gcode", app._on_generate_die_gcode, "s_scrap1.gcode")):
                    die_form(sizes="30, 32, 34, 36, 38, 42, 44", w="130", h="130", name="s")
                    app.die_scrap_var.set(True)
                    target = os.path.join(out, name)
                    if os.path.exists(target):
                        os.remove(target)
                    del boxes[:]
                    handler()
                    root.update()
                    if app.tooling_scrap_window is not None:
                        app.tooling_scrap_window.withdraw()
                    assert not errors(), f"edge bias {bias!r}, {kind}: {errors()}"
                    assert app._count_die_remaining() == 3, \
                        f"edge bias {bias!r}, {kind}: the first sheet took {7 - app._count_die_remaining()} dies, 4 fit"
                    with open(target, encoding="utf-8") as f:
                        got[kind] = f.read()
                    same_setting(kind)
                die_form(sizes="30, 32, 34, 36, 38, 42, 44", w="130", h="130", name="s")
                app.die_scrap_var.set(True)
                runs_ = machine(app._on_frame_cut_die)
                if app.tooling_scrap_window is not None:
                    app.tooling_scrap_window.withdraw()
                assert not errors() and len(runs_) == 3, f"edge bias {bias!r}, scrap Frame & Cut: {errors()}"
                got["scrap cut"] = runs_[2]["lines"]
                assert got["scrap cut"] == [G92_ZERO] + got["scrap gcode"].splitlines(), \
                    "scrap Frame & Cut does not stream the exported sheet"
                assert app._count_die_remaining() == 3
                same_setting("scrap Frame & Cut")
                seen[bias] = got
            for bias, got in seen.items():
                for kind in got:
                    assert got[kind] == seen["sw"][kind], \
                        f"die inserts, {kind}: the output under Edge Bias {bias!r} differs from 'sw'"
            # And it is the corner layout: the engine's own, asked for 'sw'.
            want = se._nest_discs([{"size": s, "qty": 1} for s in SEVEN_DIES[:4]], "die_ring", 130.0, 130.0,
                                  dict(app.settings, edge_bias="sw"))[0]
            with tempfile.TemporaryDirectory() as d:
                path = os.path.join(d, "want.gcode")
                ge.generate_die_gcode_from_placed(want, 130.0, 130.0, path, app._get_die_settings())
                assert seen["center"]["gcode"] == read_lines(path), "the die file is not the bottom-left corner layout"
            # Read off the file itself, whatever the engine calls its corner:
            # four 50 mm rings make a 2 x 2 block about 101 mm square, and on
            # a 130 mm sheet it sits against the origin (G-code Y runs up, so
            # that is the bottom-left as drawn), leaving the far 28 mm empty.
            moves = [ln for ln in seen["center"]["gcode"] if ln.startswith(("G0", "G1"))]
            xs = [float(v) for ln in moves for v in re.findall(r"X(-?\d+(?:\.\d+)?)", ln)]
            ys = [float(v) for ln in moves for v in re.findall(r"Y(-?\d+(?:\.\d+)?)", ln)]
            assert 0 <= min(xs) and 0 <= min(ys) and max(xs) < 104 and max(ys) < 104, \
                f"the die block is not against the origin: x {min(xs)}..{max(xs)}, y {min(ys)}..{max(ys)}"
            assert max(xs) > 99 and max(ys) > 99, "fixture: four 50 mm rings should span about 101 mm each way"
        finally:
            app._set_edge_bias(keep_bias)
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            app._update_die_scrap_status()
            die_form()
    check("die inserts: SVG, G-code, Frame & Cut and Scrap Mode give the same files under every Edge Bias, "
          "and never change it", die_files_do_not_depend_on_the_pad_makers_edge_bias)

    def the_default_edge_bias_no_longer_refuses_dies_that_fit():
        keep_bias = app.settings.get("edge_bias", "center")
        try:
            app._set_edge_bias("center")            # what a new install has
            # Two 50 mm rings and a 70 mm one fit 130 x 130 mm from a corner.
            # From the middle the 70 mm ring went down first, dead centre,
            # and nothing else fitted round it: the job was refused.
            die_form(sizes="30, 32, 42", w="130", h="130", name="fit")
            del boxes[:]
            app._on_generate_die_gcode()
            root.update()
            assert not errors(), f"two 50 mm dies and a 70 mm die on 130 x 130 mm were refused: {errors()}"
            info = [b for b in boxes if b[0] == "showinfo"]
            assert len(info) == 1 and "3 die rings" in info[0][2], boxes
        finally:
            app._set_edge_bias(keep_bias)
            die_form()
    check("die inserts: with the default Edge Bias (center) three dies that fit a 130 x 130 mm sheet are cut, not refused",
          the_default_edge_bias_no_longer_refuses_dies_that_fit)


def main():
    print("Fixes made before 3.0")
    print("=" * 60)
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    engine_checks()
    box_guard_checks()

    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping the app checks: no display available ({e})")
        root = None
    if root is not None:
        root.withdraw()
        try:
            app_checks(root)
        finally:
            try:
                root.destroy()
            except tk.TclError:
                pass
        section_button_checks()

    print("=" * 60)
    passed = sum(results)
    print(f"{passed}/{len(results)} passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
