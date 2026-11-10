"""The one line every "G-code written to a file" message ends with.

A user imported die insert G-code into LightBurn, which drops the speeds
and powers the file carries, and was confused by what came out. Matt: "is
there a dialog when the user generates gcode? maybe we can put something in
there (this is not for lightburn)", and "short message for gcode is best".
So every message box shown after G-code is written to a file ends with

    Send this file to the laser as it is. For LightBurn, use Generate SVG.

This suite clicks every Generate button in the app that ends in a message
box, on the Pad Maker and on Tooling, and pins:

  - every G-code one carries the line, as its last line
  - no SVG or STL one does
  - Frame & Cut, which writes no file, never shows it
  - the sentence exists once in the source (ui_dialogs.gcode_file_note) and
    is translated in all four catalogs

Self-skips without a display. Isolated profile (SAXSHOP_CONFIG_DIR).
"""
import gettext
import glob
import inspect
import os
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-gcode-note-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

results = []
NOTE = "Send this file to the laser as it is. For LightBurn, use Generate SVG."


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


def one_sentence_defined_once():
    import ui_dialogs
    assert ui_dialogs.gcode_file_note() == NOTE
    hits = []
    for path in glob.glob(os.path.join(REPO, "*.py")):
        text = open(path, encoding="utf-8").read()
        hits += [os.path.basename(path)] * text.count("Send this file to the laser as it is")
    assert hits == ["ui_dialogs.py"], f"the sentence should be written once, in ui_dialogs.py: {hits}"
    for lang in ("es", "de", "fr", "it"):
        tr = gettext.translation("saxshop", os.path.join(REPO, "locale"), languages=[lang])
        assert tr.gettext(NOTE) not in (NOTE, ""), f"not translated in {lang}"
        assert "LightBurn" in tr.gettext(NOTE) and "SVG" in tr.gettext(NOTE)


check("the sentence is defined once (ui_dialogs.gcode_file_note) and translated in all four catalogs",
      one_sentence_defined_once)


def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import falcon_sender
    import main as main_mod
    import ui_dialogs
    app = main_mod.PadSVGGeneratorApp(root)
    out = tempfile.mkdtemp(prefix="ssc-gcode-note-out-")
    boxes = []

    def fake_save(**kw):
        return os.path.join(out, kw.get("initialfile") or "out" + kw.get("defaultextension", ""))
    main_mod.filedialog.asksaveasfilename = fake_save       # tkinter.filedialog: the tabs share it
    main_mod.filedialog.askdirectory = lambda **k: out
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(main_mod.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))
    app._open_remaining_pads_window = lambda: None
    app._update_remaining_pads_window = lambda: None
    app._open_die_scrap_window = lambda: None
    # Everything a later check needs is set here, not inside an earlier
    # check: a check that fails half way must not leave a modal preview or
    # the kerf steps dialog waiting for a click nobody will make.
    app._show_kerf_instructions = lambda name: None
    app.fs_show_preview_var.set(False)

    def tooling_forms():
        app.die_scrap_var.set(False)
        app.die_units_var.set("mm")
        app.die_width_var.set("300")
        app.die_height_var.set("300")
        app.die_size_entry.delete("1.0", tk.END)
        app.die_size_entry.insert("1.0", "30, 35.5, 42")
        app.holder_variant_var.set("small")
        app.holder_layer_count_var.set(5)
        app.holder_units_var.set("mm")
        app.holder_width_var.set("300")
        app.holder_height_var.set("300")
        app.organizer_variant_var.set("lower")
        app.organizer_units_var.set("mm")
        app.organizer_width_var.set("300")
        app.organizer_height_var.set("400")
        app.kerf_material_var.set("Felt")
        app.kerf_use_existing_var.set(True)
        app.fs_show_preview_var.set(False)
        app.fs_material_var.set("Felt")
        app.crescent_list.clear()
        app.crescent_tonehole_var.set("43")
        app.crescent_percent_var.set("20")
        app.crescent_thickness_var.set("2.0")

    def said(fn):
        """The message boxes one click shows (info only; an error fails the check)."""
        del boxes[:]
        fn()
        root.update()
        errs = [b for b in boxes if b[0] == "showerror"]
        assert not errs, errs
        return [b[2] for b in boxes if b[0] == "showinfo"]

    def ends_with_note(msgs, what):
        assert len(msgs) == 1, f"{what}: {len(msgs)} boxes"
        assert msgs[0].endswith("\n\n" + NOTE), f"{what}: the box does not end with the laser line: {msgs[0][-120:]!r}"
        assert msgs[0].count(NOTE) == 1 and len(msgs[0]) > len(NOTE) + 10

    def no_note(msgs, what):
        assert msgs, f"{what}: no box at all (the check would prove nothing)"
        assert not [m for m in msgs if NOTE in m or "LightBurn, use" in m], f"{what}: carries the G-code line"

    # ---- Pad Maker ----------------------------------------------------------
    def pad_form(pads, materials, w=200, h=150):
        app.settings["units"] = "mm"
        app.settings["show_engraving_warning"] = False
        app.preview_var.set(False)
        app.custom_polygon = app.custom_polygon_outline = None
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", pads)
        app.width_entry.delete(0, tk.END)
        app.width_entry.insert(0, str(w))
        app.height_entry.delete(0, tk.END)
        app.height_entry.insert(0, str(h))
        app.filename_entry.delete(0, tk.END)
        app.filename_entry.insert(0, "note")
        for m, v in app.material_vars.items():
            v.set(m in materials)
        root.update()

    def pad_maker_standard():
        app.scrap_mode_var.set(False)
        pad_form("20 x 3\n30 x 2", ("felt", "card", "leather"))
        ends_with_note(said(app.on_generate_gcode), "Pad Maker Generate G-code")
        no_note(said(app.on_generate_svg), "Pad Maker Generate SVG")
        # All three wordings of the G-code box (plain, SD card ejected, please eject) end with it.
        src = inspect.getsource(main_mod.PadSVGGeneratorApp.on_generate_gcode)
        n = src.count('_("G-code files generated successfully.')
        assert n == 3, f"expected the three G-code success wordings, found {n}"
        for piece in src.split('_("G-code files generated successfully.')[1:]:
            assert "gcode_file_note()" in piece.split("messagebox.")[0], \
                "one of the three G-code success boxes does not end with the line"
    check("Pad Maker: the G-code box ends with the line (all three wordings); the SVG box does not",
          pad_maker_standard)

    def pad_maker_scrap_mode():
        try:
            for handler, is_gcode in ((app.on_generate_gcode, True), (app.on_generate_svg, False)):
                app._clear_scrap_session()
                pad_form("12.5 x 30\n20.0 x 6", ("felt",), w=70, h=60)
                app.scrap_mode_var.set(True)
                more = said(handler)                     # pads remain: the continue box
                (ends_with_note if is_gcode else no_note)(more, f"scrap continue ({'G-code' if is_gcode else 'SVG'})")
                assert "pads remaining" in more[0]
                last = None
                for _i in range(12):
                    last = said(handler)
                    if app._count_remaining_pads() == 0:
                        break
                assert app._count_remaining_pads() == 0 and "Session complete" in last[0], last
                (ends_with_note if is_gcode else no_note)(last, f"scrap session complete ({'G-code' if is_gcode else 'SVG'})")
        finally:
            app._clear_scrap_session()
            app.scrap_mode_var.set(False)
            app._update_scrap_status_display()
    check("Pad Maker, Scrap Mode: the G-code continue and session-complete boxes end with the line; the SVG ones do not",
          pad_maker_scrap_mode)

    # ---- Tooling ------------------------------------------------------------------
    def tooling_gcode_boxes():
        tooling_forms()
        # Every tool is clicked even when an earlier one is wrong, so one
        # missing line does not hide another.
        wrong = []

        def expect(kind, fn, what):
            try:
                (ends_with_note if kind == "gcode" else no_note)(said(fn), what)
            except AssertionError as e:
                wrong.append(str(e))
        expect("gcode", app._on_generate_die_gcode, "Die Inserts")
        expect("other", app._on_generate_die_svg, "Die Inserts SVG")
        expect("gcode", app._on_generate_holder_gcode, "Die Holders")
        expect("other", app._on_generate_holder_svg, "Die Holders SVG")
        expect("gcode", app._on_generate_organizer_gcode, "Die Organizer")
        expect("other", app._on_generate_organizer_svg, "Die Organizer SVG")
        expect("gcode", app._on_generate_kerf_gcode, "Kerf Test")
        assert said(app._on_generate_kerf_svg) == [], "the kerf SVG shows no Done box"
        expect("gcode", app._on_generate_feeds_speeds_gcode, "Speed & Power Test")
        expect("gcode", app._on_generate_crescent_gcode, "Tonehole Crescent")
        assert not wrong, "; ".join(wrong)
        no_note(said(app._on_generate_crescent_svg), "Tonehole Crescent SVG")
        no_note(said(app._on_generate_crescent_stl), "Tonehole Crescent STL")

        stl = sorted(f for f in os.listdir(os.path.join(REPO, "pad_press_spacers")) if f.lower().endswith(".stl"))[0]
        no_note(said(lambda: app._save_pad_spacer_stl(stl)), "Pad Press Spacer STL")
    check("Tooling: Die Inserts, Die Holders, Die Organizer, Kerf Test, Speed & Power and Crescent G-code boxes end with "
          "the line; every SVG and STL box does not", tooling_gcode_boxes)

    def die_scrap_mode():
        try:
            for handler, is_gcode in ((app._on_generate_die_gcode, True), (app._on_generate_die_svg, False)):
                app._clear_die_scrap_session()
                app.die_units_var.set("mm")
                app.die_width_var.set("60")
                app.die_height_var.set("60")
                app.die_size_entry.delete("1.0", tk.END)
                app.die_size_entry.insert("1.0", "30, 32")
                app.die_scrap_var.set(True)
                first, second = said(handler), said(handler)
                assert "dies remaining" in first[0] and "Session complete" in second[0], (first, second)
                for msgs, what in ((first, "sheet generated"), (second, "session complete")):
                    (ends_with_note if is_gcode else no_note)(msgs, f"die scrap {what} ({'G-code' if is_gcode else 'SVG'})")
        finally:
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            app._update_die_scrap_status()
    check("Tooling, Die Scrap Mode: both G-code boxes end with the line; both SVG boxes do not", die_scrap_mode)

    def frame_and_cut_never_shows_it():
        class Run:
            def __init__(self, parent, sender, lines, **kw):
                self._final_reason = "complete"

        class Sender:
            def __init__(self, port=None):
                pass

            def connect(self):
                pass

            def unlock(self):
                pass

            def disconnect(self):
                pass

        class Preview:
            def __init__(self, *a, **k):
                self.result = "save"
        keep = ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main_mod.NestingPreviewWindow
        ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main_mod.NestingPreviewWindow = Run, Sender, Preview
        app.falcon_port = "TEST"
        app.live_camera_var.set(False)
        continues = []
        real_continue = app._show_scrap_continue_dialog
        try:
            shown = []
            tooling_forms()
            # Tooling, including a die scrap session (its "Sheet Cut" boxes).
            app.die_size_entry.delete("1.0", tk.END)
            app.die_size_entry.insert("1.0", "30, 32")
            app.die_width_var.set("60")
            app.die_height_var.set("60")
            app.die_scrap_var.set(True)
            app._clear_die_scrap_session()
            shown += said(app._on_frame_cut_die) + said(app._on_frame_cut_die)
            assert len(shown) == 2 and "Cut 1 dies on sheet" in shown[0], shown
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            for fn in (app._on_frame_cut_holder, app._on_frame_cut_organizer, app._on_frame_cut_kerf,
                       app._on_frame_cut_feeds_speeds, app._on_frame_cut_crescent):
                shown += said(fn)
            # The Pad Maker in scrap mode: its continue dialog after a cut.
            app._show_scrap_continue_dialog = lambda *a: continues.append(a) or real_continue(*a)
            app._clear_scrap_session()
            pad_form("12.5 x 30\n20.0 x 6", ("felt",), w=70, h=60)
            app.scrap_mode_var.set(True)
            shown += said(app.on_frame_and_cut)
            assert len(continues) == 1 and len(continues[0]) == 3, \
                "Frame & Cut's continue dialog must not be told a G-code file was written"
            assert not [m for m in shown if NOTE in m], "Frame & Cut writes no file and must not show the line"
        finally:
            app._show_scrap_continue_dialog = real_continue
            ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main_mod.NestingPreviewWindow = keep
            app.falcon_port = None
            app._clear_scrap_session()
            app.scrap_mode_var.set(False)
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            app._update_die_scrap_status()
    check("Frame & Cut writes no file and never shows the line (Tooling, die scrap, and the Pad Maker's scrap dialog)",
          frame_and_cut_never_shows_it)

    def the_guide_says_why():
        # "Getting a job to the laser" > "G-code or SVG?" (the guide upgrade,
        # 2026-11-09, moved it out of the Pad Maker section).
        src = inspect.getsource(ui_dialogs.UserGuideWindow._section_laser)
        assert "G-code or SVG?" in src and "LightBurn" in src and "ignores those values" in src, \
            "the reason belongs in the user guide, not in the box"
        assert "ignores" not in NOTE
    check("the reason (LightBurn ignores the file's speeds and powers) is in the user guide, not in the box",
          the_guide_says_why)

    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
