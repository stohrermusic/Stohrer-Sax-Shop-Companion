"""The Analyzer tab (2026-11-10): the Toner tab turned round so the intonation chart
comes first and the harmonic analyzer stands behind the HARMONICS switch.

What is pinned, on a profile that has NEVER unlocked the harmonics (so the tab's own
migration and the switch's terms run as a new user sees them):
  - the one-time migration (config.migrate_analyzer_tab) from each starting state:
    never unlocked (the old startup rule had hidden the tab), unlocked and visible,
    unlocked and hidden on purpose, a fresh profile; run once only; in memory at the
    build (which writes nothing) and on disk with the next regular save
  - the tab is called Analyzer, shows by default, opens on the chart with the
    HARMONICS switch off and the MODE switch at free (labels free / guided), with
    Compare... and Export... under the chart; Feature Set lists it as Analyzer
  - the HARMONICS switch: declined terms put it back to off and unlock nothing;
    accepted terms unlock (saved) and show the spectrum view, which says
    "experimental"; once unlocked the terms are not asked again
  - the required fields both ways: horn type, make, model, player and mouthpiece
    for an intonation run; the mic fields only while HARMONICS is on, and the
    "Mic Type Required" stop at Capture likewise
  - a free run and a guided run through the synthetic source (the real engine
    and animate loop), each producing a chart with the notes played
  - the free readout: the frame count a capture would keep, turning green at the
    34th frame (the shortest note kept), and a note let go sooner named too short
  - Compare...: sessions picked from two presets draw together, legend and markers;
    Clear goes back to the loaded preset
  - Export...: a PNG (opens, twice the chart's size) and a one-page PDF with the title
  - the Settings dialog's "Low notes on the left" tick, saved and flipping the axis
  - the tour's stop for the tab
  - the window (Matt, 2026-11-10: "have it resize itself"): the start-up width is the
    old rule's (tab labels and the Tooling tab, never the Analyzer's strip); selecting
    the tab widens a narrower window to the strip (Capture on the page), never past the
    screen; leaving puts the width back; a size the user chose while on the tab, a
    window already wide enough and a maximized window are left alone
  - the HARMONICS switch remembers its position across launches, but "on" in a profile
    that never unlocked counts as off
  - in every language, in a child process each: the tab builds, the start-up width is
    not the strip's, the select widens and the leave restores; German labels translated
Nothing waits a fixed time for the window manager: each geometry change is awaited
with a bounded poll under the mainloop (settle), so Aqua's late geometry cannot fail a
check. Frames run under a real mainloop with a quit timer (see test_tuner_canvas for
why a root.update() loop can't be used). Self-skips without a display or the audio stack.
"""
import json
import os
import re
import struct
import subprocess
import sys
import tempfile
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PROFILE = tempfile.mkdtemp(prefix="ssc-analyzer-test-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
# A user who never saw the harmonics terms: nothing about the toner in the profile.
with open(os.path.join(_PROFILE, "app_settings.json"), "w", encoding="utf-8") as _f:
    json.dump({"toner_record_wav": False, "seen_calibration_tutorial": True, "capture_threshold": 5}, _f)
from i18n import init_translation  # noqa: E402
init_translation("en")
import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

F0 = 233.08             # concert A#3
HARM = [0.0, -3.0, -6.0, -10.0, -14.0, -18.0, -22.0]


def _capture(note, f0, **extra):
    return {"note": note, "fundamental_freq": f0, "harmonics_db": list(HARM),
            "harmonic_cents": [0.0] * len(HARM), "timestamp": "2026-11-01 10:00:00", "n_frames": 30,
            "method": "free", "signal_level": 0.5, "spectral_centroid": 900.0, **extra}


# Two presets with captures: Bench (Tenor, one session, a mic) and Other (Alto, two
# sessions, no mic anywhere: the backfill leaves it without one).
SEEDED = {
    "My Presets": {
        "Bench": {"horn_type": "Tenor", "horn_make": "Selmer", "horn_model": "Mark VI", "serial": "", "player": "Matt",
                  "mouthpiece": "Link 7*", "reed": "", "ligature": "", "mic_position": "", "room": "", "preamp": "",
                  "mic_type": "condenser", "mic_model": "AT2020", "notes": "", "created": "2026-11-01",
                  "sessions": [{"date": "2026-11-01 10:00:00", "mic_type": "condenser", "mic_model": "AT2020",
                                "horn_type": "Tenor", "captures": [_capture("A#3", 233.5), _capture("C4", 261.0)]}]},
        "Other": {"horn_type": "Alto", "horn_make": "Conn", "horn_model": "6M", "player": "Matt", "mouthpiece": "Meyer 5",
                  "notes": "", "created": "2026-10-01",
                  "sessions": [{"date": "2026-10-01 10:00:00", "captures": [_capture("A#3", 232.5)]},
                               {"date": "2026-10-02 10:00:00", "captures": [_capture("D4", 294.5), _capture("E4", 329.0)]}]},
        # a preset from the Toner's first weeks: its labels sit two transpositions (28 semitones on a
        # tenor) below the sounding pitch, the storage mistake found in Matt's library (2026-11-10)
        "Old": {"horn_type": "Tenor", "horn_make": "Selmer", "horn_model": "SBA", "player": "Matt", "mouthpiece": "Link",
                "notes": "", "created": "2026-03-23",
                "sessions": [{"date": "2026-03-23 11:56:47", "captures": [
                    _capture("A#0", 146.83, method="calibration"), _capture("C1", 164.81, method="calibration"),
                    _capture("D1", 185.0, method="calibration"), _capture("F1", 220.0, method="calibration")]}]},
    }
}
with open(os.path.join(_PROFILE, "toner_data.json"), "w", encoding="utf-8") as _f:
    json.dump(SEEDED, _f, indent=1)

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


class FakeClock:
    """Stands in for toner_tab's `time`: time() under the test's control, the rest real."""

    def __init__(self, start=1_000_000.0):
        self.t = start

    def time(self):
        return self.t

    def strftime(self, fmt, when=None):
        return time.strftime(fmt, time.localtime(self.t) if when is None else when)

    def __getattr__(self, name):
        return getattr(time, name)


def frame(f0, note, level=0.8, harm=HARM):
    from toner_engine import HarmonicInfo, TonerResult
    r = TonerResult()
    r.fundamental_freq = f0
    r.fundamental_note = note
    r.harmonics = [HarmonicInfo(i + 1, f0 * (i + 1), f0 * (i + 1), db, 0.0) for i, db in enumerate(harm)]
    r.signal_level = level
    r.spectral_centroid = 900.0
    return r


def silence():
    from toner_engine import TonerResult
    return TonerResult()


def settings_on_disk():
    with open(os.path.join(_PROFILE, "app_settings.json"), encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------- the migration, on dicts
def migration_from_each_state():
    from config import migrate_analyzer_tab
    # never unlocked: the old startup rule had forced the tab hidden; it comes back
    s = {"toner_unlocked": False, "visible_tabs": {"Toner": False, "Tuner": True}}
    assert migrate_analyzer_tab(s) is True
    assert s["visible_tabs"]["Toner"] is True and s["visible_tabs"]["Tuner"] is True and s["analyzer_tab_shown_v3"] is True
    # unlocked and visible: nothing to change but the flag
    s = {"toner_unlocked": True, "visible_tabs": {"Toner": True}}
    assert migrate_analyzer_tab(s) is True and s["visible_tabs"]["Toner"] is True and s["analyzer_tab_shown_v3"]
    # unlocked and hidden: the user hid it on purpose, and that stands
    s = {"toner_unlocked": True, "visible_tabs": {"Toner": False}}
    assert migrate_analyzer_tab(s) is True and s["visible_tabs"]["Toner"] is False and s["analyzer_tab_shown_v3"]
    # a fresh profile: the default (visible) applies; the flag is set
    s = {}
    assert migrate_analyzer_tab(s) is True and "Toner" not in s["visible_tabs"] and s["analyzer_tab_shown_v3"]
    # once only: a second run changes nothing and says so
    s = {"toner_unlocked": False, "visible_tabs": {"Toner": False}, "analyzer_tab_shown_v3": True}
    before = json.dumps(s, sort_keys=True)
    assert migrate_analyzer_tab(s) is False and json.dumps(s, sort_keys=True) == before


def main():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    from toner_engine import AUDIO_AVAILABLE
    if not AUDIO_AVAILABLE:
        print("Skipping: audio libraries not installed (numpy/sounddevice) — no Analyzer tab on this build")
        root.destroy()
        return 0

    print("The Analyzer tab")
    print("=" * 60)
    check("the migration from each starting state, once only", migration_from_each_state)

    import config
    import toner_tab
    import intonation_chart
    from toner_engine import FREE_MIN_FRAMES, SAMPLE_RATE, whole_window, load_tone_presets, note_to_freq
    from config import TONER_DATA_FILE
    check("profile override is in effect",
          lambda: os.path.normcase(config.get_config_dir()) == os.path.normcase(_PROFILE))

    from main import PadSVGGeneratorApp
    root.geometry("+0+0")                  # the app's own size: the window-fit rule is under test too
    app = PadSVGGeneratorApp(root)
    root.update()
    assert getattr(app, "_toner_engine", None) is not None, "Analyzer tab not built (audio libs missing?)"
    app._toner_mic_checked = True          # the first-start mic notice is a modal
    app._toner_show_coverage_summary = lambda: None    # a Toplevel with buttons; not under test here

    def run(ms):
        root.after(ms, root.quit)
        root.mainloop()

    def settle(pred, timeout=6.0, what="the window"):
        """Run the mainloop in slices until pred() holds (a geometry change applied by the
        window manager, late on Aqua) or the bound passes; never a fixed wait."""
        deadline = time.time() + timeout
        while not pred() and time.time() < deadline:
            run(100)
        assert pred(), f"{what} did not settle within {timeout:.0f} s"

    W = {}      # the window's sizes as the checks find them

    def toplevels():
        return [w for w in root.winfo_children() if isinstance(w, tk.Toplevel)]

    def widgets(top, cls):
        out = []

        def walk(w):
            for c in w.winfo_children():
                if isinstance(c, cls):
                    out.append(c)
                walk(c)
        walk(top)
        return out

    def chart_texts():
        cv = app._toner_chart_canvas
        return [cv.itemcget(i, "text") for i in cv.find_withtag("chart") if cv.type(i) == "text"]

    def chart_markers():
        cv = app._toner_chart_canvas
        return [cv.type(i) for i in cv.find_withtag("chart") if cv.type(i) in ("oval", "polygon")]

    # ---------------------------------------------------------------- the tab as built
    def the_tab_as_built():
        assert app.notebook.tab(app.toner_tab_frame, "text") == "Analyzer"
        assert str(app.toner_tab_frame) in app.notebook.tabs()
        assert app.notebook.tab(app.toner_tab_frame, "state") == "normal", "the tab is hidden on a fresh profile"
        assert app.settings.get("toner_unlocked") in (None, False)
        # the migration ran in memory; building the app writes nothing (the flag goes to
        # disk with the next regular save: the terms' accept below is one)
        assert app.settings.get("analyzer_tab_shown_v3") is True, "the build did not run the migration"
        assert "analyzer_tab_shown_v3" not in settings_on_disk(), "building the app wrote the settings file"
        assert app._toner_harmonics_var.get() == "off" and app._toner_mode_var.get() == "free"
        # the start-up width is the old rule's, not the strip's: the window is narrower than the
        # strip needs (the Pad Maker is the first tab; nothing has selected the Analyzer yet)
        root.update_idletasks()
        W["start"] = (root.winfo_width(), root.winfo_height())
        W["need"] = app._analyzer_strip_need_px()
        assert W["start"][0] < W["need"], f"the window opened {W['start'][0]} px wide, the strip's {W['need']}"
        assert getattr(app, "_analyzer_widened", None) is None
        app.notebook.select(app.toner_tab_frame)
        # selecting the tab widens the window to the strip, height unchanged, never past the screen
        want_w = min(W["need"], root.winfo_screenwidth())
        settle(lambda: root.winfo_width() == want_w, what="the widening on select")
        W["widened"] = root.winfo_width()
        assert root.winfo_height() == W["start"][1], (root.winfo_height(), W["start"][1])
        assert app._analyzer_widened == (W["start"][0], want_w), app._analyzer_widened
        assert app._toner_chart_frame.winfo_ismapped() and not app._toner_harmonics_frame.winfo_ismapped()
        assert app._toner_chart_canvas.winfo_width() > 300 and app._toner_chart_canvas.cget("bg") == "#FFFFFF"
        # the whole control strip is on the page: the Capture button, packed last and rightmost
        strip, btn = app._toner_ctrl_frame, app._toner_capture_btn
        settle(lambda: strip.winfo_reqwidth() <= app._toner_main_frame.winfo_width(), what="the page's width")
        assert btn.winfo_ismapped(), "the Capture button is off the page"
        assert btn.winfo_rootx() + btn.winfo_width() <= root.winfo_rootx() + root.winfo_width()
        texts = [w.cget("text") for w in widgets(app._toner_main_frame, tk.Label)]
        for label in ("HARMONICS", "off", "on", "MODE", "free", "guided"):
            assert label in texts, f"{label!r} not on the control strip: {texts}"
        assert "cal" not in texts
        btns = [w.cget("text") for w in widgets(app._toner_chart_frame, tk.Button)]
        assert btns == ["Compare...", "Export..."], btns
        # nothing loaded: the chart says what to do
        assert any("Load a preset" in t for t in chart_texts()), chart_texts()
        app._toner_stop()
    check("the tab is called Analyzer, shows on a never-unlocked profile (the migration recorded on disk), opens on "
          "the chart with HARMONICS off and MODE at free (labels free/guided), Compare... and Export... under it",
          the_tab_as_built)

    def feature_set_lists_the_analyzer():
        before = set(toplevels())
        app._open_feature_set()
        dlg = [w for w in toplevels() if w not in before][0]
        texts = [c.cget("text") for c in widgets(dlg, tk.Checkbutton)]
        dlg.destroy()
        assert "Analyzer" in texts and "Toner" not in texts and "Tuner" in texts, texts
    check("Feature Set offers the tab as Analyzer, in the main list", feature_set_lists_the_analyzer)

    # ---------------------------------------------------------------- the window, case by case
    def width_is(w):
        return lambda: root.winfo_width() == w

    def leaving_puts_the_width_back():
        start_w, start_h = W["start"]
        app.notebook.select(app.pad_tab)
        settle(width_is(start_w), what="the width going back on leaving")
        assert root.winfo_height() == start_h and getattr(app, "_analyzer_widened", None) is None
        # and again: select widens, leave restores
        app.notebook.select(app.toner_tab_frame)
        settle(width_is(W["widened"]), what="the second widening")
        app.notebook.select(app.pad_tab)
        settle(width_is(start_w), what="the second restore")
    check("leaving the tab puts the width back to what it was, height unchanged", leaving_puts_the_width_back)

    def a_size_the_user_chose_stays():
        start_w, start_h = W["start"]
        app.notebook.select(app.toner_tab_frame)
        settle(width_is(W["widened"]), what="the widening")
        chosen = W["widened"] + 40
        root.geometry(f"{chosen}x{start_h}")              # the user drags the window wider while on the tab
        settle(width_is(chosen), what="the user's resize")
        app.notebook.select(app.pad_tab)
        run(400)
        assert root.winfo_width() == chosen, f"the app took back a width the user chose: {root.winfo_width()}"
        assert getattr(app, "_analyzer_widened", None) is None
        root.geometry(f"{start_w}x{start_h}")
        settle(width_is(start_w), what="the reset for the next check")
    check("a size the user chose while on the tab is left alone on leaving", a_size_the_user_chose_stays)

    def a_window_already_wide_enough_is_left_alone():
        start_w, start_h = W["start"]
        wide = W["need"] + 60
        root.geometry(f"{wide}x{start_h}")
        settle(width_is(wide), what="the wide window")
        app.notebook.select(app.toner_tab_frame)
        run(400)
        assert root.winfo_width() == wide and getattr(app, "_analyzer_widened", None) is None
        assert app._toner_capture_btn.winfo_ismapped()
        app.notebook.select(app.pad_tab)
        run(400)
        assert root.winfo_width() == wide, root.winfo_width()
        root.geometry(f"{start_w}x{start_h}")
        settle(width_is(start_w), what="the reset for the next check")
    check("a window already wide enough is left alone, on select and on leave", a_window_already_wide_enough_is_left_alone)

    def a_maximized_window_is_left_alone():
        start_w, start_h = W["start"]
        # the guard itself, on every platform: a window that says it is zoomed is not widened,
        # however narrow (here: the start-up width, with state() stood in)
        root.state = lambda *a, **k: "zoomed"
        try:
            app.notebook.select(app.toner_tab_frame)
            run(400)
            assert root.winfo_width() == start_w, f"a zoomed window was widened to {root.winfo_width()}"
            assert getattr(app, "_analyzer_widened", None) is None
            app.notebook.select(app.pad_tab)
            run(400)
            assert root.winfo_width() == start_w
        finally:
            del root.state
        # and a really maximized window, where the window manager has one
        try:
            root.state("zoomed")
        except tk.TclError as e:
            print(f"        (skipped: this window manager has no zoomed state: {e})")
            return
        run(400)
        if root.state() != "zoomed":
            print("        (skipped: the window manager did not maximize the window; no WM under xvfb)")
            root.state("normal")
            return
        w = root.winfo_width()
        app.notebook.select(app.toner_tab_frame)
        run(400)
        assert root.state() == "zoomed" and root.winfo_width() == w, (root.state(), root.winfo_width(), w)
        assert getattr(app, "_analyzer_widened", None) is None
        app.notebook.select(app.pad_tab)
        run(400)
        assert root.state() == "zoomed" and root.winfo_width() == w
        root.state("normal")
        settle(width_is(start_w), what="the window back from maximized")
    check("a maximized window is left alone, on select and on leave", a_maximized_window_is_left_alone)

    def never_wider_than_the_screen():
        start_w, start_h = W["start"]
        narrow = W["need"] - 100                              # a screen narrower than the strip needs
        root.winfo_screenwidth = lambda: narrow
        try:
            app.notebook.select(app.toner_tab_frame)
            settle(width_is(narrow), what="the widening capped at the screen")
            assert app._analyzer_widened == (start_w, narrow)
            # what happens then: the strip still overflows and pack drops its last column
            assert app._toner_ctrl_frame.winfo_reqwidth() > app._toner_main_frame.winfo_width()
            app.notebook.select(app.pad_tab)
            settle(width_is(start_w), what="the restore from the capped width")
        finally:
            del root.winfo_screenwidth
    check("the widening never goes past the screen (on a screen too narrow the strip still overflows)",
          never_wider_than_the_screen)
    app.notebook.select(app.toner_tab_frame)
    settle(width_is(W["widened"]), what="the tab back for the rest")

    # ---------------------------------------------------------------- the switch and the terms
    def the_switch_and_the_terms():
        asked = []

        def decline():
            asked.append("declined")
            return False

        app._toner_show_harmonics_terms = decline
        app._toner_harmonics_var.set("on")
        run(300)
        assert asked == ["declined"], asked
        assert app._toner_harmonics_var.get() == "off" and app._toner_harmonics_int.get() == 0
        assert app._toner_chart_frame.winfo_ismapped() and not app._toner_harmonics_frame.winfo_ismapped()
        assert not app.settings.get("toner_unlocked") and not settings_on_disk().get("toner_unlocked")

        app._toner_show_harmonics_terms = lambda: asked.append("accepted") or True
        app._toner_harmonics_var.set("on")
        run(300)
        assert asked == ["declined", "accepted"], asked
        assert app._toner_harmonics_frame.winfo_ismapped() and not app._toner_chart_frame.winfo_ismapped()
        assert app.settings["toner_unlocked"] is True and settings_on_disk()["toner_unlocked"] is True
        assert settings_on_disk().get("analyzer_tab_shown_v3") is True, "the save did not carry the migration's flag"
        cv = app._toner_spectrum_canvas
        texts = [cv.itemcget(i, "text") for i in cv.find_all() if cv.type(i) == "text"]
        assert "experimental" in texts, texts

        app._toner_harmonics_var.set("off")
        run(200)
        assert app._toner_chart_frame.winfo_ismapped() and not app._toner_harmonics_frame.winfo_ismapped()

        def never():
            raise AssertionError("the terms were asked again after unlocking")
        app._toner_show_harmonics_terms = never
        app._toner_harmonics_var.set("on")
        run(200)
        assert app._toner_harmonics_frame.winfo_ismapped()
        # the position is remembered in the settings (written with the next regular save)
        assert app.settings["toner_settings"]["harmonics_on"] is True
        app._toner_save_settings()
        assert app.settings["toner_settings"]["harmonics_on"] is True
        app._toner_harmonics_var.set("off")
        run(200)
        assert app.settings["toner_settings"]["harmonics_on"] is False
    check("HARMONICS: declined terms put the switch back and unlock nothing; accepted terms unlock (saved) and show "
          "the spectrum view, marked experimental; off again shows the chart; unlocked, the terms are not asked again; "
          "the position is remembered in the settings", the_switch_and_the_terms)

    # ---------------------------------------------------------------- required fields both ways
    def required_fields_both_ways():
        def form(**values):
            keys = ("name", "horn_type", "horn_make", "horn_model", "player", "mouthpiece", "mic_type", "mic_model")
            return {k: tk.StringVar(value=values.get(k, "")) for k in keys}
        full = dict(name="P", horn_type="Alto", horn_make="Conn", horn_model="6M", player="Me", mouthpiece="M5")
        app._toner_harmonics_var.set("off")
        assert app._toner_validate_required_fields(form(**full), root) is True, "no mic is fine for intonation"
        assert app._toner_validate_required_fields(form(**dict(full, player="")), root) is False
        assert app._toner_validate_required_fields(form(**dict(full, horn_type="")), root) is False
        app._toner_harmonics_var.set("on")
        run(150)
        assert app._toner_validate_required_fields(form(**full), root) is False, "harmonics on: the mic is required"
        assert app._toner_validate_required_fields(form(**dict(full, mic_type="condenser", mic_model="AT2020")),
                                                   root) is True
        app._toner_harmonics_var.set("off")
        run(150)
        # the same rule at Capture on a preset with no mic type
        app._toner_active_library, app._toner_active_preset = "My Presets", "Other"
        app._toner_active_session = None
        app._toner_update_preset_label()
        assert not app._toner_presets["My Presets"]["Other"].get("mic_type"), "the seed grew a mic type"
        app._toner_harmonics_var.set("on")
        run(150)
        app._toner_toggle_capture()
        assert app._toner_active_session is None and app._toner_capture_state is None, "harmonics on: a mic is needed"
        app._toner_harmonics_var.set("off")
        run(150)
        app._toner_toggle_capture()
        assert app._toner_active_session is not None and app._toner_capture_state == "free_listening"
        app._toner_stop_capture()
        app._toner_unload_preset()
    check("required fields both ways: horn type, make, model, player, mouthpiece for a run; the mic fields and the "
          "'Mic Type Required' stop at Capture only while HARMONICS is on", required_fields_both_ways)

    # ---------------------------------------------------------------- the runs through the synthetic source
    clock = FakeClock()
    toner_tab.time = clock                         # the session dates, under the test's control

    def load(name):
        if app._toner_capture_state is not None:
            app._toner_free_accumulator = []
            app._toner_stop_capture()
        clock.t += 60.0
        preset = app._toner_presets["My Presets"][name]
        app._toner_active_library, app._toner_active_preset = "My Presets", name
        app._toner_active_session = None
        app._toner_sax_var.set(preset["horn_type"])
        app._toner_update_preset_label()

    def free_run_on_the_synthetic_source():
        load("Bench")
        e = app._toner_engine
        e.synthetic_hz = 293.66         # D4 first: a note the seeded library does not have (written E5 on a tenor)
        app._toner_start()
        app._toner_mode_var.set("free")
        app._toner_toggle_capture()
        assert app._toner_capture_state == "free_listening"
        run(300)
        assert "E5" not in chart_texts(), "the seeded chart already shows the note about to be played"
        # D4 held: the synthetic source hands over 1024 samples a frame, at whatever rate the
        # loop manages here, so wait for the readout to turn green rather than a fixed time
        deadline = time.time() + 12.0
        while app._toner_capture_label.cget("fg") != toner_tab.CAPTURE_READY_FG and time.time() < deadline:
            run(300)
        assert app._toner_capture_label.cget("fg") == toner_tab.CAPTURE_READY_FG, \
            (app._toner_capture_label.cget("text"), app._toner_capture_label.cget("fg"))
        first = app._toner_engine.audio_time_s
        e.synthetic_hz = F0             # A#3: the note changes (silence would stop the synthetic stream)
        deadline = time.time() + 12.0
        while app._toner_engine.audio_time_s - first < 1.6 and time.time() < deadline:
            run(300)
        # the first note was saved when the second began, and the chart drew it while the run goes on
        # (nothing resized: the redraw comes from the capture itself)
        assert len(app._toner_active_session["captures"]) == 1, app._toner_active_session["captures"]
        assert "E5" in chart_texts(), ("the chart did not draw the captured note during the run", chart_texts())
        app._toner_toggle_capture()     # Stop: the last note is saved
        app._toner_stop()
        notes = [c["note"] for c in app._toner_active_session["captures"]]
        assert notes == ["D4", "A#3"], notes
        for c in app._toner_active_session["captures"]:
            assert c["method"] == "free" and c["n_frames"] >= FREE_MIN_FRAMES, c
        run(300)
        texts = chart_texts()
        # Tenor: concert A#3 is written C5, D4 is written E5; the seeded C4 is written D5
        for written in ("C5", "E5", "D5"):
            assert written in texts, (written, texts)
        assert "Bench" in texts, texts                     # the title
        assert len(chart_markers()) >= 3, chart_markers()
        disk = load_tone_presets(TONER_DATA_FILE)["My Presets"]["Bench"]
        assert len(disk["sessions"]) == 2 and [c["note"] for c in disk["sessions"][1]["captures"]] == ["D4", "A#3"]
    check("a free run on the synthetic source: two notes held, each kept, the readout green on a held note, the "
          "chart with the written notes and the preset's name, the session on disk", free_run_on_the_synthetic_source)

    def guided_run_on_the_synthetic_source():
        load("Other")                                   # Alto
        e = app._toner_engine
        old = toner_tab.CAL_COUNTDOWN_S, toner_tab.CALIBRATION_DURATION_S
        toner_tab.CAL_COUNTDOWN_S, toner_tab.CALIBRATION_DURATION_S = 0.4, 1.0    # the shape of the run, shorter
        try:
            # the first guided note is written A#3; on an alto that sounds C#3
            e.synthetic_hz = note_to_freq("C#3")
            app._toner_start()
            app._toner_mode_var.set("calibration")
            app._toner_toggle_capture()
            assert app._toner_capture_state == "cal_countdown"
            # the loop's frame rate is whatever this machine manages: wait for each state, bounded
            settle(lambda: app._toner_capture_state == "calibration", timeout=15.0, what="the countdown")
            prompt = [app._toner_chart_canvas.itemcget(i, "text")
                      for i in app._toner_chart_canvas.find_withtag("calprompt")]
            # the large note name, as a player writes it, whether still waiting or already recording
            assert prompt and prompt[0] == "Bb3", prompt
            settle(lambda: app._toner_capture_state == "cal_pause", timeout=20.0, what="the recording")
            prompt = [app._toner_chart_canvas.itemcget(i, "text")
                      for i in app._toner_chart_canvas.find_withtag("calprompt")]
            assert "A#3" not in prompt, prompt
            assert "Bb3" in app._toner_capture_label.cget("text"), app._toner_capture_label.cget("text")
            caps = app._toner_active_session["captures"]
            assert len(caps) == 1 and caps[0]["method"] == "calibration", caps
            assert caps[0]["written_note"] == "A#3" and caps[0]["note"] == "C#3", caps[0]
            assert caps[0]["n_frames"] >= FREE_MIN_FRAMES
            app._toner_toggle_capture()
            app._toner_stop()
            run(300)
            texts = chart_texts()
            assert "Bb3" in texts and "A#3" not in texts and "Other" in texts, texts
            assert len(chart_markers()) >= 1
        finally:
            toner_tab.CAL_COUNTDOWN_S, toner_tab.CALIBRATION_DURATION_S = old
    check("a guided run on the synthetic source: countdown, the note asked for (shown large on the chart canvas), "
          "recorded and filed under the asked note, the chart with it", guided_run_on_the_synthetic_source)

    # ---------------------------------------------------------------- the readout, frame by frame
    STEP_SAMPLES = 1455
    STEP = STEP_SAMPLES / SAMPLE_RATE

    def tick():
        clock.t += STEP
        app._toner_engine._samples_in += STEP_SAMPLES

    def feed(n, f0=F0, note="A#3"):
        for _ in range(n):
            tick()
            app._toner_process_capture_frame(frame(f0, note))

    def kept_of(n):
        times = [k * STEP for k in range(1, n + 1)]
        return sum(whole_window(times, times[0], times[-1]))

    SHORTEST = next(n for n in range(FREE_MIN_FRAMES, 200) if kept_of(n) >= FREE_MIN_FRAMES)

    def the_readout_turns_green_at_the_right_frame():
        assert SHORTEST == 34, SHORTEST
        load("Bench")
        app._toner_mode_var.set("free")
        app._toner_toggle_capture()
        label = app._toner_capture_label
        feed(SHORTEST - 1)
        assert label.cget("fg") == "#FFCC00" and label.cget("text") == f"Free: C5 ({kept_of(SHORTEST - 1)} frames)", \
            (label.cget("text"), label.cget("fg"))
        assert kept_of(SHORTEST - 1) < FREE_MIN_FRAMES
        feed(1)
        assert label.cget("fg") == "#00CC00" and label.cget("text") == f"Free: C5 ({FREE_MIN_FRAMES} frames)", \
            (label.cget("text"), label.cget("fg"))
        app._toner_process_capture_frame(silence())
        assert label.cget("fg") == "#FFCC00"
        assert len(app._toner_active_session["captures"]) == 1
        assert "too short" not in app._toner_capture_progress.cget("text")
        # a note let go before green
        feed(20, f0=293.66, note="D4")
        assert label.cget("fg") == "#FFCC00"
        app._toner_process_capture_frame(silence())
        assert len(app._toner_active_session["captures"]) == 1
        assert app._toner_capture_progress.cget("text").endswith("E5 too short to keep"), \
            app._toner_capture_progress.cget("text")
        feed(SHORTEST)
        app._toner_process_capture_frame(silence())
        assert "too short" not in app._toner_capture_progress.cget("text"), "a kept note clears the words"
        app._toner_stop_capture()
    check("the readout counts the frames a capture would keep and turns green at the 34th frame; a note let go "
          "sooner is named too short until the next note is kept", the_readout_turns_green_at_the_right_frame)

    # ---------------------------------------------------------------- Compare and Export
    def compare_two_presets():
        load("Bench")
        run(200)
        before = set(toplevels())
        app._toner_open_compare_dialog()
        dlg = [w for w in toplevels() if w not in before][0]
        boxes = {c.cget("text"): c for c in widgets(dlg, tk.Checkbutton)}
        # the two presets and the sessions of each (the runs above added sessions of their own)
        assert {"[My Presets] Bench", "[My Presets] Other", "2026-10-01 10:00:00", "2026-10-02 10:00:00"} <= set(boxes), \
            sorted(boxes)
        boxes["[My Presets] Bench"].invoke()
        boxes["2026-10-02 10:00:00"].invoke()
        buttons = {b.cget("text"): b for b in widgets(dlg, tk.Button)}
        buttons["Compare"].invoke()
        assert not dlg.winfo_exists()
        assert app._toner_compare_runs == [("My Presets", "Bench", None), ("My Presets", "Other", "2026-10-02 10:00:00")]
        run(300)
        texts = chart_texts()
        assert "Bench" in texts and "Other — 2026-10-02 10:00:00" in texts, texts
        kinds = chart_markers()
        assert "oval" in kinds and "polygon" in kinds, kinds        # a circle for the first run, squares for the second
        # Other's second session, an alto: D4 and E4 sound, written B4 and C#5
        for written in ("B4", "C#5", "C5", "D5", "E5"):
            assert written in texts, (written, texts)
        # Clear goes back to the loaded preset alone
        before = set(toplevels())
        app._toner_open_compare_dialog()
        dlg = [w for w in toplevels() if w not in before][0]
        {b.cget("text"): b for b in widgets(dlg, tk.Button)}["Clear"].invoke()
        assert app._toner_compare_runs == []
        run(300)
        assert "B4" not in chart_texts() and "Bench" in chart_texts()
    check("Compare...: a preset and one session of another drawn together, each with its name and marker; Clear "
          "returns to the loaded preset", compare_two_presets)

    def takes_that_are_not_their_note_are_left_out_and_said():
        load("Old")
        run(300)
        texts = chart_texts()
        assert "4 takes left out: their pitch did not match their note." in texts, texts
        assert "No take matched its note." in texts, texts
        # no point in the plot (the legend's own marker, up by the title, is not a point)
        cv = app._toner_chart_canvas
        in_plot = [i for i in cv.find_withtag("chart") if cv.type(i) in ("oval", "polygon") and cv.coords(i)[1] > 60]
        assert not in_plot, "a point drawn for a take that is not its note"
        # none of the bogus labels, as written names, reach the axis
        assert not any(t in ("C2", "D2", "E2", "G2") for t in texts), texts
        # the export carries the same line
        runs, prims = app._toner_chart_prims(toner_tab.CHART_W, toner_tab.CHART_H)
        assert runs[0].left_out["off_note"] == 4 and runs[0].readings == []
        assert any(p["kind"] == "text" and p["text"].startswith("4 takes left out") for p in prims)
        # a good preset says nothing of the kind
        load("Bench")
        run(300)
        assert not any("left out" in t for t in chart_texts()), chart_texts()
    check("a preset whose takes are not their notes (labels two transpositions low) draws no point; the chart says "
          "how many were left out and why, on screen and in the export; a good preset says nothing",
          takes_that_are_not_their_note_are_left_out_and_said)

    def export_png_and_pdf():
        import tkinter.filedialog as fd
        out = tempfile.mkdtemp(prefix="ssc-analyzer-export-")
        asked = []

        def save_as(**kw):
            asked.append(kw)
            return asked[-1]["_path"] if "_path" in kw else target[0]

        target = [os.path.join(out, "bench.png")]
        fd.asksaveasfilename = save_as
        load("Bench")
        app._toner_export_chart()
        assert asked and asked[-1]["initialfile"] == "Bench.png", asked
        with open(target[0], "rb") as f:
            head = f.read(24)
        assert head[:8] == b"\x89PNG\r\n\x1a\n", head
        w, h = struct.unpack(">II", head[16:24])
        assert (w, h) == (2 * toner_tab.CHART_W, 2 * toner_tab.CHART_H), (w, h)
        if intonation_chart.PNG_AVAILABLE:
            from PIL import Image
            img = Image.open(target[0]).convert("RGB")
            colours = {img.getpixel((x, y)) for x in range(0, img.width, 25) for y in range(0, img.height, 25)}
            assert len(colours) > 1, "a blank image"
        target[0] = os.path.join(out, "bench.pdf")
        app._toner_export_chart()
        with open(target[0], "rb") as f:
            pdf = f.read()
        assert pdf.startswith(b"%PDF-1.4") and pdf.rstrip().endswith(b"%%EOF"), pdf[:20]
        assert pdf.count(b"/Type /Page\n") == 1 or pdf.count(b"/Type /Page ") == 1 or b"/Count 1" in pdf, "not one page"
        assert b"(Bench)" in pdf, "the title is not in the PDF"
        # nothing to export: no file dialog
        app._toner_unload_preset()
        asked.clear()
        app._toner_export_chart()
        assert asked == [], "a save dialog with nothing to save"
    check("Export...: a PNG at twice the chart's size with something drawn, and a one-page PDF carrying the title; "
          "nothing to export asks for no file", export_png_and_pdf)

    # ---------------------------------------------------------------- the Settings tick
    def low_notes_left_setting():
        load("Bench")
        run(300)

        def x_of(text):
            cv = app._toner_chart_canvas
            xs = [cv.coords(i)[0] for i in cv.find_withtag("chart") if cv.type(i) == "text" and cv.itemcget(i, "text") == text]
            assert xs, text
            return xs[0]
        assert x_of("C5") < x_of("D5"), "low notes on the left by default"     # the seeded A#3 and C4, written
        before = set(toplevels())
        app._toner_open_settings()
        dlg = [w for w in toplevels() if w not in before][0]
        boxes = {c.cget("text"): c for c in widgets(dlg, tk.Checkbutton)}
        assert "Low notes on the left" in boxes, sorted(boxes)
        boxes["Low notes on the left"].invoke()
        {b.cget("text"): b for b in widgets(dlg, tk.Button)}["OK"].invoke()
        assert not dlg.winfo_exists()
        assert app.settings["toner_settings"]["chart_low_notes_left"] is False
        assert settings_on_disk()["toner_settings"]["chart_low_notes_left"] is False
        run(300)
        assert x_of("C5") > x_of("D5"), "the axis did not flip"
        app.settings["toner_settings"]["chart_low_notes_left"] = True
        app._toner_chart_changed()
        run(300)
        assert x_of("C5") < x_of("D5")
    check("Settings > General: 'Low notes on the left' is saved and flips the axis", low_notes_left_setting)

    toner_tab.time = time
    app._toner_unload_preset()

    # ---------------------------------------------------------------- the tour's stop
    def the_tours_stop():
        import main as main_mod
        log, errors = main_mod.run_tour(root, app, shots_dir=None, step_ms=150, on_done=root.quit)
        root.after(90000, root.quit)
        root.mainloop()
        assert "toner" in log, log
        assert not errors, "\n".join(errors)
        # unlocked by the terms check above: the tour showed the spectrum on the tone
        assert app._toner_harmonics_var.get() == "on" and app._toner_harmonics_frame.winfo_ismapped()
        app._toner_stop()
    check("the tour's stop for the tab raises nothing and, unlocked, shows the spectrum", the_tours_stop)

    root.destroy()

    # ---------------------------------------------------------------- in German, in a child
    CHILD = r'''
import json, os, sys
sys.path.insert(0, sys.argv[1])
from i18n import init_translation
init_translation(sys.argv[2])
import tkinter as tk
import time
import tkinter as tk
from tkinter import font as tkfont, ttk
root = tk.Tk()
root.geometry("+0+0")
import main as main_mod
from main import PadSVGGeneratorApp
app = PadSVGGeneratorApp(root)
app._toner_mic_checked = True

def run(ms):
    root.after(ms, root.quit); root.mainloop()

def settle(pred, timeout=6.0):
    deadline = time.time() + timeout
    while not pred() and time.time() < deadline:
        run(100)
    return pred()

root.update()
run(400)                              # the window-fit rule runs after_idle
# the old rule's own measure: the tab labels and the Tooling tab, never the strip
style = ttk.Style(root)
spec = style.lookup("TNotebook.Tab", "font") or "TkDefaultFont"
font = tkfont.nametofont(spec) if spec in tkfont.names(root) else tkfont.Font(root, font=spec)
visible = [t for t in app.notebook.tabs() if app.notebook.tab(t, "state") != "hidden"]
formula = sum(font.measure(app.notebook.tab(t, "text")) for t in visible) + main_mod.TAB_LABEL_PADDING_PX * len(visible) + 40
formula = max(formula, app.tooling_tab_frame.winfo_reqwidth() + app._notebook_page_chrome_px())
start_w, start_h = root.winfo_width(), root.winfo_height()
need = app._analyzer_strip_need_px()
screen_w = root.winfo_screenwidth()
app.notebook.select(app.toner_tab_frame)
want = max(start_w, min(need, screen_w))
settled = settle(lambda: root.winfo_width() == want)
after_select = (root.winfo_width(), root.winfo_height())
harm_mapped = app._toner_harmonics_frame.winfo_ismapped()
chart_mapped = app._toner_chart_frame.winfo_ismapped()
capture_mapped = app._toner_capture_btn.winfo_ismapped()
page_w = app._toner_main_frame.winfo_width()
labels = []
def walk(w):
    for c in w.winfo_children():
        if isinstance(c, tk.Label):
            labels.append(c.cget("text"))
        walk(c)
walk(app._toner_ctrl_frame)
buttons = [b.cget("text") for b in app._toner_chart_frame.winfo_children()[1].winfo_children()]
app.notebook.select(app.pad_tab)
settled_back = settle(lambda: root.winfo_width() == start_w)
after_leave = (root.winfo_width(), root.winfo_height())
out = {"tab": app.notebook.tab(app.toner_tab_frame, "text"), "labels": labels, "buttons": buttons,
       "formula_w": formula, "start": [start_w, start_h], "need": need, "screen_w": screen_w,
       "settled": settled, "after_select": list(after_select), "page_w": page_w,
       "harm_mapped": harm_mapped, "chart_mapped": chart_mapped, "capture_mapped": capture_mapped,
       "settled_back": settled_back, "after_leave": list(after_leave),
       "harmonics_var": app._toner_harmonics_var.get()}
app._toner_stop()
root.destroy()
print("REPORT " + json.dumps(out))
'''

    CASES = [
        ("en", {"language": "en"}), ("es", {"language": "es"}), ("de", {"language": "de"}),
        ("fr", {"language": "fr"}), ("it", {"language": "it"}),
        # the switch remembered on, with the terms accepted; and "on" in a profile that never
        # unlocked (a hand edit), which counts as off
        ("en-remembered-on", {"language": "en", "toner_unlocked": True, "toner_settings": {"harmonics_on": True}}),
        ("en-hand-edited-on", {"language": "en", "toner_settings": {"harmonics_on": True}}),
    ]

    def in_every_language_and_the_remembered_switch():
        reports = {}
        for case, profile in CASES:
            folder = tempfile.mkdtemp(prefix=f"ssc-analyzer-{case}-")
            with open(os.path.join(folder, "app_settings.json"), "w", encoding="utf-8") as f:
                json.dump(dict(profile, toner_record_wav=False), f)
            env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
            p = subprocess.run([sys.executable, "-c", CHILD, REPO, profile["language"]], capture_output=True,
                               text=True, env=env, timeout=240, cwd=REPO, encoding="utf-8")
            line = [ln for ln in (p.stdout or "").splitlines() if ln.startswith("REPORT ")]
            assert p.returncode == 0 and line, (case, p.returncode, (p.stdout or "")[-300:], (p.stderr or "")[-1200:])
            reports[case] = json.loads(line[0][7:])
        for case, r in reports.items():
            start_w, start_h = r["start"]
            # the start-up width is the old rule's: at least its measure, and never the strip's
            assert start_w >= r["formula_w"], f"{case}: opened {start_w} px, the tab/Tooling rule asks {r['formula_w']}"
            assert start_w < r["need"], f"{case}: the start-up window ({start_w} px) is already the strip's ({r['need']})"
            want = max(start_w, min(r["need"], r["screen_w"]))
            assert r["settled"] and r["after_select"] == [want, start_h], (case, r["after_select"], want, start_h)
            assert r["capture_mapped"] or r["need"] > r["screen_w"], f"{case}: Capture off the page at {r['after_select']}"
            assert r["settled_back"] and r["after_leave"] == [start_w, start_h], (case, r["after_leave"], r["start"])
            assert len(r["buttons"]) == 2, (case, r["buttons"])
        plain = {c: r for c, r in reports.items() if "-" not in c}
        for case, r in plain.items():
            assert r["chart_mapped"] and not r["harm_mapped"] and r["harmonics_var"] == "off", case
        en, de = reports["en"], reports["de"]
        assert "HARMONICS" in en["labels"] and "guided" in en["labels"], en["labels"]
        assert len(de["labels"]) == len(en["labels"]), (de["labels"], en["labels"])
        changed = [a for a, b in zip(en["labels"], de["labels"]) if a != b and re.search(r"[A-Za-z]{3,}", a)]
        assert changed, "no control-strip label is translated in German"
        on = reports["en-remembered-on"]
        assert on["harm_mapped"] and not on["chart_mapped"] and on["harmonics_var"] == "on", on
        hand = reports["en-hand-edited-on"]
        assert hand["chart_mapped"] and not hand["harm_mapped"] and hand["harmonics_var"] == "off", hand
    check("in every language (a child process each): the start-up width is the old rule's, selecting the tab widens "
          "to the strip and leaving restores, German labels translated; the switch remembered on shows the spectrum "
          "on an unlocked profile and counts as off on one that never unlocked",
          in_every_language_and_the_remembered_switch)

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
