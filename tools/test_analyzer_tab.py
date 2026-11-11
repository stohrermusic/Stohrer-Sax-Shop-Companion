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
  - in German, in a child process: the tab builds and its control strip fits
Frames run under a real mainloop with a quit timer (see test_tuner_canvas for why a
root.update() loop can't be used). Self-skips without a display or the audio stack.
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
        app.notebook.select(app.toner_tab_frame)
        run(400)
        assert app._toner_chart_frame.winfo_ismapped() and not app._toner_harmonics_frame.winfo_ismapped()
        assert app._toner_chart_canvas.winfo_width() > 300 and app._toner_chart_canvas.cget("bg") == "#FFFFFF"
        # the window opened wide enough for the whole control strip: the Capture button, packed
        # last and rightmost, is on the page (at the 640 px start it was not)
        strip, btn = app._toner_ctrl_frame, app._toner_capture_btn
        assert strip.winfo_reqwidth() <= app._toner_main_frame.winfo_width(), \
            (strip.winfo_reqwidth(), app._toner_main_frame.winfo_width(), root.winfo_width())
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
        app._toner_harmonics_var.set("off")
        run(200)
    check("HARMONICS: declined terms put the switch back and unlock nothing; accepted terms unlock (saved) and show "
          "the spectrum view, marked experimental; off again shows the chart; unlocked, the terms are not asked again",
          the_switch_and_the_terms)

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
            run(900)
            assert app._toner_capture_state in ("calibration",), app._toner_capture_state
            prompt = [app._toner_chart_canvas.itemcget(i, "text")
                      for i in app._toner_chart_canvas.find_withtag("calprompt")]
            assert prompt and prompt[0] == "A#3", prompt      # the large note name, on the chart canvas too
            run(2200)
            assert app._toner_capture_state == "cal_pause", app._toner_capture_state
            caps = app._toner_active_session["captures"]
            assert len(caps) == 1 and caps[0]["method"] == "calibration", caps
            assert caps[0]["written_note"] == "A#3" and caps[0]["note"] == "C#3", caps[0]
            assert caps[0]["n_frames"] >= FREE_MIN_FRAMES
            app._toner_toggle_capture()
            app._toner_stop()
            run(300)
            texts = chart_texts()
            assert "A#3" in texts and "Other" in texts, texts
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
root = tk.Tk()
root.geometry("+0+0")
from main import PadSVGGeneratorApp
app = PadSVGGeneratorApp(root)
app._toner_mic_checked = True
root.update()
root.after(300, root.quit)
root.mainloop()                       # the window-fit rule runs after_idle
app.notebook.select(app.toner_tab_frame)
root.update()
labels = []
def walk(w):
    for c in w.winfo_children():
        if isinstance(c, tk.Label):
            labels.append(c.cget("text"))
        walk(c)
walk(app._toner_ctrl_frame)
out = {"tab": app.notebook.tab(app.toner_tab_frame, "text"), "labels": labels,
       "ctrl_req": app._toner_ctrl_frame.winfo_reqwidth(), "page_w": app._toner_main_frame.winfo_width(),
       "window_w": root.winfo_width(), "capture_mapped": app._toner_capture_btn.winfo_ismapped(),
       "chart": app._toner_chart_frame.winfo_ismapped(), "buttons": [b.cget("text") for b in app._toner_chart_frame.winfo_children()[1].winfo_children()]}
app._toner_stop()
root.destroy()
print("REPORT " + json.dumps(out))
'''

    def in_german():
        reports = {}
        for lang in ("en", "de"):
            folder = tempfile.mkdtemp(prefix=f"ssc-analyzer-{lang}-")
            with open(os.path.join(folder, "app_settings.json"), "w", encoding="utf-8") as f:
                json.dump({"language": lang, "toner_record_wav": False}, f)
            env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
            p = subprocess.run([sys.executable, "-c", CHILD, REPO, lang], capture_output=True, text=True, env=env,
                               timeout=240, cwd=REPO, encoding="utf-8")
            line = [ln for ln in (p.stdout or "").splitlines() if ln.startswith("REPORT ")]
            assert p.returncode == 0 and line, (lang, p.returncode, (p.stdout or "")[-300:], (p.stderr or "")[-1200:])
            reports[lang] = json.loads(line[0][7:])
        en, de = reports["en"], reports["de"]
        for lang, r in reports.items():
            assert r["chart"], f"{lang}: the chart is not the view shown"
            assert len(r["buttons"]) == 2, (lang, r["buttons"])
            assert r["ctrl_req"] <= r["page_w"] and r["capture_mapped"], \
                f"{lang}: the control strip asks {r['ctrl_req']} px of a {r['page_w']} px page ({r['window_w']} px window)"
        assert "HARMONICS" in en["labels"] and "guided" in en["labels"], en["labels"]
        assert len(de["labels"]) == len(en["labels"]), (de["labels"], en["labels"])
        # the German catalog carries the new strings once it is compiled; the switch labels are
        # translated (a word the two languages share may stay)
        changed = [a for a, b in zip(en["labels"], de["labels"]) if a != b and re.search(r"[A-Za-z]{3,}", a)]
        assert changed, "no control-strip label is translated in German"
    check("in German (a child process): the tab builds on the chart, its control strip fits the page, and its "
          "labels are translated", in_german)

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
