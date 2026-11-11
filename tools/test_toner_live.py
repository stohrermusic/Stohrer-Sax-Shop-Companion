"""Drive the Toner tab end to end on a synthetic tone — no microphone needed.

`TonerEngine.synthetic_hz` replaces the audio stream with a generated Bb3
(233.08 Hz) carrying six harmonics, so the whole tab runs on a machine
with no input device: the real `_toner_animate` loop, the spectrum bars,
the note and frequency readouts, the intonation gauge. This is the toner's
only end-to-end check on the macOS CI runner.

The harmonic analyzer is behind the terms (the tab itself shows by default since
2026-11-10, opening on the intonation chart), so the test hands the app a prepared
profile through SAXSHOP_CONFIG_DIR (harmonics unlocked, tab visible) — which also
keeps it away from the user's real settings — and turns the HARMONICS switch on. The first-start
microphone notice is a modal; `_toner_mic_checked` is set so it never opens.

Frames run under a real mainloop with a quit timer (see test_tuner_canvas
for why a root.update() loop can't be used). Self-skips without a display.

The second half (2026-11-10) pins the CAPTURE path by feeding the tab's state
machine fabricated analysis frames on the engine's audio clock: the tab built
on a seeded toner_data.json (the mic backfill, the preset label), free capture
(only the frames whose whole analysis window lies inside the note make the
micro-capture, how short a note is dropped, the keys written, the deferred
save), calibration capture (the countdown and the 5 s recording on audio time,
the frames from WINDOW_S in, the release dropped when the note stops early, the
capture stored under the note ASKED for with `written_note` and `detected_as`,
the wait for silence), the session saved to disk once per date and read back,
and the exit flush. The capture rules date from 2026-11-10; the characterization
of the rules before them is in the report of that day.
"""
import copy
import json
import math
import os
import re
import sys
import tempfile
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-toner-test-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
with open(os.path.join(_PROFILE, "app_settings.json"), "w", encoding="utf-8") as _f:
    json.dump({"toner_unlocked": True, "visible_tabs": {"Toner": True},
               "toner_record_wav": False, "seen_calibration_tutorial": True}, _f)
from i18n import init_translation  # noqa: E402
init_translation("en")
import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

# A tone library as the app writes it: a complete preset with one earlier session,
# and one without a mic type whose sessions carry it (the backfill's case).
F0 = 233.08
HARM = [0.0, -3.0, -6.0, -10.0, -14.0, -18.0, -22.0]


def _capture(note, f0, method="free", **extra):
    return {"note": note, "fundamental_freq": f0, "harmonics_db": list(HARM),
            "harmonic_cents": [0.0] * len(HARM), "timestamp": "2026-11-01 10:00:00", "n_frames": 30,
            "method": method, "signal_level": 0.5, "spectral_centroid": 900.0, **extra}


SEEDED = {
    "My Presets": {
        "Bench": {"horn_type": "Tenor", "horn_make": "Selmer", "horn_model": "Mark VI", "serial": "", "player": "Matt",
                  "mouthpiece": "Link 7*", "reed": "", "ligature": "", "mic_position": "", "room": "", "preamp": "",
                  "mic_type": "condenser", "mic_model": "AT2020", "notes": "", "created": "2026-11-01",
                  "sessions": [{"date": "2026-11-01 10:00:00", "mic_type": "condenser", "mic_model": "AT2020",
                                "horn_type": "Tenor", "captures": [_capture("A#3", F0), _capture("C4", 261.63)]}]},
        "NoMic": {"horn_type": "Alto", "horn_make": "Conn", "horn_model": "6M", "player": "Matt", "mouthpiece": "Meyer 5",
                  "notes": "", "created": "2026-10-01",
                  "sessions": [{"date": "2026-10-01 10:00:00", "mic_type": "ribbon", "mic_model": "R121",
                                "captures": [_capture("A#3", F0)]},
                               {"date": "2026-10-02 10:00:00", "mic_type": "ribbon", "mic_model": "",
                                "captures": []}]},
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
    """One analysis frame as the engine reports it."""
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


def main():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    from toner_engine import AUDIO_AVAILABLE
    if not AUDIO_AVAILABLE:
        print("Skipping: audio libraries not installed (numpy/sounddevice) — no toner tab on this build")
        root.destroy()
        return 0

    print("Toner tab on a synthetic tone")
    print("=" * 60)
    import config
    check("profile override is in effect",
          lambda: os.path.normcase(config.get_config_dir()) == os.path.normcase(_PROFILE))

    from main import PadSVGGeneratorApp
    root.geometry("1100x760+0+0")
    app = PadSVGGeneratorApp(root)
    root.update()
    check("toner unlocked and tab visible from the profile",
          lambda: app.settings.get("toner_unlocked") is True
          and str(app.toner_tab_frame) in app.notebook.tabs())
    assert getattr(app, "_toner_engine", None) is not None, "toner tab not built (audio libs missing?)"

    app._toner_mic_checked = True          # the first-start mic notice is a modal
    app._toner_engine.synthetic_hz = F0
    app._toner_harmonics_var.set("on")      # the spectrum view; unlocked in the profile, so no terms
    app.notebook.select(app.toner_tab_frame)   # on_tab_changed -> _toner_start
    root.after(2000, root.quit)
    root.mainloop()

    check("the HARMONICS switch puts the spectrum in the top area (no terms: the profile is unlocked)",
          lambda: app._toner_harmonics_frame.winfo_ismapped() and not app._toner_chart_frame.winfo_ismapped())

    check("toner is running on the synthetic source (no stream opened)",
          lambda: app._toner_engine.is_running and app._toner_engine._stream is None)
    check("spectrum bars built", lambda: app._toner_bars_built)

    def no_error_text():
        c = getattr(app, "_toner_spectrum_canvas", None)
        if c is not None:
            for item in c.find_all():
                if c.type(item) == "text" and "error" in (c.itemcget(item, "text") or "").lower():
                    raise AssertionError(c.itemcget(item, "text"))
    check("no error text on the spectrum canvas", no_error_text)

    def readouts():
        note = app._toner_note_label.cget("text")
        freq_txt = app._toner_freq_label.cget("text")
        assert note and note != "—", f"note readout {note!r}"
        m = re.search(r"([0-9]+(?:\.[0-9]+)?)", freq_txt)
        assert m, f"frequency readout {freq_txt!r}"
        shown = float(m.group(1))
        assert abs(shown - F0) < 0.15, f"frequency readout {freq_txt!r} vs {F0}"
    check("note and frequency readouts show the tone", readouts)

    def engine_live_result():
        r = app._toner_engine.analyze()
        assert r.fundamental_freq > 0, "no fundamental"
        cents = 1200.0 * math.log2(r.fundamental_freq / F0)
        assert abs(cents) < 0.5, f"fundamental {r.fundamental_freq:.3f} Hz = {cents:+.2f} c"
        assert len(r.harmonics) >= 5, f"only {len(r.harmonics)} harmonics"
        h2 = next((h for h in r.harmonics if h.harmonic_number == 2), None)
        assert h2 is not None and abs(h2.magnitude_db + 3.0) < 1.0, f"H2 {h2 and h2.magnitude_db}"
    check("engine reads Bb3 within 0.5 c with its harmonics through the live path", engine_live_result)

    app._toner_stop()
    root.update_idletasks()
    check("stops cleanly", lambda: not app._toner_engine.is_running)

    # ------------------------------------------------------------------
    # The capture path as it is (2026-11-10): the tab's state machine fed
    # fabricated frames, on the seeded library.
    # ------------------------------------------------------------------
    print("-" * 60)
    print("Capture path on the seeded library")
    import toner_tab
    from toner_engine import FREE_MIN_FRAMES, SAMPLE_RATE, WINDOW_S, load_tone_presets, whole_window
    from config import TONER_DATA_FILE
    clock = FakeClock()
    toner_tab.time = clock                         # the tab's clock, under the test's control
    app._toner_show_coverage_summary = lambda: None    # a Toplevel with buttons; not under test here

    def on_disk():
        return load_tone_presets(TONER_DATA_FILE)

    def seeded_library_loaded_and_mic_backfilled():
        lib = app._toner_presets["My Presets"]
        assert set(lib) == {"Bench", "NoMic"}, "set(lib) == {\"Bench\", \"NoMic\"}"
        # the preset without a mic type took its sessions' most common mic type and first model, and was saved
        assert lib["NoMic"]["mic_type"] == "ribbon" and lib["NoMic"]["mic_model"] == "R121", lib["NoMic"]
        disk = on_disk()["My Presets"]["NoMic"]
        assert disk["mic_type"] == "ribbon" and disk["mic_model"] == "R121", "the backfill was not saved"
        assert lib["Bench"]["mic_type"] == "condenser" and len(lib["Bench"]["sessions"]) == 1, "lib[\"Bench\"][\"mic_type\"] == \"condenser\" and len(lib[\"Bench\"][\"sessions\"]) == 1"
    check("the tab loads the seeded library and backfills a missing mic type from the sessions, saving it",
          seeded_library_loaded_and_mic_backfilled)

    def load_bench():
        app._toner_active_library, app._toner_active_preset = "My Presets", "Bench"
        app._toner_active_session = None
        app._toner_sax_var.set("Tenor")
        app._toner_update_preset_label()

    def loading_a_preset_shows_it():
        load_bench()
        assert app._toner_preset_label.cget("text") == "Bench", "app._toner_preset_label.cget(\"text\") == \"Bench\""
        assert app._toner_mic_label.cget("text") == "AT2020", "app._toner_mic_label.cget(\"text\") == \"AT2020\""
        app._toner_active_preset, app._toner_active_library = None, None
        app._toner_update_preset_label()
        assert app._toner_preset_label.cget("text") == "no preset" and app._toner_mic_label.cget("text") == "no preset", "app._toner_preset_label.cget(\"text\") == \"no preset\" and app._toner_mic_label.cget(\"text\") == \"no preset\""
    check("loading a preset shows its name and its mic model; unloading shows 'no preset'", loading_a_preset_shows_it)

    def start_capture(mode):
        if app._toner_capture_state is not None:   # a check that failed mid-capture must not spoil the next
            app._toner_free_accumulator = []
            app._toner_stop_capture()
        clock.t += 60.0                        # a new minute: sessions are keyed by their date to the second
        load_bench()
        app._toner_mode_var.set(mode)
        app._toner_toggle_capture()            # a new session, then listening
        return app._toner_active_session

    STEP_SAMPLES = 1455                            # one frame at 30 fps: 33 ms of sound handed over
    STEP = STEP_SAMPLES / SAMPLE_RATE              # ... on the audio clock, exactly

    def tick(seconds=STEP):
        clock.t += seconds
        app._toner_engine._samples_in += STEP_SAMPLES if seconds == STEP else int(round(seconds * SAMPLE_RATE))

    def feed(n, f0=F0, note="A#3", **kw):
        for _ in range(n):
            tick()
            app._toner_process_capture_frame(frame(f0, note, **kw))

    def kept_of(n):
        """How many of n frames fed STEP apart have their whole window inside the note."""
        times = [k * STEP for k in range(1, n + 1)]
        return sum(whole_window(times, times[0], times[-1]))

    SHORTEST = next(n for n in range(FREE_MIN_FRAMES, 200) if kept_of(n) >= FREE_MIN_FRAMES)   # 34 frames, 1.1 s

    def the_audio_clock_runs_on_the_sound_handed_over():
        e = app._toner_engine
        e.synthetic_hz = F0
        e.start()
        assert e.audio_time_s == 0.0, "start() resets the clock"
        for _ in range(10):
            e.analyze()
        assert abs(e.audio_time_s - 10 * 1024 / SAMPLE_RATE) < 1e-9, e.audio_time_s
        e.stop()
        assert abs(WINDOW_S - 0.3715) < 0.001 and kept_of(40) == 16 and SHORTEST == 34, (kept_of(40), SHORTEST)
    check("the capture clock is audio time: the synthetic source hands over 1024 samples a frame; a 0.37 s window; "
          "16 of 40 frames are whole-window, 34 frames is the shortest note kept",
          the_audio_clock_runs_on_the_sound_handed_over)

    def free_capture_rules():
        session = start_capture("free")
        assert app._toner_capture_state == "free_listening" and session["captures"] == [], "app._toner_capture_state == \"free_listening\" and session[\"captures\"] == []"
        assert session["mic_type"] == "condenser" and session["horn_type"] == "Tenor" and session["player"] == "Matt", "session[\"mic_type\"] == \"condenser\" and session[\"horn_type\"] == \"Tenor\" and session[\"player\"] == \"Matt\""
        assert app._toner_capture_btn.cget("text") == "Stop", "app._toner_capture_btn.cget(\"text\") == \"Stop\""
        # a note held 40 frames, then silence: one capture of its whole-window frames (16: the 12 at
        # either end hold the onset or the release)
        feed(40)
        assert len(app._toner_free_accumulator) == 40, "every frame is gathered; the ends go at save time"
        app._toner_process_capture_frame(silence())
        caps = session["captures"]
        assert len(caps) == 1, caps
        c = caps[0]
        assert set(c) == {"note", "fundamental_freq", "harmonics_db", "harmonic_cents", "timestamp", "n_frames",
                          "method", "signal_level", "level_dbfs", "spectral_centroid"}, sorted(c)
        assert c["note"] == "A#3" and c["method"] == "free" and c["n_frames"] == kept_of(40) == 16, "c[\"note\"] == \"A#3\" and c[\"method\"] == \"free\" and c[\"n_frames\"] == kept_of(40) == 16"
        assert c["fundamental_freq"] == round(F0, 2) and c["harmonics_db"] == HARM and c["signal_level"] == 0.8, "c[\"fundamental_freq\"] == round(F0, 2) and c[\"harmonics_db\"] == HARM and c[\"signal_level\"] == 0.8"
        assert app._toner_captured_notes == {"A#3"} and app._toner_save_pending is True, "the free save is deferred"
        assert app._toner_capture_progress.cget("text") == "(1 notes captured)", "app._toner_capture_progress.cget(\"text\") == \"(1 notes captured)\""
        assert app._toner_too_short is None, "app._toner_too_short is None"
        # the shortest note that is kept: FREE_MIN_FRAMES whole-window frames
        feed(SHORTEST - 1, f0=261.63, note="C4")
        app._toner_process_capture_frame(silence())
        assert len(caps) == 1, "a note one frame too short was kept"
        note, heard, kept = app._toner_too_short
        assert note == "C4" and abs(heard - (SHORTEST - 2) * STEP) < 1e-6 and kept == FREE_MIN_FRAMES - 1, app._toner_too_short
        feed(SHORTEST, f0=261.63, note="C4")
        app._toner_process_capture_frame(silence())
        assert len(caps) == 2 and caps[1]["note"] == "C4" and caps[1]["n_frames"] == FREE_MIN_FRAMES, caps[1:]
        assert app._toner_too_short is None, "app._toner_too_short is None"
        # a change of note closes the capture of the one before it
        feed(40, f0=F0, note="A#3")
        feed(40, f0=261.63, note="C4")
        assert len(caps) == 3 and caps[2]["note"] == "A#3" and caps[2]["n_frames"] == 16, "len(caps) == 3 and caps[2][\"note\"] == \"A#3\" and caps[2][\"n_frames\"] == 16"
        assert app._toner_stable_note == "C4" and len(app._toner_free_accumulator) == 40, "app._toner_stable_note == \"C4\" and len(app._toner_free_accumulator) == 40"
        # nothing on disk yet: the save is deferred
        assert len(on_disk()["My Presets"]["Bench"]["sessions"]) == 1, "len(on_disk()[\"My Presets\"][\"Bench\"][\"sessions\"]) == 1"
        # stop: the note in progress is kept, the save flushed, the session written once with its rolloff rate
        app._toner_stop_capture()
        assert len(caps) == 4 and caps[3]["note"] == "C4" and app._toner_capture_state is None, "len(caps) == 4 and caps[3][\"note\"] == \"C4\" and app._toner_capture_state is None"
        assert app._toner_save_pending is False and app._toner_capture_btn.cget("text") == "Capture", "app._toner_save_pending is False and app._toner_capture_btn.cget(\"text\") == \"Capture\""
        sessions = on_disk()["My Presets"]["Bench"]["sessions"]
        assert len(sessions) == 2 and sessions[1]["date"] == session["date"], [s["date"] for s in sessions]
        assert len(sessions[1]["captures"]) == 4 and "rolloff_rate" not in sessions[1], \
            "seven harmonics give no rolloff rate (compute_rolloff_rate needs eight)"
    check("free capture: a held note's whole-window frames, one capture per held note, at least 10 of them (a "
          "shorter note is remembered as too short), the keys written, the deferred save, the stop flushing to disk",
          free_capture_rules)

    def a_session_is_saved_once_per_date():
        session = start_capture("free")
        feed(SHORTEST)
        app._toner_process_capture_frame(silence())
        app._toner_save_active_session()
        app._toner_save_active_session()
        sessions = on_disk()["My Presets"]["Bench"]["sessions"]
        assert [s["date"] for s in sessions].count(session["date"]) == 1, "saved twice"
        assert sessions[-1]["captures"][0]["n_frames"] == FREE_MIN_FRAMES, "sessions[-1][\"captures\"][0][\"n_frames\"] == FREE_MIN_FRAMES"
        app._toner_stop_capture()
        assert len(on_disk()["My Presets"]["Bench"]["sessions"]) == 3, "len(on_disk()[\"My Presets\"][\"Bench\"][\"sessions\"]) == 3"
    check("a session is replaced by its date on every save, never appended twice", a_session_is_saved_once_per_date)

    def calibration_rules():
        session = start_capture("calibration")
        assert app._toner_capture_state == "cal_countdown", "app._toner_capture_state == \"cal_countdown\""
        notes = app._toner_cal_notes
        assert notes[0] == "A#3" and len(notes) == 32, notes       # a tenor sounds every calibration note
        feed(1)
        assert app._toner_capture_state == "cal_countdown", "app._toner_capture_state == \"cal_countdown\""
        clock.t += toner_tab.CAL_COUNTDOWN_S                         # the wall clock does not count...
        app._toner_process_capture_frame(silence())
        assert app._toner_capture_state == "cal_countdown", "the countdown runs on audio time, not the wall clock"
        tick(toner_tab.CAL_COUNTDOWN_S)                              # ...the sound handed over does
        app._toner_process_capture_frame(silence())
        assert app._toner_capture_state == "calibration" and app._toner_capture_label.cget("text") == "Play Bb3 (1/32)", "app._toner_capture_state == \"calibration\" and app._toner_capture_label.cget(\"text\") == \"Play Bb3 (1/32)\""
        # a frame under the capture threshold does not start the recording
        g_sharp2 = 103.83
        app._toner_process_capture_frame(frame(g_sharp2, "G#2", level=0.4))
        assert app._toner_cal_recording is False, "app._toner_cal_recording is False"
        # the recording starts on the first frame over it; a frame is kept once its whole window is inside
        # the note (WINDOW_S after the start), every frame over the threshold from there until the 5 s are up
        app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
        assert app._toner_cal_recording is True, "app._toner_cal_recording is True"
        kept = 0
        k = 0
        while app._toner_capture_state == "calibration":
            k += 1
            tick()
            app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
            if k * STEP >= WINDOW_S:
                kept += 1
        assert app._toner_capture_state == "cal_pause", app._toner_capture_state
        assert k * STEP > toner_tab.CALIBRATION_DURATION_S and (k - 1) * STEP <= toner_tab.CALIBRATION_DURATION_S, "k * STEP > toner_tab.CALIBRATION_DURATION_S and (k - 1) * STEP <= toner_tab.CALIBRATION_DURATION_S"
        caps = session["captures"]
        assert len(caps) == 1, caps
        c = caps[0]
        assert set(c) == {"note", "written_note", "detected_as", "fundamental_freq", "harmonics_db", "harmonic_cents",
                          "timestamp", "n_frames", "method", "signal_level", "level_dbfs", "spectral_centroid"}, sorted(c)
        # stored under the note ASKED for (concert G#2 for written A#3 on a tenor), with what was detected
        assert c["note"] == "G#2" and c["written_note"] == "A#3" and c["detected_as"] == "A#3" and c["method"] == "calibration", "c[\"note\"] == \"G#2\" and c[\"written_note\"] == \"A#3\" and c[\"detected_as\"] == \"A#3\" and c[\"method\"] == \"calibra..."
        assert c["n_frames"] == kept == 141, (c["n_frames"], kept)
        assert c["fundamental_freq"] == g_sharp2, "c[\"fundamental_freq\"] == g_sharp2"
        assert len(on_disk()["My Presets"]["Bench"]["sessions"]) == 4, "a calibration capture is saved at once"
        # the next note waits for silence
        app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
        assert app._toner_capture_state == "cal_pause", "app._toner_capture_state == \"cal_pause\""
        app._toner_process_capture_frame(silence())
        assert app._toner_capture_state == "calibration" and app._toner_cal_index == 1, "app._toner_capture_state == \"calibration\" and app._toner_cal_index == 1"
        assert app._toner_capture_label.cget("text") == "Play B3 (2/32)", "app._toner_capture_label.cget(\"text\") == \"Play B3 (2/32)\""
        # the detector disagreeing changes nothing: the capture is filed under the asked note, detected_as says
        # what was heard; and a note that stops inside the 5 s loses the frames whose window reaches past its end
        a2 = 110.0
        app._toner_process_capture_frame(frame(a2, "A2"))
        held = 60                                                      # about 2 s of sound, then silence
        for _k in range(held):
            tick()
            app._toner_process_capture_frame(frame(a2, "A2"))
        while app._toner_capture_state == "calibration":
            tick()
            app._toner_process_capture_frame(silence())
        times = [k * STEP for k in range(1, held + 1)]
        expected = sum(1 for t in times if t >= WINDOW_S and t <= times[-1] - WINDOW_S)
        assert len(caps) == 2 and caps[1]["note"] == "A2" and caps[1]["written_note"] == "B3", "len(caps) == 2 and caps[1][\"note\"] == \"A2\" and caps[1][\"written_note\"] == \"B3\""
        assert caps[1]["detected_as"] == "B3" and caps[1]["fundamental_freq"] == a2, "caps[1][\"detected_as\"] == \"B3\" and caps[1][\"fundamental_freq\"] == a2"
        assert caps[1]["n_frames"] == expected == 37, (caps[1]["n_frames"], expected)
        app._toner_stop_capture()
        assert app._toner_capture_state is None, "app._toner_capture_state is None"
    check("calibration: a 10 s countdown and a 5 s recording on audio time, the frames from WINDOW_S in, the release "
          "dropped when the note stops early, stored under the note asked for with written_note and detected_as, "
          "saved at once, the next note after silence", calibration_rules)

    def a_guided_note_knows_its_stretch_of_the_recording():
        """With a WAV being recorded, each guided capture carries wav_start_s / wav_end_s: where in
        the recording its frames were read from (2026-11-10), so the
        recording can be read again under the asked names (analyze_asked_notes)."""
        session = start_capture("calibration")
        e = app._toner_engine
        tick(2.0)                                                  # 2 s of sound before the recording starts
        e.start_recording()                                        # the recording's clock starts here
        assert e.recording_time_s == 0.0, "e.recording_time_s == 0.0"
        rec0 = e.audio_time_s
        tick(toner_tab.CAL_COUNTDOWN_S + 0.1)
        app._toner_process_capture_frame(silence())
        assert app._toner_capture_state == "calibration", "app._toner_capture_state == \"calibration\""
        g_sharp2 = 103.83
        app._toner_process_capture_frame(frame(g_sharp2, "G#2"))   # the recording of this note starts now
        t_start = e.audio_time_s - rec0
        while app._toner_capture_state == "calibration":
            tick()
            app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
        t_end = e.audio_time_s - rec0                              # the 5 s ran out: the end is now
        c = session["captures"][0]
        assert c["wav_start_s"] == round(t_start, 3) and c["wav_end_s"] == round(t_end, 3), (c["wav_start_s"], c["wav_end_s"], t_start, t_end)
        assert 4.9 < c["wav_end_s"] - c["wav_start_s"] < 5.2, "4.9 < c[\"wav_end_s\"] - c[\"wav_start_s\"] < 5.2"
        # a note that stops early ends at its last sound, not at the 5 s
        app._toner_process_capture_frame(silence())
        assert app._toner_capture_state == "calibration" and app._toner_cal_index == 1, "app._toner_capture_state == \"calibration\" and app._toner_cal_index == 1"
        a2 = 110.0
        app._toner_process_capture_frame(frame(a2, "A2"))
        t2_start = e.audio_time_s - rec0
        for _k in range(60):
            tick()
            app._toner_process_capture_frame(frame(a2, "A2"))
        t2_last = e.audio_time_s - rec0
        while app._toner_capture_state == "calibration":
            tick()
            app._toner_process_capture_frame(silence())
        c2 = session["captures"][1]
        assert c2["wav_start_s"] == round(t2_start, 3) and c2["wav_end_s"] == round(t2_last, 3), (c2, t2_start, t2_last)
        # with no recording there are no stretches (the next session, recording stopped)
        e.stop_recording()
        assert e.recording_time_s is None, "e.recording_time_s is None"
        app._toner_stop_capture()
        session = start_capture("calibration")
        tick(toner_tab.CAL_COUNTDOWN_S + 0.1)
        app._toner_process_capture_frame(silence())
        app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
        while app._toner_capture_state == "calibration":
            tick()
            app._toner_process_capture_frame(frame(g_sharp2, "G#2"))
        assert "wav_start_s" not in session["captures"][0] and "wav_end_s" not in session["captures"][0], "\"wav_start_s\" not in session[\"captures\"][0] and \"wav_end_s\" not in session[\"captures\"][0]"
        app._toner_stop_capture()
    check("a guided note records where in the WAV its frames were read from (wav_start_s / wav_end_s on the "
          "recording's clock; a note that stops early ends at its last sound; none without a recording)",
          a_guided_note_knows_its_stretch_of_the_recording)

    def the_wav_reanalysis_reads_a_guided_session_under_the_asked_names():
        """The session's WAV re-analysis: a guided session whose captures all carry their stretch is read
        again with analyze_asked_notes, each note under the name asked for (its method, written_note and
        detected_as kept); any other session goes through analyze_audio_file as before. Both readers stood in."""
        calls = []

        def fake_asked(path, engine, asked, progress_cb=None):
            calls.append(("asked", path, [a["note"] for a in asked]))
            out = []
            for a in asked:
                out.append(dict(a, fundamental_freq=a["fundamental_freq"] * 1.001, n_frames=99, method="calibration",
                                detected_note=a["note"], level_dbfs=-20.0))
            out[-1] = None                       # one note the recording did not read: the live capture stays
            return out

        def fake_by_ear(path, engine, progress_cb=None):
            calls.append(("ear", path))
            return [{"note": "A#3", "fundamental_freq": F0, "harmonics_db": list(HARM), "harmonic_cents": [0.0] * len(HARM),
                     "n_frames": 40, "method": "file", "timestamp": ""}]

        real_asked, real_ear = toner_tab.analyze_asked_notes, toner_tab.analyze_audio_file
        toner_tab.analyze_asked_notes, toner_tab.analyze_audio_file = fake_asked, fake_by_ear
        try:
            # a guided session with two captures carrying their stretches
            session = start_capture("calibration")
            app._toner_active_session["captures"] = [
                dict(_capture("G#2", 103.83, method="calibration", written_note="A#3", detected_as="A#3",
                              wav_start_s=1.0, wav_end_s=6.0)),
                dict(_capture("A2", 110.0, method="calibration", written_note="B3", detected_as="B3",
                              wav_start_s=8.0, wav_end_s=10.5)),
            ]
            app._toner_capture_state = None
            app._toner_reanalyze_from_wav("guided.wav")
            assert calls == [("asked", "guided.wav", ["G#2", "A2"])], calls
            caps = app._toner_active_session["captures"]
            assert [c["note"] for c in caps] == ["G#2", "A2"] and [c["method"] for c in caps] == ["calibration"] * 2, "[c[\"note\"] for c in caps] == [\"G#2\", \"A2\"] and [c[\"method\"] for c in caps] == [\"calibration\"] * 2"
            assert caps[0]["n_frames"] == 99 and caps[0]["written_note"] == "A#3" and caps[0]["detected_as"] == "A#3", "caps[0][\"n_frames\"] == 99 and caps[0][\"written_note\"] == \"A#3\" and caps[0][\"detected_as\"] == \"A#3\""
            assert caps[1]["n_frames"] == 30, "the note the recording did not read keeps its live capture"
            assert all(c.get("timestamp") for c in caps), "all(c.get(\"timestamp\") for c in caps)"
            assert session is app._toner_active_session, "session is app._toner_active_session"
            # a session without the stretches (an older guided session, or free mode) is read by ear
            calls.clear()
            app._toner_active_session["captures"] = [dict(_capture("G#2", 103.83, method="calibration", written_note="A#3"))]
            app._toner_reanalyze_from_wav("free.wav")
            assert calls == [("ear", "free.wav")], calls
            caps = app._toner_active_session["captures"]
            assert [c["note"] for c in caps] == ["A#3"] and caps[0]["method"] == "file", "[c[\"note\"] for c in caps] == [\"A#3\"] and caps[0][\"method\"] == \"file\""
        finally:
            toner_tab.analyze_asked_notes, toner_tab.analyze_audio_file = real_asked, real_ear
            app._toner_active_session = None
            app._toner_capture_state = None
            app._toner_free_accumulator = []
    check("the WAV re-analysis reads a guided session under the asked names (method, written_note and detected_as "
          "kept; a note the recording did not read keeps its live capture) and anything else by ear",
          the_wav_reanalysis_reads_a_guided_session_under_the_asked_names)

    def pause_freezes_either_mode():
        start_capture("free")
        feed(20)
        app._toner_toggle_pause()
        assert app._toner_capture_state == "paused" and app._toner_paused, app._toner_capture_state
        assert app._toner_pause_btn.cget("text") == "Resume", app._toner_pause_btn.cget("text")
        n = len(app._toner_free_accumulator)
        feed(10)
        assert len(app._toner_free_accumulator) == n, "frames were taken while paused"
        app._toner_toggle_pause()
        assert app._toner_capture_state == "free_listening", app._toner_capture_state
        assert app._toner_pause_btn.cget("text") == "Pause", app._toner_pause_btn.cget("text")
        app._toner_free_accumulator = []
        app._toner_stop_capture()
    check("pause freezes the state machine and resume restores it", pause_freezes_either_mode)

    def the_exit_path_keeps_a_note_in_progress():
        session = start_capture("free")
        feed(40)
        app._toner_save_settings()                 # what on_exit calls
        assert len(session["captures"]) == 1 and session["captures"][0]["n_frames"] == 16, "len(session[\"captures\"]) == 1 and session[\"captures\"][0][\"n_frames\"] == 16"
        disk = on_disk()["My Presets"]["Bench"]["sessions"]
        assert disk[-1]["date"] == session["date"] and len(disk[-1]["captures"]) == 1, "disk[-1][\"date\"] == session[\"date\"] and len(disk[-1][\"captures\"]) == 1"
        ts = app.settings["toner_settings"]
        assert ts["sax_type"] == "Tenor" and ts["reference_pitch"] == 440.0 and "analysis_descriptors" in ts, "ts[\"sax_type\"] == \"Tenor\" and ts[\"reference_pitch\"] == 440.0 and \"analysis_descriptors\" in ts"
        app._toner_free_accumulator = []
        app._toner_stop_capture()
    check("the exit path saves a free-mode note in progress and the session, and merges the toner settings",
          the_exit_path_keeps_a_note_in_progress)

    def the_library_on_disk_reads_back_whole():
        disk = on_disk()
        bench = disk["My Presets"]["Bench"]
        # the seeded session and the four that made a capture; the paused session made none and was never written
        # the seeded session, the four that captured above, the two of the recording-stretch check and the
        # one the WAV re-analysis check filled
        assert len(bench["sessions"]) == 8, [len(s["captures"]) for s in bench["sessions"]]
        assert bench["sessions"][0]["captures"][0]["note"] == "A#3", "bench[\"sessions\"][0][\"captures\"][0][\"note\"] == \"A#3\""
        for s in bench["sessions"][1:]:
            assert set(s) >= {"date", "captures", "mic_type", "mic_model", "horn_type", "horn_make", "horn_model",
                              "serial", "player", "mouthpiece", "reed", "mic_position"}, sorted(s)
        assert copy.deepcopy(disk) == disk, "copy.deepcopy(disk) == disk"
    check("the library on disk reads back whole: the seeded session and the four that captured (a session with no "
          "capture is never written), each with the preset's details copied in", the_library_on_disk_reads_back_whole)

    toner_tab.time = time

    # No microphone: the error must survive a spectrum rebuild.
    app._toner_engine.synthetic_hz = None
    app._toner_engine.start = lambda device=None: (False, "no microphone (test)")
    app._toner_start()
    app._toner_build_spectrum_bars()                 # the rebuild that used to wipe it
    root.update_idletasks()

    def no_mic_error_survives_rebuild():
        c = app._toner_spectrum_canvas
        items = c.find_withtag("error")
        assert items, "audio-error message gone after the spectrum rebuild"
        assert "no microphone" in c.itemcget(items[0], "text"), "\"no microphone\" in c.itemcget(items[0], \"text\")"
    check("no-mic: 'Audio error' stays on the spectrum canvas after a rebuild", no_mic_error_survives_rebuild)
    root.destroy()

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
