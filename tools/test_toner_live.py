"""Drive the Toner tab end to end on a synthetic tone — no microphone needed.

`TonerEngine.synthetic_hz` replaces the audio stream with a generated Bb3
(233.08 Hz) carrying six harmonics, so the whole tab runs on a machine
with no input device: the real `_toner_animate` loop, the spectrum bars,
the note and frequency readouts, the intonation gauge. This is the toner's
only end-to-end check on the macOS CI runner.

The toner is hidden behind the beta terms, so the test hands the app a
prepared profile through SAXSHOP_CONFIG_DIR (toner unlocked, tab visible)
— which also keeps it away from the user's real settings. The first-start
microphone notice is a modal; `_toner_mic_checked` is set so it never opens.

Frames run under a real mainloop with a quit timer (see test_tuner_canvas
for why a root.update() loop can't be used). Self-skips without a display.
"""
import json
import math
import os
import re
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-toner-test-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
with open(os.path.join(_PROFILE, "app_settings.json"), "w", encoding="utf-8") as _f:
    json.dump({"toner_unlocked": True, "visible_tabs": {"Toner": True}}, _f)

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []
F0 = 233.08


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
    app.notebook.select(app.toner_tab_frame)   # on_tab_changed -> _toner_start
    root.after(2000, root.quit)
    root.mainloop()

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
        assert "no microphone" in c.itemcget(items[0], "text")
    check("no-mic: 'Audio error' stays on the spectrum canvas after a rebuild", no_mic_error_survives_rebuild)
    root.destroy()

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
