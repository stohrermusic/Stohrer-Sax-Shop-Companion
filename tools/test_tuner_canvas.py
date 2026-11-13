"""Drive the Tuner tab end to end on a synthetic tone — no microphone needed.

The engine's `synthetic_hz` replaces the audio stream with a generated
440 Hz tone (plus a -6 dB 2nd harmonic), so the whole tab — tab switch,
engine start, the real `_tuner_animate` loop, StrobeWheel updates, the VU
readout — runs on a machine with no input device. That is what lets the
macOS CI runner exercise the tuner at all; on a dev box it also covers the
canvas path that a GPU machine otherwise never draws.

Two passes:
  1. Canvas mode (forced by clearing tuner_tab._HAS_GPU_RENDERER before the
     app is built): 12 wheels, the A wheel lit and the rest dark, readout
     "A" within 1 cent, no error overlay.
  2. GPU mode, only where the tuner_render wheel is built (never macOS):
     renderer alive after the frames, zero render failures, readout "A".

Frames run under a real mainloop with a quit timer. A `root.update()` loop
never returns here: a canvas frame outlasts its 16 ms interval, so update()
keeps servicing the already-due reschedule (25 s stall in canvas.coords).

Self-skips without a display. Uses an isolated config profile
(SAXSHOP_CONFIG_DIR) so it never reads or writes the user's settings.
"""
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-tuner-test-")
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


def _run_frames(root, app, seconds=3.0, min_frames=20):
    """Let the tab's real animate loop run under mainloop, then return.

    The GPU renderer's first-use setup can eat the first second, and the
    ring buffer holds the pre-tone silence until the engine has consumed
    four chunks, so the labels are only read once the engine has fed at
    least min_frames chunks (the readout is per frame, not damped)."""
    for _ in range(3):
        root.after(int(seconds * 1000), root.quit)
        root.mainloop()
        if app._tuner_engine._synth_pos >= min_frames * 1024:
            return
    raise AssertionError(f"engine processed only {app._tuner_engine._synth_pos // 1024} frames")


def _readout(app):
    note = app._vu_note_label.cget("text")
    cents_txt = app._vu_cents_label.cget("text")
    cents = None
    for tok in cents_txt.replace("¢", " ").replace("+", " ").split():
        try:
            cents = float(tok)
            break
        except ValueError:
            continue
    return note, cents, cents_txt


def _drive(root, gpu):
    import tuner_tab
    tuner_tab._HAS_GPU_RENDERER = gpu
    from main import PadSVGGeneratorApp
    root.geometry("1000x720+0+0")
    app = PadSVGGeneratorApp(root)
    root.update()
    assert getattr(app, "_tuner_engine", None) is not None, "tuner tab not built (audio libs missing?)"
    app._tuner_engine.synthetic_hz = 440.0
    app.notebook.select(app.tuner_tab_frame)   # on_tab_changed -> _tuner_start
    _run_frames(root, app)
    return app


def the_bottom_row_drifts_the_same_way():
    """A sharp note drifts right in every window (Matt, 2026-11-12: the bottom row moved the
    wrong way). The bottom row's wedge opens downward, where a disc turning the same way
    travels the other way across the window, so those discs are turned the other way."""
    import inspect
    import tuner_tab
    layout = tuner_tab.TunerTabMixin._tuner_compute_layout(None, 900, 500)
    assert layout and len(layout) == 12, layout
    bottom = {pc for pc, _cx, _cy, _r, is_up in layout if not is_up}
    assert bottom == set(tuner_tab.BOTTOM_ROW_PITCH_CLASSES) == {0, 2, 4, 5, 7, 9, 11}, bottom
    phases = [10.0, -20.0, 0.0, 3.5, 4.0, 5.0, 6.0]
    for pc in range(12):
        got = tuner_tab.window_phases(pc, list(phases))
        want = [-p for p in phases] if pc in bottom else phases
        assert got == want, (pc, got)
    # THE RULE, from the picture itself: which way the stripes travel across each window for a
    # sharp note. A stripe sits at angle a (0 at the top of the disc, clockwise, as the shader's
    # atan2(x, y) and the canvas both count) where a - phase is constant, so it turns with the
    # phase the renderer is given; on screen its x is sin(a), so at the middle of a window
    # centred on angle c it travels right at cos(c) times that turn. A sharp note raises the
    # engine's phase. Right is positive, in all twelve windows.
    import math
    engine_src = open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                                   "tuner_engine.py"), encoding="utf-8").read()
    assert "phase_offsets[pc] += cents" in engine_src, "the engine's phase no longer rises for a sharp note"
    for pc, _cx, _cy, _r, is_up in layout:
        window_centre = 0.0 if is_up else math.pi
        turn = tuner_tab.window_phases(pc, [1.0])[0]          # what the renderer gets for +1 from the engine
        travel = math.cos(window_centre) * turn
        assert travel > 0, (f"pitch class {pc} ({'top' if is_up else 'bottom'} row): a sharp note's stripes travel "
                            f"{'left' if travel < 0 else 'nowhere'} across its window")
    # both renderers take their phases through it, and the shader's own sign is untouched
    src = inspect.getsource(tuner_tab.TunerTabMixin)
    assert src.count("window_phases(i, ") == 3, src.count("window_phases(i, ")
    assert "ring_phases.append(list(result.ring_phase_offsets[i]))" not in src, "the GPU path sends raw phases"
    assert "rp = list(result.ring_phase_offsets[i])" not in src, "the canvas path sends raw phases"
    shader = open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                               "tuner_renderer", "src", "shader.wgsl"), encoding="utf-8").read()
    assert "angle - wheel.ring_phases[ring_idx]" in shader


def the_rescan_goes_round_the_engines():
    """main._rescan_audio_devices stops the running engines, rescans, starts them again, in that
    order; never during a capture; every device list in the app is read after a rescan."""
    import inspect
    import config
    import main as main_mod
    import toner_tab
    import tuner_tab
    calls = []

    class Engine:
        def __init__(self, running):
            self.is_running = running

    class App:
        _rescan_audio_devices = main_mod.PadSVGGeneratorApp._rescan_audio_devices
        _toner_capture_state = None

        def __init__(self, tuner_on, toner_on):
            self._tuner_engine, self._toner_engine = Engine(tuner_on), Engine(toner_on)

        def _tuner_stop(self):
            calls.append("tuner stop")

        def _tuner_start(self):
            calls.append("tuner start")

        def _toner_stop(self):
            calls.append("toner stop")

        def _toner_start(self):
            calls.append("toner start")
    keep = main_mod.rescan_input_devices
    main_mod.rescan_input_devices = lambda: calls.append("rescan") or True
    try:
        assert App(True, False)._rescan_audio_devices() is True
        assert calls == ["tuner stop", "rescan", "tuner start"], calls
        del calls[:]
        assert App(False, True)._rescan_audio_devices() is True
        assert calls == ["toner stop", "rescan", "toner start"], calls
        del calls[:]
        assert App(False, False)._rescan_audio_devices() is True
        assert calls == ["rescan"], calls
        del calls[:]
        capturing = App(False, True)
        capturing._toner_capture_state = "free_listening"
        assert capturing._rescan_audio_devices() is False
        assert calls == [], calls
    finally:
        main_mod.rescan_input_devices = keep
    # the three places that list devices rescan first, and nothing opens a raw saved index
    for mod, fn in ((main_mod.PadSVGGeneratorApp, "_open_input_device_dialog"),
                    (toner_tab.TonerTabMixin, "_toner_open_settings"),
                    (tuner_tab.TunerTabMixin, None)):
        src = inspect.getsource(getattr(mod, fn)) if fn else inspect.getsource(mod)
        a, b = src.find("self._rescan_audio_devices()"), src.find("devices = get_input_devices()")
        assert 0 <= a < b, f"{mod.__name__}.{fn or '*'}: the list is read before the rescan"
    for mod in (main_mod, toner_tab, tuner_tab):
        src = inspect.getsource(mod)
        assert 'settings.get("audio_input_device")' not in src, f"{mod.__name__} reads the saved index raw"
        assert 'settings["audio_input_device"] =' not in src, \
            f"{mod.__name__} writes the saved index without its name"
    assert callable(config.input_device_for) and callable(config.set_input_device)


def main():
    check("a device rescan stops the running engines, rescans and starts them again; never during a "
          "capture; every device list is read after one, and no saved index is opened raw",
          the_rescan_goes_round_the_engines)
    check("the bottom row's discs are turned the other way, so a sharp note drifts right in every window; "
          "both renderers go through window_phases", the_bottom_row_drifts_the_same_way)
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    from tuner_engine import AUDIO_AVAILABLE
    if not AUDIO_AVAILABLE:
        # The Intel Mac build ships without numpy/sounddevice; its Tuner tab
        # is the "not available on this Mac" panel by design.
        print("Skipping: audio libraries not installed (numpy/sounddevice) — no tuner tab on this build")
        root.destroy()
        return 0

    print("Tuner tab on a synthetic tone")
    print("=" * 60)
    import tuner_tab
    gpu_available = tuner_tab._HAS_GPU_RENDERER

    # ---- pass 0: no microphone — the error must survive a canvas rebuild ----
    import tuner_tab as _tt
    _tt._HAS_GPU_RENDERER = False
    from main import PadSVGGeneratorApp as _App
    root.geometry("1000x720+0+0")
    app0 = _App(root)
    root.update()
    app0._tuner_engine.start = lambda device=None: (False, "no microphone (test)")
    app0.notebook.select(app0.tuner_tab_frame)       # _tuner_start fails -> error text
    root.after(800, root.quit)
    root.mainloop()
    app0._tuner_build_wheels()                       # the rebuild that used to wipe it
    root.update_idletasks()

    def no_mic_error_survives_rebuild():
        items = app0._tuner_canvas.find_withtag("error")
        assert items, "audio-error message gone after the canvas rebuild"
        assert "no microphone" in app0._tuner_canvas.itemcget(items[0], "text")
    check("no-mic: 'Audio error' stays on the canvas after a rebuild", no_mic_error_survives_rebuild)
    root.destroy()
    root = tk.Tk()

    # ---- pass 1: canvas ---------------------------------------------
    app = _drive(root, gpu=False)
    check("tuner is running on the synthetic source (no stream opened)",
          lambda: app._tuner_running and app._tuner_engine.is_running and app._tuner_engine._stream is None)
    check("canvas mode with 12 StrobeWheels",
          lambda: (not app._tuner_use_gpu) and app._tuner_canvas is not None and len(app._tuner_wheels) == 12)
    check("no error overlay on the canvas",
          lambda: app._tuner_canvas.find_withtag("error") == ())

    def a_wheel_brightest_far_wheels_dark():
        # A pure tone also lights the two neighbouring wheels dimly (FFT
        # leakage at 10.8 Hz bins); that is how the app looks. The contract
        # is: A clearly brightest, everything not adjacent to A dark.
        b = [float(w._brightness) for w in app._tuner_wheels]
        assert b[9] > 0.5, f"A wheel brightness {b[9]:.2f}"
        others = [x for i, x in enumerate(b) if i != 9]
        assert b[9] >= 2.0 * max(others), f"A not clearly brightest: {[round(x, 2) for x in b]}"
        far = [x for i, x in enumerate(b) if i not in (8, 9, 10)]
        assert max(far) < 0.05, f"a non-adjacent wheel is lit: {[round(x, 2) for x in b]}"
    check("A wheel clearly brightest, non-adjacent wheels dark", a_wheel_brightest_far_wheels_dark)

    def readout_a_in_tune():
        note, cents, raw = _readout(app)
        assert note.rstrip("0123456789") == "A", f"note readout {note!r}"
        # 'IN TUNE' replaces the number inside ±4 c; a number means off by that much.
        assert "IN TUNE" in raw or (cents is not None and abs(cents) < 1.0), f"cents readout {raw!r}"
    check("VU readout says A, in tune", readout_a_in_tune)

    def engine_result_is_a_in_tune():
        r = app._tuner_engine.analyze()
        assert r.magnitudes[9] > 0.5, r.magnitudes
        assert abs(r.cents_errors[9]) < 0.2, r.cents_errors[9]
    check("engine reads A at 0 cents through the live path", engine_result_is_a_in_tune)

    app._tuner_stop()
    root.update_idletasks()
    check("stops cleanly", lambda: not app._tuner_running)
    root.destroy()

    # ---- pass 2: GPU, where the wheel exists --------------------------
    if gpu_available:
        root = tk.Tk()
        app2 = _drive(root, gpu=True)
        if app2._tuner_use_gpu:
            check("GPU mode: renderer alive after the frames",
                  lambda: app2._tuner_use_gpu and app2._gpu_renderer is not None)
            check("GPU mode: zero render failures",
                  lambda: app2._tuner_gpu_fail_count == 0)
        else:
            # CI runners have a software adapter (or no Vulkan): the tab
            # must have dropped to the canvas on purpose and said so.
            check("software adapter / no GPU: fell back to the canvas as designed",
                  lambda: app2._tuner_canvas is not None and len(app2._tuner_wheels) == 12
                  and app2._gpu_renderer is None and hasattr(app2, '_cpu_mode_lbl'))

        def gpu_readout():
            note, cents, raw = _readout(app2)
            assert note.rstrip("0123456789") == "A", (note, raw)
            assert "IN TUNE" in raw or (cents is not None and abs(cents) < 1.0), (note, raw)
        check("GPU build: VU readout says A, in tune", gpu_readout)
        app2._tuner_stop()
        root.update_idletasks()
        root.destroy()
    else:
        print("  SKIP  GPU pass: tuner_render not built here" +
              (" (macOS is canvas-only by design)" if sys.platform == "darwin" else ""))

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
