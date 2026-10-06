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

results = []


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


def main():
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
