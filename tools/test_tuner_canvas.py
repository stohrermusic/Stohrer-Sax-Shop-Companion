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

Which way the stripes travel is read from the picture itself in both
passes (2026-11-25, from JustATuner): the canvas items' own coordinates in
pass 1, screen pixels of the live GPU surface in pass 2. A sharp note
drifts right in all twelve windows. Both passes also time the tab's own
frame loop: the 60 setting gives 60 frames a second.

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
    # THE RULE: which way the stripes travel across each window for a sharp note. Both
    # renderers turn a disc CLOCKWISE for a rising phase: a stripe sits at angle a (0 at the
    # top of the disc, clockwise, as the shader's atan2(x, y) counts) where a - phase is
    # constant, and StrobeWheel.update draws base - phase in the canvas's counter-clockwise
    # angles, which is the same turn. (Until 2026-11-25 the canvas drew base + phase, the
    # opposite turn, and this comment said the two counted alike; the travel gates below read
    # the drawn picture instead of trusting it.) On screen a stripe's x is sin(a), so at the
    # middle of a window centred on angle c it travels right at cos(c) times the turn. A sharp
    # note raises the engine's phase. Right is positive, in all twelve windows.
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


# ---------------------------------------------------------------------
# Which way the stripes travel, and how fast frames come (2026-11-25)
# ---------------------------------------------------------------------
# The engine RAISES a wheel's phase for a sharp note (test_tuner_engine).
# The rule, from the picture: while the phase rises, the stripes in the
# middle of every one of the twelve windows travel to the RIGHT. Ported from
# JustATuner, where the canvas was measured turning the opposite way to the
# GPU (+3 deg of phase: GPU +3.1 deg clockwise, canvas -3.1).
NOTE_NAMES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
MEASURED_RING = 3          # 32 segments: 22.5 degrees per stripe pair
PHASE_STEP = 3.0           # degrees; well under half a stripe pair


def _all_wheels(phase_deg):
    """Stand-in for engine.analyze: every wheel lit, every phase at phase_deg."""
    from tuner_engine import TunerResult, NUM_RINGS

    def analyze():
        r = TunerResult()
        r.magnitudes = [10.0] * 12
        r.active = [True] * 12
        r.phase_offsets = [phase_deg] * 12
        r.ring_magnitudes = [[10.0] * NUM_RINGS for _ in range(12)]
        r.ring_phase_offsets = [[phase_deg] * NUM_RINGS for _ in range(12)]
        return r
    return analyze


def _show_phase(root, app, phase_deg, drawn):
    """Run the tab's own animate loop until it has drawn phase_deg."""
    app._tuner_engine.analyze = _all_wheels(phase_deg)
    for _ in range(20):
        root.after(150, root.quit)
        root.mainloop()
        if drawn():
            return
    raise AssertionError(f"the tab never drew phase {phase_deg}")


def _canvas_stripes(app):
    """{pitch class: {polygon id: x}} for the stripes of one ring that sit in
    the middle of that wheel's window, from the canvas's own coordinates."""
    import math
    out = {}
    c = app._tuner_canvas
    for pc, wheel in enumerate(app._tuner_wheels):
        xs = {}
        for poly_id, _base, _span, _steps in wheel._segments[MEASURED_RING]:
            xy = c.coords(poly_id)
            x, y = sum(xy[0::2]) / (len(xy) / 2), sum(xy[1::2]) / (len(xy) / 2)
            ang = math.degrees(math.atan2(-(y - wheel.cy), x - wheel.cx))     # counter-clockwise from east
            if abs((ang - wheel._wedge_center + 180.0) % 360.0 - 180.0) < 20.0:
                xs[poly_id] = x
        out[pc] = xs
    return out


def _canvas_travel(root, app, p0, p1):
    """Mean x travel (px) of the stripes in each window as the phase goes p0 -> p1."""
    def drawn(p):
        return lambda: all(abs(abs(w._last_ring_phases[MEASURED_RING]) - p) < 1e-6 for w in app._tuner_wheels)
    _show_phase(root, app, p0, drawn(p0))
    a = _canvas_stripes(app)
    _show_phase(root, app, p1, drawn(p1))
    b = _canvas_stripes(app)
    travel = {}
    for pc in range(12):
        dxs = [b[pc][i] - a[pc][i] for i in a[pc] if i in b[pc]]
        assert dxs, f"no stripe of ring {MEASURED_RING} in the window of pitch class {pc}"
        assert len({d > 0 for d in dxs}) == 1, f"pitch class {pc}: stripes in one window travel both ways {dxs}"
        travel[pc] = sum(dxs) / len(dxs)
    return travel


def _screen_ring(img, scale, ox, oy, cx, cy, radius, centre_deg):
    """Green level along one ring, 0.1 degree steps over +-30 degrees round the
    window's centre; the angle is clockwise from the top OF THE SCREEN."""
    import math
    import tuner_tab as tv
    gap = radius * tv.CENTER_GAP_FRACTION
    r = gap + (MEASURED_RING + 0.5) * (radius - gap) / tv.NUM_RINGS
    out = []
    for k in range(-300, 301):
        th = math.radians(centre_deg + k * 0.1)
        x = (ox + cx + r * math.sin(th)) * scale
        y = (oy + cy - r * math.cos(th)) * scale
        out.append(img.getpixel((int(round(x)), int(round(y))))[1])
    return out


def _clockwise_turn(a, b):
    """How far (degrees, clockwise on screen) profile b is profile a turned."""
    ma, mb = sum(a) / len(a), sum(b) / len(b)
    best = None
    for lag in range(-80, 81):
        s = n = 0
        for i in range(len(a)):
            j = i + lag
            if 0 <= j < len(b):
                s += (a[i] - ma) * (b[j] - mb)
                n += 1
        if best is None or s / n > best[0]:
            best = (s / n, lag)
    return best[1] * 0.1


def _gpu_travel(root, app, p0, p1):
    """{pitch class: travel to the right, in degrees of turn} read from screen
    pixels of the live GPU surface, or a string saying why the screen could
    not be read."""
    try:
        from PIL import ImageGrab
    except ImportError as e:
        return f"Pillow's ImageGrab is not available ({e})"
    frame = app._tuner_gpu_frame
    layout = app._tuner_compute_layout(frame.winfo_width(), frame.winfo_height())
    root.lift()
    root.attributes("-topmost", True)
    shots = []
    try:
        for p in (p0, p1):
            _show_phase(root, app, p, lambda: app._tuner_gpu_fail_count == 0)
            try:
                shots.append(ImageGrab.grab())
            except Exception as e:  # noqa: BLE001 — no screen to grab is a skip, said out loud
                return f"the screen could not be grabbed ({e})"
    finally:
        root.attributes("-topmost", False)
    # Tk reports logical pixels, a grab is physical (a scaled display).
    scale = shots[0].width / root.winfo_screenwidth()
    ox, oy = frame.winfo_rootx(), frame.winfo_rooty()
    travel = {}
    for pc, cx, cy, radius, is_up in layout:
        centre = 0.0 if is_up else 180.0
        a = _screen_ring(shots[0], scale, ox, oy, cx, cy, radius, centre)
        b = _screen_ring(shots[1], scale, ox, oy, cx, cy, radius, centre)
        if max(a) - min(a) < 60:
            return f"no stripes on screen in the window of pitch class {pc} (is the window covered, or off screen?)"
        cw = _clockwise_turn(a, b)
        travel[pc] = cw if is_up else -cw      # clockwise is rightward at the top of a disc, leftward at the bottom
    return travel


def _wrong_way(travel, want_right):
    return [NOTE_NAMES[pc] for pc in sorted(travel) if (travel[pc] > 0) != want_right]


def _loop_rate(root, app, seconds=2.0):
    """(median fps, mean fps, mean ms a frame keeps Tk busy) of the tab's own
    animate loop, timed round the real method. Busy is the frame's work plus
    the redraws Tk does for it afterwards as idle work, which on the canvas
    path is most of the cost."""
    import time
    stamps, work = [], []
    real = app._tuner_animate
    idle = {"pending": False}

    def idle_ran():
        idle["pending"] = False
        work.append(time.perf_counter() - stamps[-1])

    def timed():
        stamps.append(time.perf_counter())
        # Every frame's redraws must run before the next frame starts,
        # however short the wait the pacer hands back.
        assert not idle["pending"], "a frame started before Tk reached the last frame's idle work (redraws starved)"
        idle["pending"] = True
        real()
        root.after_idle(idle_ran)       # queued behind this frame's redraws
    app._tuner_animate = timed
    root.after(int(seconds * 1000), root.quit)
    root.mainloop()
    app.__dict__.pop("_tuner_animate", None)
    gaps = sorted(b - a for a, b in zip(stamps, stamps[1:]))
    assert len(gaps) >= 5, f"only {len(stamps)} frames in {seconds} s"
    # The median frame, not the mean: one stall (a screen grab just before,
    # another program) is not the loop's pace.
    return 1.0 / gaps[len(gaps) // 2], len(gaps) / sum(gaps), 1000.0 * sum(work) / len(work)


def _check_loop_rate(label, root, app):
    """The 60 setting gives 60 frames a second wherever a frame leaves room
    for it (2026-11-25: JustATuner measured 32 on Windows, from the same
    loop as SSC's). Where a frame keeps Tk busy for over 85 % of the period
    (the canvas renderer, a slow CI runner) the rate is printed, not judged."""
    app._tuner_fps_var.set("60")
    _run_frames(root, app, seconds=0.5, min_frames=0)
    fps, mean_fps, work_ms = _loop_rate(root, app)
    if sys.platform == "darwin" and not app._tuner_use_gpu:
        # Aqua draws the canvas when the window is displayed, outside Tk's idle
        # work, so the busy figure cannot say whether a frame leaves room for
        # 60. Measured on the macOS runner, 2026-11-25: busy 3-7 ms against
        # frames of 17-35 ms, while a plain paced canvas ran 57-61 fps there
        # and the paced loop beat the old after(16) loop in every pairing
        # (1.1-1.7x). The rate is reported, not judged.
        print(f"  SKIP  {label}: {fps:.1f} fps median on the 60 setting on macOS (Tk busy {work_ms:.1f} ms per "
              f"frame, but Aqua's drawing is not in that figure, so it is not judged)")
        return
    if work_ms > 0.85 * 1000.0 / 60.0:
        print(f"  SKIP  {label}: {mean_fps:.0f} fps on the 60 setting with Tk busy {work_ms:.1f} ms per frame "
              f"(over 85 % of the 16.7 ms period: the loop cannot keep up here)")
        return
    check(f"{label}: the 60 setting runs at 60 frames a second (median frame {fps:.1f} fps, mean {mean_fps:.1f}, "
          f"Tk busy {work_ms:.1f} ms per frame)", lambda: 55.0 <= fps <= 65.0)


def pacer_on_a_fake_clock():
    import frame_pacing as tv
    p = tv.FramePacer()
    period = 1.0 / 60.0
    # a frame whose work ended 4 ms after its start: wait the rest of the period, not a fixed 16 ms
    assert p.next_delay_ms(period, now=10.000) == 17              # first frame: a full period from now
    assert p.next_delay_ms(period, now=10.0167 + 0.004) == 13, "the work was not taken off the wait"
    assert p.next_delay_ms(period, now=10.0333 + 0.012) == 5
    # a frame that overran (20 ms): the next follows at once, on the 1 ms floor, never 0
    assert p.next_delay_ms(period, now=10.0500 + 0.020) == 1
    # far behind (a stall): one frame at once, then the period again; no burst of catch-up frames
    assert p.next_delay_ms(period, now=10.500) == 1
    assert p.next_delay_ms(period, now=10.501 + 0.004) == 12
    p.reset()
    t = 0.0
    for _ in range(600):                                           # ten seconds of 4 ms frames, whole-ms waits
        t += 0.004
        t += p.next_delay_ms(period, now=t) / 1000.0
    assert abs(600 / t - 60.0) < 0.2, f"{600 / t:.2f} fps over ten simulated seconds"


def pacer_on_a_real_timer():
    import time
    import tkinter as tk
    import frame_pacing as tv
    root = tk.Tk()
    root.withdraw()
    holds_before = tv._fine_timer_holds
    held = tv.fine_timer(True)
    pacer, stamps = tv.FramePacer(), []

    def frame():
        stamps.append(time.perf_counter())
        while time.perf_counter() - stamps[-1] < 0.004:            # 4 ms of work, as the GPU tab does
            pass
        root.after(pacer.next_delay_ms(1.0 / 60.0), frame)
    frame()
    root.after(2000, root.quit)
    root.mainloop()
    root.destroy()
    tv.fine_timer(False)
    assert held == (sys.platform == "win32"), held
    assert tv._fine_timer_holds == holds_before, (tv._fine_timer_holds, holds_before)
    gaps = sorted(b - a for a, b in zip(stamps, stamps[1:]))
    fps = 1.0 / gaps[len(gaps) // 2]
    # Windows, where the 1 ms tick is what this proves: 57. Elsewhere timers are
    # already fine but jitter more: 54.9-60.8 measured on the macOS runner.
    floor = 57.0 if sys.platform == "win32" else 50.0
    assert floor <= fps <= 63.0, f"median frame {fps:.1f} fps (worst frame {1000 * gaps[-1]:.0f} ms)"


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

    def canvas_sharp_right():
        t = _canvas_travel(root, app, 20.0, 20.0 + PHASE_STEP)
        wrong = _wrong_way(t, True)
        assert not wrong, f"a sharp note's stripes travel LEFT in the windows of {wrong}"
    check("canvas: a sharp note's stripes travel RIGHT in all twelve windows", canvas_sharp_right)

    def canvas_flat_left():
        t = _canvas_travel(root, app, 20.0 + PHASE_STEP, 20.0)
        wrong = _wrong_way(t, False)
        assert not wrong, f"a flat note's stripes travel RIGHT in the windows of {wrong}"
    check("canvas: a flat note's stripes travel LEFT in all twelve windows", canvas_flat_left)
    app._tuner_engine.__dict__.pop("analyze", None)      # the engine's own analyze again
    _check_loop_rate("canvas", root, app)

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

        if app2._tuner_use_gpu:
            t = _gpu_travel(root, app2, 20.0, 20.0 + PHASE_STEP)
            app2._tuner_engine.__dict__.pop("analyze", None)      # the engine's own analyze again
            if isinstance(t, str) and os.environ.get("SSC_REQUIRE_GPU") != "1":
                print(f"  SKIP  GPU: stripe travel from screen pixels: {t}")
            else:
                def gpu_sharp_right():
                    assert not isinstance(t, str), t
                    wrong = _wrong_way(t, True)
                    assert not wrong, f"a sharp note's stripes travel LEFT in the windows of {wrong}"
                    off = {NOTE_NAMES[pc]: round(v, 1) for pc, v in t.items() if abs(abs(v) - PHASE_STEP) > 1.0}
                    assert not off, f"the picture did not turn by the {PHASE_STEP} degrees it was given: {off}"
                check("GPU, from screen pixels: a sharp note's stripes travel RIGHT in all twelve windows",
                      gpu_sharp_right)
            _run_frames(root, app2, seconds=1.0)               # back on the tone for the readout below
            _check_loop_rate("GPU", root, app2)

            def fine_timer_held():
                import frame_pacing as tv
                assert tv._fine_timer_holds == (1 if sys.platform == "win32" else 0), tv._fine_timer_holds
            check("the 1 ms timer tick is held while the tuner runs (Windows) and by nothing else", fine_timer_held)

        def gpu_readout():
            note, cents, raw = _readout(app2)
            assert note.rstrip("0123456789") == "A", (note, raw)
            assert "IN TUNE" in raw or (cents is not None and abs(cents) < 1.0), (note, raw)
        check("GPU build: VU readout says A, in tune", gpu_readout)
        app2._tuner_stop()
        root.update_idletasks()

        def fine_timer_released():
            import frame_pacing as tv
            assert tv._fine_timer_holds == 0, tv._fine_timer_holds
        check("stopping the tuner releases the 1 ms timer tick", fine_timer_released)
        root.destroy()
    else:
        print("  SKIP  GPU pass: tuner_render not built here" +
              (" (macOS is canvas-only by design)" if sys.platform == "darwin" else ""))

    check("FramePacer: the wait is the period less the work, never 0, no catch-up burst, 60.0 fps over 600 frames",
          pacer_on_a_fake_clock)
    check("a Tk loop paced by FramePacer under fine_timer runs 4 ms frames at 60 a second; the hold is released",
          pacer_on_a_real_timer)

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
