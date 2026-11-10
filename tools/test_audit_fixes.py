"""Gates for the faults the exception audit listed and the 2026-11-10 bug round fixed. Each check
fails on the code as it was (the frozen tree a10c8dd8) and passes now; the message says what was
seen. The machine-facing ones first.

  1. CameraCalibrationDialog._engrave_check_position_safe: a position that cannot be read refuses
     the engrave/frame with a box ("Head Position Unknown") and a log line, instead of passing.
  2. main._on_machine_reset_falcon: one outcome, one box; the homed flag is cleared only by a reset
     that went out; a reset byte that could not be written is a failure.
  3. CameraCalibrationDialog._engrave_jog: the port is closed in a finally, so a jog that raises
     does not leave it open.
  4. FalconRunDialog: a jog the sender did not write does not unlock Start Frame; Pause and Resume
     flip their buttons only when the byte went out; a stand-in's None still counts as sent (so the
     suites' fake senders keep working). FalconSender.jog / send_realtime / pause / resume return
     the answer.
  5. CameraCaptureDialog and LiveCameraWindow refresh loops: a frame that fails to process is
     logged once with its traceback, counted, said at close, and the loop goes on.
  6. parse_pad_list: inf, 1e400 and nan as a size or quantity are rejected like any other bad line.
  7. toner_engine.load_tone_presets: a corrupt library is kept as .corrupt.bak; bad bytes and an
     unreadable path are a failed load, not a crash; a good library loads unchanged.
  8. pixels_to_mm with a short card_engrave_offset_mm; sleep_lock saying once that Windows refused;
     a non-numeric stored reference pitch (the tabs still build, the handlers swallow TclError).
  9. tools/run_tests.py: a hung suite is killed at the timeout with its children, and its output so
     far is returned; tools/test_tuner_engine.py bounds every call that touches the audio device.

Isolated profile (SAXSHOP_CONFIG_DIR); message boxes through box_guard; no real serial port, no
real camera (fakes stand in for falcon_sender.FalconSender and the capture objects). The GUI
checks self-skip without a display.
"""
import logging
import os
import sys
import tempfile
import threading
import time
import types

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))
_PROFILE = tempfile.mkdtemp(prefix="ssc-audit-fixes-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

import box_guard  # noqa: E402
guard = box_guard.install()

from i18n import init_translation  # noqa: E402
init_translation("en")

import falcon_sender  # noqa: E402
import ui_dialogs  # noqa: E402

results = []


def check(name, fn):
    try:
        ok = fn()
        if ok is False:
            raise AssertionError("returned False")
        box_guard.settle()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:  # noqa: BLE001
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        print(f"        boxes pending: {box_guard.reset()}")
        results.append(False)


class Catch(logging.Handler):
    """Collects the records of the named loggers while attached."""

    def __init__(self, *names):
        super().__init__(level=logging.DEBUG)
        self.records = []
        self.loggers = [logging.getLogger(n) for n in names]

    def emit(self, record):
        self.records.append(record)

    def __enter__(self):
        for lg in self.loggers:
            lg.addHandler(self)
            lg.setLevel(logging.DEBUG)
        return self

    def __exit__(self, *exc):
        for lg in self.loggers:
            lg.removeHandler(self)

    def messages(self, level=None):
        return [r.getMessage() for r in self.records if level is None or r.levelno >= level]


class FakeSender:
    """falcon_sender.FalconSender stand-in. Class attributes script it:
    connect_raises, status (what get_status returns), jog_raises,
    realtime_answer (what send_realtime returns)."""
    connect_raises = None
    status = {'mpos': [200.0, 200.0, 0.0], 'state': 'Idle'}
    jog_raises = None
    realtime_answer = True
    made = []

    def __init__(self, port=None, **kw):
        self.port = port
        self.connected = False
        self.disconnects = 0
        self.jogs = []
        self.realtime = []
        FakeSender.made.append(self)

    def connect(self):
        if FakeSender.connect_raises:
            raise FakeSender.connect_raises
        self.connected = True

    def disconnect(self):
        self.disconnects += 1
        self.connected = False

    def get_status(self, timeout=0.5):
        return FakeSender.status

    def jog(self, **kw):
        self.jogs.append(kw)
        if FakeSender.jog_raises:
            raise FakeSender.jog_raises
        return True

    def send_realtime(self, byte):
        self.realtime.append(byte)
        return FakeSender.realtime_answer


def with_fake_sender(fn, **script):
    saved = (falcon_sender.FalconSender, FakeSender.connect_raises, FakeSender.status,
             FakeSender.jog_raises, FakeSender.realtime_answer)
    falcon_sender.FalconSender = FakeSender
    for k, v in script.items():
        setattr(FakeSender, k, v)
    del FakeSender.made[:]
    try:
        return fn()
    finally:
        (falcon_sender.FalconSender, FakeSender.connect_raises, FakeSender.status,
         FakeSender.jog_raises, FakeSender.realtime_answer) = saved


# ---------------------------------------------------------------------------
# 1. The bed-bounds check refuses when the position cannot be read
# ---------------------------------------------------------------------------
Dialog = ui_dialogs.CameraCalibrationDialog


def bed_check_fake():
    fake = types.SimpleNamespace(
        _falcon_port="COM-T", _port_lock=threading.RLock(),     # @_holds_port takes the dialog's port lock
        _settings={"laser_bed_x_max": 400.0, "laser_bed_y_max": 415.0},
        _engrave_geometry=lambda: (6, 4, 20.0, 120.0, 80.0, 10.0, 8.0),
    )
    fake._position_unknown = types.MethodType(Dialog._position_unknown, fake)
    return fake


def a_check_that_cannot_be_made_does_not_pass():
    for script, reason in ((dict(connect_raises=OSError("port busy")), "port busy"),
                           (dict(status=None), "no status report"),
                           (dict(status={'state': 'Idle'}), "no status report")):
        del guard.shown[:]
        fake = bed_check_fake()
        with Catch("ssc.machine") as log:
            ok = with_fake_sender(lambda: Dialog._engrave_check_position_safe(fake), **script)
        assert ok is False, f"{script}: passed although the position could not be read"
        assert len(guard.shown) == 1 and guard.shown[0][0] == "showerror", guard.shown
        kind, title, message = guard.shown[0]
        assert title == "Head Position Unknown" and reason in message, (title, message)
        assert any("could not be read" in m for m in log.messages(logging.WARNING)), log.messages()
        assert FakeSender.made[0].disconnects == 1, "the port was not closed"
    del guard.shown[:]


def a_readable_position_still_decides_as_before():
    del guard.shown[:]
    fake = bed_check_fake()
    ok = with_fake_sender(lambda: Dialog._engrave_check_position_safe(fake))
    assert ok is True and not guard.shown, (ok, guard.shown)
    assert FakeSender.made[0].disconnects == 1
    # The off-bed refusal is unchanged: head at the home corner.
    fake = bed_check_fake()
    ok = with_fake_sender(lambda: Dialog._engrave_check_position_safe(fake),
                          status={'mpos': [5.0, 5.0, 0.0], 'state': 'Idle'})
    assert ok is False and guard.shown[-1][1] == "Card Would Run Off Bed", guard.shown
    del guard.shown[:]


print("=== 1. Bed-bounds check: a position that cannot be read refuses, never passes ===")
check("a connect failure or a missing status report refuses the engrave/frame with the "
      "'Head Position Unknown' box and a log line, and the port is closed",
      a_check_that_cannot_be_made_does_not_pass)
check("a readable position decides as before: on the bed passes silently, off the bed refuses",
      a_readable_position_still_decides_as_before)

# ---------------------------------------------------------------------------
# 2. Reset Falcon: one outcome, one box, the homed flag only on success
# ---------------------------------------------------------------------------
import main as main_mod  # noqa: E402

App = main_mod.PadSVGGeneratorApp


def reset_fake():
    return types.SimpleNamespace(_machine_require_falcon=lambda: True, falcon_port="COM-T",
                                 _falcon_homed_this_session=True)


def a_failed_reset_is_one_box_and_keeps_the_homed_flag():
    for script, said in ((dict(connect_raises=OSError("no port")), "no port"),
                         (dict(realtime_answer=False), "reset byte")):
        del guard.shown[:]
        guard.expect("Reset Falcon?", answer=True)
        fake = reset_fake()
        with Catch("ssc.machine") as log:
            with_fake_sender(lambda: App._on_machine_reset_falcon(fake), **script)
        boxes = [b for b in guard.shown if b[0] != "askyesno"]
        assert len(boxes) == 1, f"{script}: {len(boxes)} boxes after a failed reset: {boxes}"
        assert boxes[0][0] == "showerror" and boxes[0][1] == "Reset Failed" and said in boxes[0][2], boxes
        assert fake._falcon_homed_this_session is True, "a reset that did not happen cleared the homed flag"
        assert any("Reset Falcon failed" in m for m in log.messages(logging.WARNING)), log.messages()
        assert FakeSender.made[0].disconnects == 1
    del guard.shown[:]


def a_reset_that_went_out_is_one_box_and_clears_the_flag():
    del guard.shown[:]
    guard.expect("Reset Falcon?", answer=True)
    fake = reset_fake()
    with_fake_sender(lambda: App._on_machine_reset_falcon(fake))
    boxes = [b for b in guard.shown if b[0] != "askyesno"]
    assert len(boxes) == 1 and boxes[0][0] == "showinfo" and boxes[0][1] == "Falcon Reset", boxes
    assert fake._falcon_homed_this_session is False
    assert FakeSender.made[0].realtime == [falcon_sender.RT_SOFT_RESET] and FakeSender.made[0].disconnects == 1
    del guard.shown[:]


print("=== 2. Reset Falcon: one outcome, one box ===")
check("a reset whose connect fails, or whose byte is not written, shows only 'Reset Failed', logs it, "
      "and leaves the homed flag as it was", a_failed_reset_is_one_box_and_keeps_the_homed_flag)
check("a reset that went out shows only 'Falcon Reset' and clears the homed flag",
      a_reset_that_went_out_is_one_box_and_clears_the_flag)

# ---------------------------------------------------------------------------
# 3. The calibration jog closes its port whatever happens
# ---------------------------------------------------------------------------


def the_calibration_jog_disconnects_in_a_finally():
    fake = types.SimpleNamespace(_falcon_port="COM-T", _port_lock=threading.RLock(),
                                 _engrave_jog_step=types.SimpleNamespace(get=lambda: "10"))
    with Catch("ssc.machine") as log:
        with_fake_sender(lambda: Dialog._engrave_jog(fake, 1, 0), jog_raises=OSError("write failed"))
    s = FakeSender.made[0]
    assert s.jogs and s.disconnects == 1, f"jog raised and the port stayed open (disconnects {s.disconnects})"
    assert any("jog failed" in m for m in log.messages(logging.WARNING))
    with_fake_sender(lambda: Dialog._engrave_jog(fake, 0, -1))
    s = FakeSender.made[0]
    assert s.jogs == [dict(x=0.0, y=-10.0, feed=3000, relative=True)] and s.disconnects == 1, (s.jogs, s.disconnects)


print("=== 3. Calibration jog ===")
check("a jog that raises still closes the port; a jog that works sends once and closes once",
      the_calibration_jog_disconnects_in_a_finally)

# ---------------------------------------------------------------------------
# 4. The run dialog follows the sender's answer; the sender gives one
# ---------------------------------------------------------------------------
Run = ui_dialogs.FalconRunDialog


class Btn:
    def __init__(self):
        self.calls = []

    def config(self, **kw):
        self.calls.append(kw)


def run_fake(sender):
    return types.SimpleNamespace(
        _jog_step_var=types.SimpleNamespace(get=lambda: "1"), _sender=sender, _has_jogged=False,
        _finished=True, _stop_btn=Btn(), _state_var=types.SimpleNamespace(set=lambda v: None),
        _done_button_label="Start Frame", _pause_btn=Btn(), _resume_btn=Btn())


class Answering:
    """A sender whose jog / pause / resume answer as scripted (or raise)."""

    def __init__(self, answer):
        self.answer = answer

    def _answer(self):
        if isinstance(self.answer, Exception):
            raise self.answer
        return self.answer

    def jog(self, **kw):
        return self._answer()

    def pause(self):
        return self._answer()

    def resume(self):
        return self._answer()


def a_jog_that_did_not_go_out_does_not_unlock_start_frame():
    for answer in (False, OSError("write failed")):
        fake = run_fake(Answering(answer))
        Run._on_jog_clicked(fake, 1, 0)
        assert fake._has_jogged is False and fake._stop_btn.calls == [], (answer, fake._stop_btn.calls)
    for answer in (True, None):               # None: a stand-in or an older sender
        fake = run_fake(Answering(answer))
        Run._on_jog_clicked(fake, 1, 0)
        assert fake._has_jogged is True and fake._stop_btn.calls == [dict(state="normal")], (answer, fake._stop_btn.calls)


def pause_and_resume_flip_only_when_the_byte_went_out():
    for answer in (False, OSError("write failed")):
        fake = run_fake(Answering(answer))
        Run._on_pause_clicked(fake)
        Run._on_resume_clicked(fake)
        assert fake._pause_btn.calls == [] and fake._resume_btn.calls == [], (answer, fake._pause_btn.calls)
    for answer in (True, None):
        fake = run_fake(Answering(answer))
        Run._on_pause_clicked(fake)
        assert fake._pause_btn.calls == [dict(state="disabled")] and fake._resume_btn.calls == [dict(state="normal")]
        Run._on_resume_clicked(fake)
        assert fake._pause_btn.calls[-1] == dict(state="normal") and fake._resume_btn.calls[-1] == dict(state="disabled")


class FakeSerial:
    def __init__(self, fail=False):
        self.is_open = True
        self.fail = fail
        self.written = []

    def write(self, data):
        if self.fail:
            raise OSError("write failed")
        self.written.append(bytes(data))


def the_sender_says_whether_a_command_went_out():
    if not falcon_sender.HAS_PYSERIAL:
        print("    (pyserial not installed: the sender-level check is skipped)")
        return
    s = falcon_sender.FalconSender(port="COM-T")
    s._serial = FakeSerial()
    assert s.jog(x=1.0, relative=True) is True and s.send_realtime(b"!") is True
    assert s.pause() is True and s.resume() is True
    assert s._serial.written[0].startswith(b"$J=") and b"!" in s._serial.written, s._serial.written
    s._serial = FakeSerial(fail=True)
    assert s.jog(x=1.0, relative=True) is False and s.send_realtime(b"!") is False
    assert s.pause() is False and s.resume() is False
    s._serial = None
    assert s.jog(x=1.0, relative=True) is False and s.send_realtime(b"!") is False


print("=== 4. The run dialog follows the sender's answer ===")
check("a jog the sender did not write (False, or raised) leaves Start Frame locked; a written one (True) or a "
      "stand-in's None unlocks it", a_jog_that_did_not_go_out_does_not_unlock_start_frame)
check("Pause and Resume flip their buttons only when the byte went out; a stand-in's None still flips",
      pause_and_resume_flip_only_when_the_byte_went_out)
check("FalconSender.jog, send_realtime, pause and resume return True when written, False when not or when "
      "there is no port", the_sender_says_whether_a_command_went_out)

# ---------------------------------------------------------------------------
# 5. The two refresh loops keep running and say so once
# ---------------------------------------------------------------------------


def capture_loop_fake(frame_read):
    scheduled = []
    fake = types.SimpleNamespace(
        winfo_exists=lambda: True, _cap=types.SimpleNamespace(read=frame_read, release=lambda: None),
        _refresh_after_id=None, _loop_errors=0, LIVE_REFRESH_MS=100,
        after=lambda ms, fn: scheduled.append(fn) or "after#1", after_cancel=lambda i: None,
        destroy=lambda: None, _render_frame_with_overlay=None, _refresh_loop=lambda: None)
    fake.scheduled = scheduled
    return fake


def the_capture_dialogs_loop_survives_a_bad_frame():
    def boom(frame):
        raise RuntimeError("shape mismatch")
    fake = capture_loop_fake(lambda: (True, object()))
    fake._render_frame_with_overlay = boom
    with Catch("ssc.machine") as log:
        for _i in range(5):
            ui_dialogs.CameraCaptureDialog._refresh_loop(fake)
        assert len(fake.scheduled) == 5, f"the loop stopped: {len(fake.scheduled)} reschedules in 5 frames"
        assert fake._loop_errors == 5
        errs = [r for r in log.records if r.levelno >= logging.ERROR]
        assert len(errs) == 1 and errs[0].exc_info and "camera capture" in errs[0].getMessage(), \
            [r.getMessage() for r in log.records]
        ui_dialogs.CameraCaptureDialog._on_close(fake)
        warns = [m for m in log.messages(logging.WARNING) if "5 frames failed" in m]
        assert len(warns) == 1, log.messages()
    # A read() that raises is a missed frame, not the end of the loop.
    def raising_read():
        raise RuntimeError("camera gone")
    fake = capture_loop_fake(raising_read)
    ui_dialogs.CameraCaptureDialog._refresh_loop(fake)
    assert len(fake.scheduled) == 1 and fake._loop_errors == 0


def live_loop_fake(frame_read):
    import camera_capture
    if not camera_capture.HAS_OPENCV:
        return None
    from PIL import Image, ImageTk
    scheduled = []
    fake = types.SimpleNamespace(
        _closed=False, winfo_exists=lambda: True, _cap=types.SimpleNamespace(read=frame_read, release=lambda: None),
        _refresh_after_id=None, _loop_errors=0, REFRESH_MS=200, PREVIEW_W=480, PREVIEW_H=360,
        _PIL_Image=Image, _PIL_ImageTk=ImageTk, _preview_label=Btn(), _latest_photo=None,
        after=lambda ms, fn: scheduled.append(fn) or "after#1", after_cancel=lambda i: None, destroy=lambda: None,
        _refresh_loop=lambda: None)
    fake.scheduled = scheduled
    return fake


def the_live_windows_loop_survives_a_bad_frame():
    fake = live_loop_fake(lambda: (True, "not a frame"))   # cv2.cvtColor raises on it
    if fake is None:
        print("    (OpenCV not installed: skipped)")
        return
    with Catch("ssc.machine") as log:
        for _i in range(4):
            ui_dialogs.LiveCameraWindow._refresh_loop(fake)
        assert len(fake.scheduled) == 4, f"the loop stopped: {len(fake.scheduled)} reschedules in 4 frames"
        assert fake._loop_errors == 4
        errs = [r for r in log.records if r.levelno >= logging.ERROR]
        assert len(errs) == 1 and "live camera" in errs[0].getMessage(), [r.getMessage() for r in log.records]
        ui_dialogs.LiveCameraWindow.close(fake)
        assert any("4 frames failed" in m for m in log.messages(logging.WARNING)), log.messages()
        assert fake._closed is True


print("=== 5. Camera refresh loops ===")
check("CameraCaptureDialog: a frame that fails to process is logged once with its traceback, counted, said at "
      "close, and the loop keeps running; a read() that raises is a missed frame",
      the_capture_dialogs_loop_survives_a_bad_frame)
check("LiveCameraWindow: the same rule", the_live_windows_loop_survives_a_bad_frame)

# ---------------------------------------------------------------------------
# 6. parse_pad_list
# ---------------------------------------------------------------------------


def inf_and_nan_are_rejected_lines():
    rejected = []
    pads = App.parse_pad_list(None, "7 x inf\n7 x 1e400\n7 x nan\ninf x 2\nnan x 2\n1e400 x 2\n8 x 2\n9 x max\n", rejected)
    assert pads == [{'size': 8.0, 'qty': 2}, {'size': 9.0, 'qty': 'max'}], pads
    assert rejected == ["7 x inf", "7 x 1e400", "7 x nan", "inf x 2", "nan x 2", "1e400 x 2"], rejected
    rejected = []
    assert App.parse_pad_list(None, "18.0 x 5\n 7.5x3 \nbad line\n", rejected) == \
        [{'size': 18.0, 'qty': 5}, {'size': 7.5, 'qty': 3}] and rejected == ["bad line"]


print("=== 6. Pad list ===")
check("inf, 1e400 and nan as a size or a quantity are rejected lines, like 'bad line'; good lines still parse",
      inf_and_nan_are_rejected_lines)

# ---------------------------------------------------------------------------
# 7. The tone library
# ---------------------------------------------------------------------------
import toner_engine  # noqa: E402


def a_corrupt_tone_library_is_kept_and_bad_bytes_do_not_crash():
    d = tempfile.mkdtemp(prefix="ssc-tone-lib-")
    empty = {toner_engine.DEFAULT_LIBRARY: {}}
    p = os.path.join(d, "garbage.json")
    open(p, "w", encoding="utf-8").write("{not json")
    with Catch("toner_engine") as log:
        assert toner_engine.load_tone_presets(p) == empty
    assert open(p + ".corrupt.bak", encoding="utf-8").read() == "{not json", "no .corrupt.bak was kept"
    assert any("could not read the tone library" in m for m in log.messages(logging.WARNING)), log.messages()
    p2 = os.path.join(d, "bytes.json")
    open(p2, "wb").write(b"\xff\xfe\x00{")
    assert toner_engine.load_tone_presets(p2) == empty          # UnicodeDecodeError used to escape
    assert os.path.exists(p2 + ".corrupt.bak")
    assert toner_engine.load_tone_presets(d) == empty           # a directory: OSError, not a crash
    good = os.path.join(d, "good.json")
    lib = {"Default": {"My alto": {"horn_type": "Alto", "sessions": []}}}
    import json
    open(good, "w", encoding="utf-8").write(json.dumps(lib))
    assert toner_engine.load_tone_presets(good) == lib and not os.path.exists(good + ".corrupt.bak")


print("=== 7. Tone library ===")
check("a corrupt library loads as empty AND is kept as .corrupt.bak; bad bytes and an unreadable path are a "
      "failed load, not a crash; a good library loads unchanged", a_corrupt_tone_library_is_kept_and_bad_bytes_do_not_crash)

# ---------------------------------------------------------------------------
# 8. The small ones
# ---------------------------------------------------------------------------


def a_short_engrave_offset_does_not_raise():
    import camera_capture
    if not camera_capture.HAS_OPENCV:
        print("    (OpenCV not installed: skipped)")
        return
    cal = {'homography_px_to_machine_mm': [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]],
           'camera_matrix': [[800.0, 0.0, 320.0], [0.0, 800.0, 240.0], [0.0, 0.0, 1.0]],
           'dist_coeffs': [0.0, 0.0, 0.0, 0.0, 0.0],
           'calibration_schema_version': 1, 'card_engrave_offset_mm': [5.0]}
    camera_capture._schema1_fix_failures['count'] = 0
    with Catch("camera_capture") as log:
        out = camera_capture.pixels_to_mm([(0, 0), (10, 0), (10, 10)], cal)
        out2 = camera_capture.pixels_to_mm([(0, 0), (10, 0), (10, 10)], cal)
    assert len(out) == 3 and len(out2) == 3
    assert len([m for m in log.messages(logging.WARNING) if "schema 1" in m]) == 1, log.messages()


def windows_refusing_to_stay_awake_is_said_once():
    import platform
    import sleep_lock
    if platform.system() != "Windows":
        print("    (not Windows: skipped)")
        return
    import ctypes
    k32 = ctypes.windll.kernel32
    real = k32.SetThreadExecutionState
    k32.SetThreadExecutionState = lambda flags: 0
    sleep_lock._state.pop('zero_warned', None)
    try:
        with Catch("sleep_lock") as log:
            sleep_lock.prevent_sleep()
            sleep_lock.prevent_sleep()
        zero = [m for m in log.messages(logging.WARNING) if "returned 0" in m]
        assert len(zero) == 1, log.messages()
    finally:
        k32.SetThreadExecutionState = real
        sleep_lock._state.pop('zero_warned', None)
        sleep_lock.allow_sleep()


print("=== 8. The small ones ===")
check("pixels_to_mm on a schema-1 calibration with a one-value engrave offset returns points and warns once",
      a_short_engrave_offset_does_not_raise)
check("sleep_lock: SetThreadExecutionState returning 0 is warned about once per process",
      windows_refusing_to_stay_awake_is_said_once)

# ---------------------------------------------------------------------------
# 9. The runner and the tuner suite's bound
# ---------------------------------------------------------------------------
import run_tests as rt  # noqa: E402


def a_hung_suite_is_killed_with_its_children_and_its_output_kept():
    d = tempfile.mkdtemp(prefix="ssc-hang-")
    script = os.path.join(d, "hangs.py")
    open(script, "w", encoding="utf-8").write(
        "import subprocess, sys, time\n"
        "print('partial line from the suite', flush=True)\n"
        "subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(60)'])\n"   # inherits stdout
        "time.sleep(60)\n")
    t0 = time.time()
    code, out = rt.run_suite([sys.executable, script], d, dict(os.environ), timeout_s=3)
    took = time.time() - t0
    assert code == 124, code
    assert took < 40, f"the runner was held {took:.0f} s past a 3 s timeout (a grandchild holding the pipe?)"
    assert "TIMEOUT after 3s" in out and "partial line from the suite" in out, out
    # The suite's own faulthandler dump, armed by the runner's sitecustomize,
    # names where it sat: the script's frame, in every thread's stack.
    assert "hangs.py" in out and ("Thread 0x" in out or "Current thread" in out), \
        "no faulthandler dump in the hung suite's output:\n" + out[-1200:]
    assert rt.TIMEOUT_S <= 300 or os.environ.get("SSC_SUITE_TIMEOUT_S"), rt.TIMEOUT_S
    assert rt.dump_at(300) == 285 and 0 < rt.dump_at(3) < 3


def the_tuner_suite_bounds_every_audio_call():
    src = open(os.path.join(ROOT, "tools", "test_tuner_engine.py"), encoding="utf-8").read()
    assert "AUDIO_BOUND_S" in src and "def bounded(" in src
    for line in src.splitlines():
        s = line.strip()
        if ("player.play(" in s or "player.stop" in s or "query_devices(" in s) and not s.startswith("#"):
            assert "bounded(" in s or s.startswith("_sd.query_devices") or s.startswith("def "), \
                f"an unbounded audio call: {s}"
    assert "os._exit(" in src, "a hung audio thread would keep the process alive at exit"


def the_toner_tabs_mic_check_is_bounded():
    """toner_tab._toner_start calls check_mic_quality() on the tab's first start, which asks PortAudio
    for the default input device: the one device query the Toner tab, its suite and the CI tour make
    with no bound (found 2026-11-10 while looking for what else can hang like the tuner suite did)."""
    if not toner_engine.AUDIO_AVAILABLE:
        print("    (no audio stack here: skipped)")
        return
    real_sd = toner_engine.sd

    class Hanging:
        default = types.SimpleNamespace(device=[0, 0])

        @staticmethod
        def query_devices(index=None):
            time.sleep(12)
            return {'name': 'never seen'}

    builtin_name = f"{toner_engine._BUILTIN_MIC_KEYWORDS[0].title()} Microphone"   # whatever the engine calls built-in

    class Answering:
        default = types.SimpleNamespace(device=[0, 0])

        @staticmethod
        def query_devices(index=None):
            return {'name': builtin_name}

    toner_engine.sd = Hanging
    try:
        with Catch("toner_engine") as log:
            t0 = time.time()
            r = toner_engine.check_mic_quality()
            took = time.time() - t0
        assert r == (False, "unknown"), r
        assert took < 10, f"check_mic_quality waited {took:.0f} s on a device query that never answers"
        assert any("did not answer" in m for m in log.messages(logging.WARNING)), log.messages()
        toner_engine.sd = Answering
        assert toner_engine.check_mic_quality() == (True, builtin_name), toner_engine.check_mic_quality()
    finally:
        toner_engine.sd = real_sd


print("=== 9. The runner and the tuner suite ===")
check("toner_engine.check_mic_quality (the Toner tab's first-start device query) answers within its bound when "
      "PortAudio never does, and still recognises a built-in mic", the_toner_tabs_mic_check_is_bounded)
check("run_tests.run_suite kills a hung suite and the grandchild holding its stdout at the timeout, and returns "
      "the output so far; the per-suite timeout is 300 s", a_hung_suite_is_killed_with_its_children_and_its_output_kept)
check("test_tuner_engine runs every audio-device call under a time bound and exits without waiting for a hung one",
      the_tuner_suite_bounds_every_audio_call)

# ---------------------------------------------------------------------------
# 10. A non-numeric stored reference pitch (GUI)
# ---------------------------------------------------------------------------


def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI check skipped: no display: {e})")
        return
    root.withdraw()
    import config
    s = config.load_settings()
    s.setdefault("tuner_settings", {})["reference_pitch"] = "abc"
    s.setdefault("toner_settings", {})["reference_pitch"] = "abc"
    config.save_settings(s)

    def the_tabs_build_with_a_bad_stored_pitch():
        app = App(root)               # used to fail while building the Tuner tab
        tcl_error = tk.TclError('expected floating-point number but got "abc"')

        def raising_get():
            raise tcl_error
        if hasattr(app, "_tuner_pitch_var"):
            assert app._tuner_pitch_var.get() == 440.0, app._tuner_pitch_var.get()
            # A Scale-bound DoubleVar refuses a non-number at set(); the way a
            # bad value reaches get() is a stored one, so stand in for get().
            app._tuner_pitch_var.get = raising_get
            app._tuner_on_pitch_changed()      # must swallow the TclError
        else:
            print("    (no Tuner tab in this build: the tuner checks are skipped)")
        if hasattr(app, "_toner_pitch_var"):
            assert app._toner_pitch_var.get() == 440.0
            app._toner_pitch_var.get = raising_get
            app._toner_on_pitch_changed()
    check("a non-numeric reference pitch in the settings file: the tabs build at 440 and the pitch handlers "
          "swallow the TclError", the_tabs_build_with_a_bad_stored_pitch)
    root.destroy()


print("=== 10. Stored pitch ===")
gui()

print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
