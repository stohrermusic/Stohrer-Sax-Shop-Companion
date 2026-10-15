"""Camera Calibration dialog driven end to end with a fake camera.

Gates for the v2.86 calibration fixes and diagnostics:

  - A board in view gets its green corners, the "Detected N corners"
    status and an enabled Capture button. v2.80 indexed each corner as
    pt[0][0], which raises IndexError on OpenCV 5's (N, 2) corner array;
    the refresh loop swallowed it, so the preview froze on the last frame
    the board was NOT found in and Capture never enabled. A user on a
    Falcon hit exactly that (and a black preview after Reconnect, the
    same bug with the preview blanked first).
  - Reconnect brings the live preview back.
  - Engrave phase -> Next opens the camera once (it opened it twice:
    two handles on one device, two refresh loops).
  - Machine INFO lines reach app.log (the root logger is WARNING, and
    every falcon_sender INFO line was being dropped).
  - The preview logs a throttled summary, logs a frame error once with
    its traceback, and saves the raw frame to the config folder.
  - corner_points accepts both OpenCV corner shapes; the file-manager
    command per platform; falcon_sender's repeat throttle and its
    home/jog/unlock lines.

The camera is a fake that serves a rendered ChArUco card; only message
boxes are stubbed. Isolated profile (SAXSHOP_CONFIG_DIR). Self-skips
without a display or without OpenCV.
"""
import logging
import os
import sys
import tempfile
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-calib-dialog-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import config  # noqa: E402

LOG_PATH = config.setup_logging()

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


def expect(actual, expected, shown=None):
    """Assert, showing the value on failure. (A check written as
    `lambda: a == b or a` passes whenever `a` is truthy -- three checks in
    this file did exactly that and passed against the old code.)"""
    assert actual == expected, f"got {shown if shown is not None else actual!r}"


def log_text():
    for h in logging.getLogger().handlers:
        h.flush()
    with open(LOG_PATH, encoding="utf-8") as f:
        return f.read()


class FakeCap:
    """Stands in for cv2.VideoCapture: every read() serves the same frame."""
    opened = []

    def __init__(self, frame):
        self.frame = frame
        self.released = False
        self.reads = 0
        FakeCap.opened.append(self)

    def read(self):
        self.reads += 1
        return True, self.frame.copy()

    def release(self):
        self.released = True

    def isOpened(self):
        return not self.released


def board_frame(cam):
    import cv2
    import numpy as np
    board = cam.make_charuco_board()
    card = board.generateImage((420, 420), marginSize=20)
    img = np.full((480, 640), 60, np.uint8)
    img[30:450, 110:530] = card
    return cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)


def pump(root, ms):
    """Run the Tk main loop for `ms` milliseconds (lets worker threads'
    after() callbacks land)."""
    root.after(ms, root.quit)
    root.mainloop()


def run_dialog(root, ui, steps, cls=None, args=None, **kwargs):
    """Construct a modal dialog with `steps` scheduled on the root.

    Each step is (delay_ms, fn(dlg)); they run in order, each `delay_ms`
    after the previous one, and the dialog is closed after the last
    (unless a step already closed it). Returns the dialog once closed.
    Defaults to CameraCalibrationDialog(root, 0, <profile>/calib.json)."""
    found = {}
    cls = cls or ui.CameraCalibrationDialog
    args = args or (0, os.path.join(_PROFILE, "calib.json"))

    def find():
        for w in root.winfo_children():
            if isinstance(w, cls):
                return w
        return None

    def run(i):
        dlg = found.get('dlg') or find()
        if dlg is None:
            root.after(20, run, i)
            return
        found['dlg'] = dlg
        if i < len(steps):
            delay, fn = steps[i]

            def go():
                try:
                    fn(dlg)
                finally:
                    run(i + 1)
            root.after(delay, go)
        elif dlg.winfo_exists():
            dlg._on_close()

    def construct():
        # Inside root.mainloop(), as in the app: the camera-open worker
        # thread calls after(), which Tk only accepts while the main
        # thread is in mainloop (the dialog's own wait_window isn't one).
        try:
            cls(root, *args, **kwargs)
        finally:
            root.quit()

    root.after(0, construct)
    root.after(50, run, 0)
    root.mainloop()
    return found['dlg']


class FakeSender:
    """falcon_sender.FalconSender stand-in: a Falcon that is always Idle
    at machine (200, 200). `connect_delay` mimics the real connect()'s
    fixed sleeps."""
    connect_delay = 0.0

    def __init__(self, port=None, **kwargs):
        self.port = port

    def connect(self):
        if FakeSender.connect_delay:
            time.sleep(FakeSender.connect_delay)

    def disconnect(self):
        pass

    def unlock(self):
        pass

    def get_status(self, timeout=0.3):
        return {'state': 'Idle', 'mpos': [200.0, 200.0, 0.0]}

    def home(self, timeout_s=60.0):
        return True, 'ok'

    def jog(self, *args, **kwargs):
        pass


def make_fake_run(tk):
    class FakeRun(tk.Toplevel):
        """FalconRunDialog stand-in: takes the input grab like the real
        one, records the G-code it was handed, ends with `reason`."""
        reason = "complete"
        calls = []

        def __init__(self, parent, sender, gcode_lines, **kwargs):
            super().__init__(parent)
            self.geometry("1x1+0+0")
            self.update()
            try:
                self.grab_set()
            except tk.TclError:
                pass
            FakeRun.calls.append(list(gcode_lines))
            self._final_reason = FakeRun.reason
            self.destroy()
    return FakeRun


def run_fix_cases(root, ui, cam, frame, seen):
    """The v2.86 fixes from the calibration bug hunt, each through the
    real dialog."""
    import json
    import tkinter as tk
    import falcon_sender
    print()
    print("v2.86 calibration fixes")
    print("=" * 60)

    good_open = cam.open_camera
    real_sender = falcon_sender.FalconSender
    real_run = ui.FalconRunDialog
    real_calibrate = cam.calibrate_from_frames
    real_save = cam.save_calibration
    real_enum = cam.enumerate_cameras
    real_hold = ui.CameraCalibrationDialog.STATUS_HOLD_S
    FakeRun = make_fake_run(tk)
    falcon_sender.FalconSender = FakeSender
    ui.FalconRunDialog = FakeRun
    warnings = []
    ui.messagebox.showwarning = lambda title, msg, **k: warnings.append(str(title))
    ui.messagebox.askyesno = lambda *a, **k: seen.get('yes', True)
    ui.messagebox.askyesnocancel = lambda *a, **k: False  # "already homed"
    engrave_kw = {"falcon_port": "COM-FAKE", "settings": {}}

    def fail_open(*a, **k):
        raise RuntimeError("no camera (test)")

    try:
        # ---- #2 the controls survive a camera that won't open ------------
        def layout(tag):
            def look(d):
                d.update_idletasks()
                panel = d._btn_frame.master
                seen[tag] = (panel.winfo_width(),
                             d._cancel_btn.winfo_viewable(),
                             d._switch_cam_btn.winfo_viewable())
            return look

        cam.open_camera = fail_open
        run_dialog(root, ui, [(900, layout('fail_capture'))])
        run_dialog(root, ui, [(900, layout('fail_engrave'))], **engrave_kw)
        cam.open_camera = good_open
        run_dialog(root, ui, [(900, layout('ok_engrave'))], **engrave_kw)

        def controls_survive():
            for tag in ('fail_capture', 'fail_engrave', 'ok_engrave'):
                width, cancel, switch = seen[tag]
                assert width >= 300, f"{tag}: right panel {width} px wide"
                assert cancel and switch, f"{tag}: Cancel/Switch not on screen"
        check("camera fails to open: the control panel and Cancel stay on "
              "screen; engrave phase fits its controls", controls_survive)

        # ---- #3 status messages are not overwritten by the preview -------
        ui.CameraCalibrationDialog.STATUS_HOLD_S = 1.0

        def put(text):
            return lambda d: d._status_var.set(text)

        def read(tag):
            return lambda d: seen.__setitem__(tag, d._status_var.get())

        run_dialog(root, ui, [(900, put("probe-engrave-msg")),
                              (1500, read('st_engrave'))], **engrave_kw)
        run_dialog(root, ui, [(900, put("probe-capture-msg")),
                              (500, read('st_held')),
                              (1300, read('st_after'))])

        def status_held():
            assert seen['st_engrave'] == "probe-engrave-msg", seen['st_engrave']
            assert seen['st_held'] == "probe-capture-msg", seen['st_held']
            assert seen['st_after'].startswith("Detected "), seen['st_after']
        check("status messages stay up; capture progress resumes after the hold",
              status_held)
        ui.CameraCalibrationDialog.STATUS_HOLD_S = real_hold

        # ---- #4a Retake after Done goes back to taking references --------
        def retake_case(d):
            d._on_capture_clicked()
            d._on_done_refs_clicked()
            assert len(d._captures) == 1 and d._capture_substate == 'intrinsics', (
                "setup: no reference capture taken")
            d._on_retake_last_clicked()
            seen['retake'] = (d._capture_substate, d._reference_count,
                              d._capture_btn.cget('text'),
                              str(d._done_refs_btn.cget('state')))
        run_dialog(root, ui, [(1000, retake_case)])
        check("Retake past the last reference returns to Capture reference",
              lambda: expect(seen['retake'], ('reference', 0, "Capture reference",
                                              "disabled")))

        # ---- #4b-c Calibrate refuses capture sets that can't be right -----
        calls = []
        cam.calibrate_from_frames = lambda caps, size, board, **kw: (
            calls.append((len(caps), kw)) or {
                'rms_reprojection_error_px': 0.5, 'frame_count': len(caps),
                'camera_matrix': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
                'dist_coeffs': [[0, 0, 0, 0, 0]]})
        cam.save_calibration = lambda calib, path: None

        def no_done(d):
            for _i in range(6):
                d._on_capture_clicked()
            d._captured_engrave_offset = (50.0, 60.0)
            d._on_calibrate_clicked()

        def moved_set(offset):
            def go(d):
                d._on_capture_clicked()
                d._on_done_refs_clicked()
                for _i in range(5):
                    d._on_capture_clicked()
                d._captured_engrave_offset = offset
                d._on_calibrate_clicked()
            return go

        warnings.clear()
        run_dialog(root, ui, [(1000, no_done)])
        seen['calls_no_done'] = len(calls)
        seen['warn_no_done'] = list(warnings)
        seen['yes'] = False  # decline to enter an offset by hand
        run_dialog(root, ui, [(1000, moved_set(None))])
        seen['calls_no_offset'] = len(calls)
        seen['yes'] = True
        run_dialog(root, ui, [(1000, moved_set((50.0, 60.0)))])
        seen['calls_ok'] = len(calls)

        def guards():
            assert seen['calls_no_done'] == 0 and "Move the Card" in seen['warn_no_done'], \
                (seen['calls_no_done'], seen['warn_no_done'])
            assert seen['calls_no_offset'] == 0, "calibrated with no engrave offset"
            assert seen['calls_ok'] == 1, seen['calls_ok']
            n, kw = calls[0]
            assert n == 6 and kw['reference_indices'] == (0,), (n, kw)
            assert kw['card_offset_x_mm'] == 50.0, kw
        check("Calibrate refuses: no 'Done with references', no engrave "
              "offset; a proper set calibrates", guards)
        cam.calibrate_from_frames = real_calibrate
        cam.save_calibration = real_save

        # ---- #5/#7/#11 framing: abort clears trust; G92.1; grab returns ---
        def frame_case(tag):
            def go(d):
                d._on_frame_clicked()
                seen[tag] = (d._framing_done, d._engrave_homed,
                             str(d._engrave_btn.cget('state')),
                             str(d.grab_current()) == str(d))
            return go

        FakeRun.calls.clear()
        FakeRun.reason = None  # the window's X: soft reset mid-trace
        run_dialog(root, ui, [(800, frame_case('frame_abort'))], **engrave_kw)
        FakeRun.reason = "complete"
        run_dialog(root, ui, [(800, frame_case('frame_ok'))], **engrave_kw)

        def framing():
            done, homed, state, _grab = seen['frame_abort']
            assert (done, homed, state) == (False, False, "disabled"), seen['frame_abort']
            done, homed, state, grab = seen['frame_ok']
            assert done and homed and state == "normal", seen['frame_ok']
            assert grab, "calibration dialog lost its grab after framing"
            assert FakeRun.calls and FakeRun.calls[-1][0] == "G92.1", \
                FakeRun.calls[-1][:3] if FakeRun.calls else "no framing G-code"
        check("framing: aborted run drops home+frame; G92.1 first; dialog "
              "modal again", framing)

        # ---- engrave: offset saved, G92.1 before motion, grab, then
        # #6 Engrave new card -> Frame enabled again ------------------------
        settings = {}

        def engrave_then_new(d):
            d._engrave_homed = True
            d._framing_done = True
            d._on_engrave_clicked()
            seen['eng_grab'] = str(d.grab_current()) == str(d)
            d._on_reengrave_clicked()
            seen['frame_state'] = str(d._frame_btn.cget('state'))

        FakeRun.calls.clear()
        run_dialog(root, ui, [(800, engrave_then_new)],
                   falcon_port="COM-FAKE", settings=settings)

        def engrave_case():
            lines = FakeRun.calls[-1]
            # A move has coordinates; the header's "G00 G17 G40 G21 G54"
            # only sets modes.
            first_motion = next(i for i, ln in enumerate(lines)
                                if ln.startswith(('G0', 'G1'))
                                and ('X' in ln or 'Y' in ln))
            assert "G92.1" in lines[:first_motion], "no G92.1 before motion"
            assert seen['eng_grab'], "calibration dialog lost its grab after engrave"
            assert seen['frame_state'] == "normal", seen['frame_state']
        check("engrave: G92.1 before motion, modal again; Engrave new card "
              "re-enables Frame", engrave_case)

        # ---- #9 the head-position poll doesn't block the UI ---------------
        FakeSender.connect_delay = 0.4

        def poll(d):
            t0 = time.perf_counter()
            d._poll_mpos()
            seen['poll_ms'] = (time.perf_counter() - t0) * 1000

        run_dialog(root, ui, [(500, poll), (1500, read('mpos_text'))],
                   **engrave_kw)
        FakeSender.connect_delay = 0.0

        def poll_case():
            assert seen['poll_ms'] < 100, f"poll blocked the UI {seen['poll_ms']:.0f} ms"
        check("head-position poll runs off the UI thread", poll_case)

        def mpos_shown(d):
            seen['mpos_label'] = d._mpos_var.get()
        run_dialog(root, ui, [(1500, mpos_shown)], **engrave_kw)
        check("...and still shows the position",
              lambda: expect("200.00" in seen['mpos_label'], True, seen['mpos_label']))

        # ---- #8 closed while the camera is still opening -----------------
        def slow_open(index, *a, **k):
            time.sleep(0.6)
            return FakeCap(frame)

        FakeCap.opened.clear()
        cam.open_camera = slow_open
        run_dialog(root, ui, [(100, lambda d: None)])
        pump(root, 1200)
        check("closed while the camera opens: the handle is released",
              lambda: expect(bool(FakeCap.opened) and FakeCap.opened[-1].released, True))
        cam.open_camera = good_open

        # ---- #13 Switch camera discards captures from the old camera -----
        cam.enumerate_cameras = lambda *a, **k: [
            {'index': 0, 'width': 640, 'height': 480},
            {'index': 1, 'width': 640, 'height': 480}]

        def switch(d):
            d._on_capture_clicked()
            assert len(d._captures) == 1, "setup: no capture taken"
            d._on_switch_camera_clicked()
            seen['after_switch'] = (len(d._captures), d._reference_count)
        run_dialog(root, ui, [(1000, switch)])
        cam.enumerate_cameras = real_enum
        check("Switch camera with captures: discarded, not mixed",
              lambda: expect(seen['after_switch'], (0, 0)))

        # ---- scrap capture dialog: size mismatch and a dead camera -------
        cal_path = os.path.join(_PROFILE, "capture_calib.json")
        with open(cal_path, "w", encoding="utf-8") as f:
            json.dump({'camera_matrix': [[500, 0, 400], [0, 500, 300], [0, 0, 1]],
                       'dist_coeffs': [[0, 0, 0, 0, 0]],
                       'homography_px_to_machine_mm': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
                       'image_size': [800, 600],
                       'calibration_schema_version': 2}, f)

        def capture_look(tag):
            def look(d):
                d.update_idletasks()
                # Not winfo_viewable: the dialog is transient to the
                # (withdrawn) test root, so Tk never maps it. What broke
                # was the requested size: a text-holding 640x480 label
                # asked for ~7000x10800 px, so in the fixed 690x720
                # window everything under the preview was cut off.
                fits = (d.winfo_reqheight() <= d.PREVIEW_H + 240
                        and d.winfo_reqwidth() <= d.PREVIEW_W + 50)
                seen[tag] = (str(d._use_btn.cget('state')),
                             d._status_var.get(), fits,
                             (d.winfo_reqwidth(), d.winfo_reqheight()))
            return look

        run_dialog(root, ui, [(1200, capture_look('cap_mismatch'))],
                   cls=ui.CameraCaptureDialog, args=(0, cal_path), settings={})
        cam.open_camera = fail_open
        run_dialog(root, ui, [(900, capture_look('cap_fail'))],
                   cls=ui.CameraCaptureDialog, args=(0, cal_path), settings={})
        cam.open_camera = good_open

        def capture_dialog():
            state, status, _fits, _req = seen['cap_mismatch']
            assert state == "disabled" and "800×600" in status, seen['cap_mismatch']
            assert seen['cap_fail'][2], \
                f"dialog asks for {seen['cap_fail'][3]}: Use/Cancel cut off"
        check("scrap capture: wrong-size camera refused; buttons survive a "
              "dead camera", capture_dialog)
    finally:
        cam.open_camera = good_open
        cam.calibrate_from_frames = real_calibrate
        cam.save_calibration = real_save
        cam.enumerate_cameras = real_enum
        falcon_sender.FalconSender = real_sender
        ui.FalconRunDialog = real_run
        ui.CameraCalibrationDialog.STATUS_HOLD_S = real_hold


def main():
    import tkinter as tk
    try:
        import camera_capture as cam
    except Exception as e:
        print(f"Skipping: camera_capture import failed ({e})")
        return 0
    if not cam.HAS_OPENCV:
        print("Skipping: OpenCV not available")
        return 0
    try:
        from PIL import Image, ImageTk  # noqa: F401
    except ImportError:
        print("Skipping: Pillow not available")
        return 0

    print("Engine-level")
    print("=" * 60)
    import numpy as np

    def corner_points_both_shapes():
        a = np.arange(10, dtype=np.float32).reshape(5, 2)
        assert cam.corner_points(a).shape == (5, 2)
        assert cam.corner_points(a.reshape(5, 1, 2)).tolist() == a.tolist()
    check("corner_points takes OpenCV 4 (N,1,2) and 5 (N,2) corners",
          corner_points_both_shapes)

    def detect_reports_markers():
        frame = board_frame(cam)
        st = {}
        c, _ids = cam.detect_charuco(frame, cam.make_charuco_board(), stats=st)
        assert c is not None and st['markers'] >= 40 and st['corners'] >= 60, st
        st = {}
        c, _ids = cam.detect_charuco(np.full((480, 640), 128, np.uint8),
                                     cam.make_charuco_board(), stats=st)
        assert c is None and st == {'markers': 0, 'corners': 0}, st
    check("detect_charuco reports marker and corner counts", detect_reports_markers)

    def reveal_commands():
        import main as main_mod
        cmd = main_mod.reveal_in_file_manager_command(
            r"C:\Users\Some One\AppData\Roaming\X\app.log", platform="win32")
        assert cmd == r'explorer /select,"C:\Users\Some One\AppData\Roaming\X\app.log"', cmd
        assert main_mod.reveal_in_file_manager_command(
            "/Users/a/Library/x/app.log", platform="darwin") == [
                "open", "-R", "/Users/a/Library/x/app.log"]
        assert main_mod.reveal_in_file_manager_command(
            "/home/a/.config/x/app.log", platform="linux") == [
                "xdg-open", "/home/a/.config/x"]
    check("Open Log File selects the file in Explorer/Finder", reveal_commands)

    print()
    print("Logging")
    print("=" * 60)

    def machine_info_reaches_log():
        logging.getLogger("falcon_sender").info("probe-falcon-info-7731")
        logging.getLogger("camera_capture").info("probe-camera-info-7731")
        logging.getLogger("ssc.machine").info("probe-machine-info-7731")
        logging.getLogger("some_other_module").info("probe-other-info-7731")
        text = log_text()
        for k in ("probe-falcon-info-7731", "probe-camera-info-7731",
                  "probe-machine-info-7731"):
            assert k in text, f"{k} missing from app.log"
        assert "probe-other-info-7731" not in text, \
            "INFO from the rest of the app should still be filtered"
    check("machine INFO reaches app.log; the rest of the app stays at WARNING",
          machine_info_reaches_log)

    def versions_logged():
        config.log_runtime_versions()
        line = next((ln for ln in log_text().splitlines()
                     if "runtime: python" in ln), None)
        assert line, "no runtime line in app.log"
        for k in ("python", "numpy", "OpenCV", "Pillow", "Tk"):
            assert f" {k} " in line or f": {k} " in line, f"{k} missing: {line}"
        import cv2
        assert f"OpenCV {cv2.__version__}" in line, line
        assert f"OpenCV {cv2.__version__} loaded" in log_text(), \
            "camera_capture didn't log the OpenCV it loaded"
    check("startup logs Python/Tk/library versions; OpenCV logged on load",
          versions_logged)

    try:
        import falcon_sender as fs
    except Exception as e:
        fs = None
        print(f"  (falcon_sender checks skipped: {e})")
    if fs is not None and fs.HAS_PYSERIAL:
        def throttle():
            fs._log_last.clear()
            old = fs.LOG_REPEAT_S
            fs.LOG_REPEAT_S = 0.3
            try:
                for _i in range(5):
                    fs._log_throttled(logging.WARNING, ('t', 1),
                                      "throttle-probe-5519 %s", "x")
                time.sleep(0.35)
                fs._log_throttled(logging.WARNING, ('t', 1),
                                  "throttle-probe-5519 %s", "y")
            finally:
                fs.LOG_REPEAT_S = old
            lines = [ln for ln in log_text().splitlines()
                     if "throttle-probe-5519" in ln]
            assert len(lines) == 2, lines
            assert "4 repeats" in lines[1], lines[1]
        check("a repeating connect message is written once per window, "
              "with the count", throttle)

        class FakeSerial:
            is_open = True
            in_waiting = 0

            def __init__(self):
                self.written = []
                self.pending = b''

            def write(self, b):
                self.written.append(b)
                if b == b'$H\n':
                    self.pending = b'ok\n'
                return len(b)

            def read(self, n=1):
                out, self.pending = self.pending, b''
                return out

            def reset_input_buffer(self):
                pass

            def close(self):
                self.is_open = False

        def sender_lines():
            s = fs.FalconSender(port="COM-TEST")
            s._serial = FakeSerial()
            assert s.home(timeout_s=1.0) == (True, 'ok')
            s.jog(x=10, y=-5, feed=3000, relative=True)
            s.unlock()
            text = log_text()
            assert "home ($H) on COM-TEST: ok after" in text
            assert "jog on COM-TEST: $J=G91X10Y-5F3000" in text
            assert "unlock ($X) sent on COM-TEST" in text
        check("home, jog and unlock are logged with the port", sender_lines)

    print()
    print("Calibration dialog with a fake camera")
    print("=" * 60)
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping the dialog cases: no display available ({e})")
        return 0 if all(results) else 1
    root.withdraw()

    import ui_dialogs as ui
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(ui.messagebox, name, lambda *a, **k: None)
    ui.messagebox.askyesno = lambda *a, **k: True

    frame = board_frame(cam)
    real_open = cam.open_camera
    cam.open_camera = lambda index, *a, **k: FakeCap(frame)
    ui.CameraCalibrationDialog.SUMMARY_INTERVAL_S = 0.5
    ui.CameraCalibrationDialog.SNAPSHOT_FIRST_S = 0.3
    snap_path = os.path.join(_PROFILE, ui.CameraCalibrationDialog.SNAPSHOT_NAME)
    seen = {}

    try:
        # ---- capture phase, board in view ---------------------------------
        def look(dlg):
            seen['status'] = dlg._status_var.get()
            seen['capture_state'] = str(dlg._capture_btn.cget('state'))
            seen['photo'] = dlg._latest_photo
            seen['errors'] = dlg._loop_errors

        def capture(dlg):
            dlg._on_capture_clicked()
            seen['captures'] = len(dlg._captures)
            seen['refs'] = dlg._reference_count

        def reconnect(dlg):
            seen['photo_before'] = dlg._latest_photo
            seen['cap_before'] = dlg._cap
            dlg._on_reconnect_camera_clicked()
            seen['photo_after_click'] = dlg._latest_photo

        def after_reconnect(dlg):
            seen['photo_after'] = dlg._latest_photo
            seen['status_after'] = dlg._status_var.get()
            seen['cap_after'] = dlg._cap

        FakeCap.opened.clear()
        run_dialog(root, ui, [(1200, look), (50, capture), (50, reconnect),
                              (900, after_reconnect)])

        def board_found():
            assert seen['errors'] == 0, f"{seen['errors']} frames raised"
            assert seen['status'].startswith("Detected "), seen['status']
            assert seen['capture_state'] == "normal", seen['capture_state']
            assert seen['photo'] is not None, "preview never drawn"
        check("board in view: corners found, preview drawn, Capture enabled",
              board_found)
        check("Capture reference records one reference capture",
              lambda: seen['captures'] == 1 and seen['refs'] == 1)

        def reconnect_ok():
            assert seen['cap_before'].released, "old handle not released"
            assert seen['cap_after'] is not seen['cap_before'], "no new handle"
            assert seen['photo_after'] is not seen['photo_before'], \
                "preview not redrawn after Reconnect"
            assert seen['status_after'].startswith("Detected "), seen['status_after']
        check("Reconnect: old handle released, live preview comes back", reconnect_ok)

        def preview_logged():
            text = log_text()
            assert "calibration dialog opened: phase=capture" in text
            assert "camera #0 live (capture phase): first frame 640x480x3" in text
            assert "camera #0 live (Reconnect clicked)" in text
            summaries = [ln for ln in text.splitlines()
                         if "calibration preview, last" in ln]
            assert 2 <= len(summaries) <= 6, f"{len(summaries)} summary lines"
            assert "board found in" in summaries[0] and " frames 640x480" in summaries[0]
            assert "capture 1 (reference):" in text
            assert "calibration dialog closed: result=None, 1 captures" in text
        check("preview summary is throttled; open, capture, close are logged",
              preview_logged)

        def snapshot_saved():
            import cv2
            assert os.path.exists(snap_path), "no snapshot written"
            img = cv2.imread(snap_path)
            assert img is not None and img.shape == frame.shape, \
                None if img is None else img.shape
            assert "saved the raw camera frame to" in log_text()
        check("raw camera frame saved to the config folder", snapshot_saved)

        # ---- a frame that errors is logged once, then counted -------------
        real_cp = cam.corner_points

        def boom(_c):
            raise ValueError("synthetic overlay failure 4417")
        cam.corner_points = boom
        try:
            run_dialog(root, ui, [(1300, look)])
        finally:
            cam.corner_points = real_cp

        def errors_logged():
            text = log_text()
            assert text.count("ValueError: synthetic overlay failure 4417") == 1, \
                "the traceback should be logged exactly once"
            assert "processing a frame failed" in text
            assert seen['errors'] >= 5, seen['errors']
            counted = [ln for ln in text.splitlines()
                       if "calibration preview, last" in ln
                       and "frame errors 0;" not in ln]
            assert counted, "summary never reported the frame errors"
        check("a failing frame is logged once with its traceback, then counted",
              errors_logged)

        # ---- engrave phase -> Next opens the camera once -------------------
        FakeCap.opened.clear()
        config_settings = {"camera_calibration_engrave_offset_mm": None}

        def to_capture(dlg):
            seen['opens_in_engrave'] = len(FakeCap.opened)
            dlg._captured_engrave_offset = (50.0, 60.0)
            dlg._on_next_clicked()

        def count(dlg):
            seen['opens_total'] = len(FakeCap.opened)
            seen['phase'] = dlg._phase
            reads0 = sum(c.reads for c in FakeCap.opened)
            seen['reads0'] = reads0

        def rate(dlg):
            seen['reads_1s'] = sum(c.reads for c in FakeCap.opened) - seen['reads0']

        real_poll = ui.CameraCalibrationDialog._start_mpos_polling
        ui.CameraCalibrationDialog._start_mpos_polling = lambda self: None
        try:
            run_dialog(root, ui, [(800, to_capture), (800, count), (1000, rate)],
                       falcon_port="COM-NONE", settings=config_settings)
        finally:
            ui.CameraCalibrationDialog._start_mpos_polling = real_poll

        def opened_once():
            assert seen['opens_in_engrave'] == 1, seen['opens_in_engrave']
            assert seen['phase'] == 'capture', seen['phase']
            assert seen['opens_total'] == 1, \
                f"camera opened {seen['opens_total']} times"
            # LIVE_REFRESH_MS = 100: one refresh chain reads ~10 frames/s.
            assert seen['reads_1s'] <= 14, \
                f"{seen['reads_1s']} reads in 1 s: more than one refresh loop"
        check("engrave phase -> Next: one camera handle, one refresh loop",
              opened_once)

        run_fix_cases(root, ui, cam, frame, seen)
    finally:
        cam.open_camera = real_open
        try:
            root.destroy()
        except tk.TclError:
            pass

    passed = sum(results)
    print()
    print(f"{passed}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
