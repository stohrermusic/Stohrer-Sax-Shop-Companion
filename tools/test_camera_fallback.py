"""The camera falls back when its saved number cannot be opened; what it
may fall back to, and that it never does so silently.

Matt, 2026-11-09, asked about the laser being plugged in while the app runs
and then, of the camera, "same deal?": yes. Each camera window used to
resolve an index when it opened (the saved number, camera_index_override,
winning outright) and simply fail when that number could not be opened.

Pinned here, against fakes (no camera touched; the GUI cases need OpenCV
and Pillow because the windows' refresh loops do, and self-skip without):
  - the saved number fails (or opens and gives no frame) and another camera
    opens: the window runs on it, one warning is logged with the number that
    failed, the number in use and the camera names the system reports, and
    the saved setting is untouched
  - WITH a camera calibration on disk only the Falcon-named camera may stand
    in (the calibration belongs to one camera: a laptop's webcam under the
    Falcon's homography gives a believable, wrong outline); without one,
    the last camera found as well
  - a read() that raises releases the capture before the error travels
  - the saved number opens: exactly one open call, no enumeration, no name
    lookup
  - nothing opens: the first error, as before
  - a number that is NOT the saved one fails: no fallback (that is a
    request for one specific camera)
  - all four windows that open a camera go through the one helper
    (camera_capture.open_camera_or_fallback), hand it the saved number as
    its second argument (checked in the source as code, not as text), and
    none calls open_camera itself
  - in the real windows: a fallback shows one line in the window's status
    area (camera_capture.fallback_notice); a refusal shows the old error
    and no fallback; the app's own live preview releases its camera before
    the shape dialog's "Get from camera" or Camera Calibration opens one,
    so the app never competes with itself for the Falcon camera
"""
import ast
import inspect
import json
import logging
import os
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-camera-fallback-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import camera_capture as cam  # noqa: E402

results = []
CAL_PATH = cam.default_calibration_path()
CALIBRATION = {'camera_matrix': [[500, 0, 320], [0, 500, 240], [0, 0, 1]],
               'dist_coeffs': [[0, 0, 0, 0, 0]],
               'homography_px_to_machine_mm': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
               'image_size': [640, 480],
               'calibration_schema_version': 2}


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


def calibration_on_disk(present):
    """Put a calibration at the app's path, or make sure there is none."""
    if present:
        with open(CAL_PATH, "w", encoding="utf-8") as f:
            json.dump(CALIBRATION, f)
    elif os.path.exists(CAL_PATH):
        os.remove(CAL_PATH)


class FakeCap:
    def __init__(self, world, index, frames=True, read_raises=None):
        self.world, self.index, self.frames, self.read_raises = world, index, frames, read_raises
        self.released = False

    def read(self):
        if self.read_raises is not None:
            raise self.read_raises
        if not self.frames:
            return False, None
        frame = self.world.frame
        return True, (frame.copy() if hasattr(frame, "copy") else f"frame-from-{self.index}")

    def release(self):
        if not self.released:
            self.released = True
            self.world.held.discard(self.index)


class World:
    """What the machine has: which indices open, which give frames, which
    raise on read, the Falcon-named index, the enumeration, the names.
    A device can be held by one capture at a time, as a DirectShow camera
    can: a second open while one is held fails. Records every call."""

    def __init__(self, opens=(), no_frame=(), falcon=None, cams=(), names=(), read_raises=(), frame=None):
        self.opens, self.no_frame, self.falcon = set(opens), set(no_frame), falcon
        self.cams, self.names, self.read_raises = list(cams), list(names), dict(read_raises)
        self.frame = frame
        self.calls = []
        self.caps = []
        self.held = set()

    def open_camera(self, index, *a, **k):
        self.calls.append(("open", index))
        if index not in self.opens:
            raise RuntimeError(f"Could not open camera index {index}")
        if index in self.held:
            raise RuntimeError(f"Could not open camera index {index}")      # in use by another handle
        cap = FakeCap(self, index, frames=index not in self.no_frame, read_raises=self.read_raises.get(index))
        self.held.add(index)
        self.caps.append(cap)
        return cap

    def find_falcon_camera_index(self):
        self.calls.append(("falcon",))
        return self.falcon

    def enumerate_cameras(self):
        self.calls.append(("enumerate",))
        return [{"index": i} for i in self.cams]

    def _list_camera_names(self):
        self.calls.append(("names",))
        return list(self.names)


def with_world(world, fn):
    keep = (cam.open_camera, cam.find_falcon_camera_index, cam.enumerate_cameras, cam._list_camera_names)
    cam.open_camera, cam.find_falcon_camera_index = world.open_camera, world.find_falcon_camera_index
    cam.enumerate_cameras, cam._list_camera_names = world.enumerate_cameras, world._list_camera_names
    logged = []

    class Grab(logging.Handler):
        def emit(self, record):
            logged.append((record.levelname, record.getMessage()))
    handler = Grab()
    cam.log.addHandler(handler)
    cam.log.setLevel(logging.INFO)
    try:
        return fn(logged)
    finally:
        cam.log.removeHandler(handler)
        (cam.open_camera, cam.find_falcon_camera_index, cam.enumerate_cameras, cam._list_camera_names) = keep


def warnings_in(logged):
    return [m for lvl, m in logged if lvl == "WARNING"]


def saved_number_fails_another_opens():
    calibration_on_disk(False)
    settings = {"camera_index_override": 1}
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=["Integrated Camera", "Falcon Camera"])

    def go(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, settings["camera_index_override"], settings)
        assert used == 2 and frame == "frame-from-2" and cap.index == 2, (used, frame)
        assert world.calls == [("open", 1), ("falcon",), ("enumerate",), ("open", 2), ("names",)], world.calls
        assert settings == {"camera_index_override": 1}, "the saved number must not change"
        warnings = warnings_in(logged)
        assert len(warnings) == 1, logged
        assert "camera #1" in warnings[0] and "camera #2" in warnings[0] and "Falcon Camera" in warnings[0] \
            and "saved number is unchanged" in warnings[0], warnings[0]
    with_world(world, go)
    # The Falcon-named camera comes before the last one, and the saved number is skipped on both counts.
    world = World(opens={0, 1, 3}, falcon=3, cams=[0, 1, 3], names=[])

    def go2(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 1 and world.calls == [("open", 1)]
    with_world(world, go2)
    world = World(opens={0, 3}, falcon=1, cams=[0, 3, 1], names=[])      # the saved 1 is named Falcon and is last: skipped twice

    def go3(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 3, used
        assert world.calls.count(("open", 1)) == 1, f"the saved number was tried again: {world.calls}"
        assert world.calls[-2:] == [("open", 3), ("names",)], world.calls
    with_world(world, go3)


check("no calibration: saved number fails, another camera opens: the window runs on it, one warning with both "
      "numbers and the names, the saved setting untouched", saved_number_fails_another_opens)


def with_a_calibration_only_the_falcon_named_camera_may_stand_in():
    # A calibration on disk, no camera named Falcon, a webcam at 0: today's error, nothing else tried.
    calibration_on_disk(True)
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=["Integrated Camera", "USB Camera"])

    def refused(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except RuntimeError as e:
            assert "camera index 1" in str(e), f"the first error is raised: {e}"
        else:
            raise AssertionError("fell back under a calibration with no Falcon-named camera")
        assert world.calls == [("open", 1), ("falcon",)], f"something else was tried or enumerated: {world.calls}"
        assert not warnings_in(logged), logged
        infos = [m for lvl, m in logged if lvl == "INFO" and "calibration is on disk" in m]
        assert len(infos) == 1 and "camera #1" in infos[0], f"the refusal is not in the log: {logged}"
    with_world(world, refused)
    # The same with a camera the Falcon-name heuristic identifies: that one, and only that one.
    world = World(opens={0, 2}, falcon=2, cams=[0, 2], names=["Integrated Camera", "Falcon Camera"])

    def falcon_named(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 2, used
        assert world.calls == [("open", 1), ("falcon",), ("open", 2), ("names",)], world.calls
        assert len(warnings_in(logged)) == 1
    with_world(world, falcon_named)
    # The Falcon-named one fails too: the first error, not the webcam.
    world = World(opens={0}, falcon=2, cams=[0, 2], names=[])

    def falcon_named_fails(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except RuntimeError as e:
            assert "camera index 1" in str(e), e
        else:
            raise AssertionError("fell back to the webcam after the Falcon-named camera failed")
        assert ("open", 0) not in world.calls and ("enumerate",) not in world.calls, world.calls
    with_world(world, falcon_named_fails)
    # The caller may say so itself: calibration_present=True with no file, False with one.
    calibration_on_disk(False)
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=[])

    def told_present(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {}, calibration_present=True)
        except RuntimeError:
            pass
        else:
            raise AssertionError("calibration_present=True was ignored")
        assert ("enumerate",) not in world.calls, world.calls
    with_world(world, told_present)
    calibration_on_disk(True)
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=[])

    def told_absent(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {}, calibration_present=False)
        assert used == 2, used
    with_world(world, told_absent)
    # An unreadable calibration file counts as present (the safe side).
    with open(CAL_PATH, "w", encoding="utf-8") as f:
        f.write("{not json")
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=[])

    def unreadable(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except RuntimeError:
            pass
        else:
            raise AssertionError("an unreadable calibration file opened the full fallback")
        assert ("enumerate",) not in world.calls, world.calls
    with_world(world, unreadable)
    calibration_on_disk(False)


check("with a calibration on disk only the Falcon-named camera may stand in for the saved one; none: today's "
      "error, nothing enumerated; the caller's calibration_present flag wins; an unreadable file counts as present",
      with_a_calibration_only_the_falcon_named_camera_may_stand_in)


def a_read_that_raises_releases_the_capture():
    calibration_on_disk(False)
    # Not the saved number: the error travels, the handle does not stay open.
    world = World(opens={0, 1}, read_raises={1: OSError("usb reset (test)")})

    def plain(logged):
        try:
            cam.open_camera_or_fallback(1, 0, {})
        except OSError as e:
            assert "usb reset" in str(e), e
        else:
            raise AssertionError("a read that raised was swallowed")
        assert world.caps and world.caps[0].released, "the capture whose read raised was not released"
        assert world.held == set(), f"a device is still held: {world.held}"
    with_world(world, plain)
    # The saved number: released, then the fallback runs as for any other failure.
    world = World(opens={0, 1}, cams=[0, 1], read_raises={1: OSError("usb reset (test)")})

    def then_falls_back(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 0 and world.caps[0].released and world.held == {0}, (used, world.held)
    with_world(world, then_falls_back)
    # Every candidate raises on read: every handle released, the first error raised.
    world = World(opens={0, 1}, cams=[0, 1], read_raises={1: OSError("usb reset (test)"), 0: OSError("also (test)")})

    def all_raise(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except OSError as e:
            assert "usb reset" in str(e), f"the FIRST error is raised: {e}"
        else:
            raise AssertionError("nothing opened and no error was raised")
        assert all(c.released for c in world.caps) and world.held == set(), world.held
    with_world(world, all_raise)


check("a read() that raises releases the capture before the error travels; the saved number then falls back as "
      "for any failure", a_read_that_raises_releases_the_capture)


def saved_number_opens_but_gives_no_frame():
    calibration_on_disk(False)
    world = World(opens={1, 2}, no_frame={1}, falcon=None, cams=[1, 2], names=["A", "B"])

    def go(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 2 and world.caps[0].released, "the frameless capture must be released before falling back"
        assert warnings_in(logged), logged
    with_world(world, go)


check("saved number opens but returns no frame: released, and the fallback runs", saved_number_opens_but_gives_no_frame)


def saved_number_opens_nothing_else_happens():
    for present in (False, True):
        calibration_on_disk(present)
        world = World(opens={1}, falcon=0, cams=[0, 1], names=["x"])

        def go(logged):
            cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
            assert (used, frame) == (1, "frame-from-1")
            assert world.calls == [("open", 1)], f"an extra call when the saved camera opened: {world.calls}"
            assert not warnings_in(logged), logged
        with_world(world, go)
    calibration_on_disk(False)


check("saved number opens: exactly one open call, no enumeration, no name lookup, no warning, calibration or not",
      saved_number_opens_nothing_else_happens)


def nothing_opens_is_the_old_error():
    calibration_on_disk(False)
    world = World(opens=set(), falcon=0, cams=[0, 2], names=["x"])

    def go(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except RuntimeError as e:
            assert "camera index 1" in str(e), f"the FIRST error is raised: {e}"
        else:
            raise AssertionError("nothing opened and no error was raised")
        assert ("open", 0) in world.calls and ("open", 2) in world.calls
        assert not warnings_in(logged), "nothing to say when nothing worked"
    with_world(world, go)
    # Not the saved number: that is a request for one camera; no fallback.
    world = World(opens={0, 2}, falcon=2, cams=[0, 2], names=["x"])

    def go2(logged):
        try:
            cam.open_camera_or_fallback(1, 0, {})
        except RuntimeError:
            pass
        else:
            raise AssertionError("a number that is not the saved one fell back")
        assert world.calls == [("open", 1)], world.calls
    with_world(world, go2)
    world = World(opens={0}, falcon=None, cams=[0], names=[])

    def go3(logged):
        try:
            cam.open_camera_or_fallback(1, None, {})
        except RuntimeError:
            pass
        else:
            raise AssertionError("with no saved number a failure fell back")
        assert world.calls == [("open", 1)], world.calls
    with_world(world, go3)


check("nothing opens: the first error, as before; a number that is not the saved one never falls back",
      nothing_opens_is_the_old_error)


def the_notice_is_one_sentence():
    text = cam.fallback_notice(1, 2)
    assert text == "Camera #1 could not be opened; showing camera #2.", text
    # Translated in every catalog (the sentence is user-facing).
    import i18n
    for lang, _name in i18n.available_languages():
        if lang == "en":
            continue
        init_translation(lang)
        try:
            other = cam.fallback_notice(1, 2)
        finally:
            init_translation("en")
        assert other != text and "1" in other and "2" in other, f"{lang}: {other!r}"


check("the fallback notice is one sentence with both numbers, translated in every catalog", the_notice_is_one_sentence)


def every_camera_window_goes_through_the_helper():
    import ui_dialogs
    windows = [ui_dialogs.PolygonDrawWindow, ui_dialogs.CameraCalibrationDialog,
               ui_dialogs.CameraCaptureDialog, ui_dialogs.LiveCameraWindow]
    for cls in windows:
        src = inspect.getsource(cls)
        assert ".open_camera(" not in src, f"{cls.__name__} still opens the camera itself"
        # The call itself, as code: its second argument must be the saved
        # number, (self._settings or {}).get("camera_index_override"). A
        # text search for the string could be satisfied by another use of
        # the key in the same class (PolygonDrawWindow's own resolver has
        # one) while the call passed None.
        tree = ast.parse(src)
        calls = [n for n in ast.walk(tree)
                 if isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                 and n.func.attr == "open_camera_or_fallback"]
        assert len(calls) == 1, f"{cls.__name__}: {len(calls)} helper calls"
        call = calls[0]
        assert len(call.args) >= 2, f"{cls.__name__}: the helper is called with {len(call.args)} positional arguments"
        second = call.args[1]
        ok = (isinstance(second, ast.Call) and isinstance(second.func, ast.Attribute) and second.func.attr == "get"
              and second.args and isinstance(second.args[0], ast.Constant)
              and second.args[0].value == "camera_index_override"
              and isinstance(second.func.value, ast.BoolOp)
              and any(isinstance(v, ast.Attribute) and v.attr == "_settings" for v in second.func.value.values))
        assert ok, f"{cls.__name__}: the helper's second argument is not the saved number: {ast.unparse(second)}"
        # The number the window runs on afterwards is the one the helper used
        # (the worker that makes the call; the calibration dialog has other workers).
        fn = next(n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)
                  and any(c is call for c in ast.walk(n)))
        assigned = [ast.unparse(s.targets[0]) for s in ast.walk(fn) if isinstance(s, ast.Assign)]
        assert "self._camera_index" in assigned, f"{cls.__name__}: the worker does not adopt the camera used"
        # A fallback is said: the worker builds the notice from the helper's answer.
        assert "fallback_notice(" in ast.unparse(fn), f"{cls.__name__}: the worker never shows the fallback notice"
    # ui_dialogs is a package (2026-11-10): its source is every submodule's, not __init__'s.
    whole = "".join(inspect.getsource(sys.modules[n]) for n in sorted(sys.modules)
                    if n.startswith("ui_dialogs.") and sys.modules[n] is not None)
    assert whole.count("open_camera_or_fallback(") == len(windows), \
        f"{whole.count('open_camera_or_fallback(')} call sites for {len(windows)} windows"
    assert ".open_camera(" not in whole.replace("open_camera_or_fallback(", ""), "a camera is opened outside the helper"
    # The two windows that map the picture through a calibration say so to the helper.
    for cls in (ui_dialogs.PolygonDrawWindow, ui_dialogs.CameraCaptureDialog):
        src = inspect.getsource(cls)
        call = next(n for n in ast.walk(ast.parse(src)) if isinstance(n, ast.Call)
                    and isinstance(n.func, ast.Attribute) and n.func.attr == "open_camera_or_fallback")
        kw = {k.arg: k.value for k in call.keywords}
        assert "calibration_present" in kw and isinstance(kw["calibration_present"], ast.Constant) \
            and kw["calibration_present"].value is True, f"{cls.__name__} does not tell the helper it holds a calibration"


check("all four camera windows open through the one helper with the saved number as its second argument (AST), "
      "adopt the number used and build the notice; the calibrated views say they hold a calibration",
      every_camera_window_goes_through_the_helper)


# ---------------------------------------------------------------- the real windows
def gui():
    import tkinter as tk
    if not cam.HAS_OPENCV:
        print("  (GUI checks skipped: OpenCV not installed; the windows' refresh loops need it)")
        return
    try:
        import numpy as np
        from PIL import Image  # noqa: F401
    except ImportError as e:
        print(f"  (GUI checks skipped: {e})")
        return
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    import ui_dialogs
    import main as main_mod

    frame = np.zeros((480, 640, 3), np.uint8)
    frame[120:360, 160:480] = 255                   # a bright scrap on a dark bed

    def pump(ms):
        root.after(ms, root.quit)
        root.mainloop()

    def run_in_loop(fn, ms):
        """Call fn() from inside mainloop (the camera workers' after() needs
        the main thread in it) and come back `ms` later."""
        out = {}

        def go():
            try:
                out['r'] = fn()
            except Exception as e:                      # noqa: BLE001 - reported by the check
                out['e'] = e
        root.after(0, go)
        pump(ms)
        if 'e' in out:
            raise out['e']
        return out.get('r')

    saved = {"camera_index_override": 1}

    # ---- LiveCameraWindow: the one window that needs no calibration ----
    def live_window_falls_back_with_the_notice_without_a_calibration():
        calibration_on_disk(False)
        world = World(opens={0}, falcon=None, cams=[0, 1], names=["Integrated Camera"], frame=frame)

        def go(logged):
            win = run_in_loop(lambda: ui_dialogs.LiveCameraWindow(root, camera_index=1, settings=saved), 900)
            try:
                assert win._cap is not None and win._camera_index == 0, (win._cap, win._camera_index)
                assert win._status_var.get() == "Camera #1 could not be opened; showing camera #0.", win._status_var.get()
                assert len(warnings_in(logged)) == 1, logged
                assert saved == {"camera_index_override": 1}
            finally:
                win.close()
            assert world.held == set(), f"closing the window left a device held: {world.held}"
        with_world(world, go)
    check("LiveCameraWindow, no calibration, saved camera gone: runs on the other camera and says so in its status "
          "line; closing releases it", live_window_falls_back_with_the_notice_without_a_calibration)

    def live_window_refuses_under_a_calibration():
        calibration_on_disk(True)
        world = World(opens={0}, falcon=None, cams=[0, 1], names=["Integrated Camera"], frame=frame)

        def go(logged):
            win = run_in_loop(lambda: ui_dialogs.LiveCameraWindow(root, camera_index=1, settings=saved), 900)
            try:
                assert win._cap is None and win._camera_index == 1, "fell back under a calibration"
                assert win._status_var.get() == "", win._status_var.get()
                assert ("open", 0) not in world.calls, world.calls
                assert not warnings_in(logged), logged
            finally:
                win.close()
        with_world(world, go)
        # The saved camera opens: no notice.
        world = World(opens={1}, falcon=None, cams=[0, 1], frame=frame)

        def ok(logged):
            win = run_in_loop(lambda: ui_dialogs.LiveCameraWindow(root, camera_index=1, settings=saved), 700)
            try:
                assert win._cap is not None and win._camera_index == 1 and win._status_var.get() == ""
                assert world.calls == [("open", 1)], world.calls
            finally:
                win.close()
        with_world(world, ok)
        calibration_on_disk(False)
    check("LiveCameraWindow, a calibration on disk, saved camera gone, only a webcam: no fallback, no notice; "
          "saved camera fine: one open, no notice", live_window_refuses_under_a_calibration)

    # ---- CameraCaptureDialog: always calibrated ----
    def capture_dialog(world, ms, **kw):
        """Open the real dialog on the fake world, look at it `ms` later, close it."""
        seen = {}

        def look(d):
            seen['index'] = d._camera_index
            seen['cap'] = d._cap
            seen['status'] = d._status_var.get()

        def construct():
            def close():
                try:
                    look(d['dlg'])
                finally:
                    d['dlg']._on_close()
            d = {}
            real_init = ui_dialogs.CameraCaptureDialog.__init__

            class Recorded(ui_dialogs.CameraCaptureDialog):
                def __init__(self, *a, **k):
                    d['dlg'] = self
                    root.after(ms, close)
                    real_init(self, *a, **k)
            Recorded(root, 1, CAL_PATH, settings=saved, **kw)
            return seen
        return run_in_loop(construct, ms + 400)

    def capture_dialog_refuses_a_webcam_and_takes_a_falcon_named_camera():
        calibration_on_disk(True)
        world = World(opens={0}, falcon=None, cams=[0, 1], names=["Integrated Camera"], frame=frame)

        def refused(logged):
            seen = capture_dialog(world, 900)
            assert seen['cap'] is None and seen['index'] == 1, f"fell back to the webcam: {seen}"
            assert seen['status'] == "Could not open camera: Could not open camera index 1", seen['status']
            assert world.calls == [("open", 1), ("falcon",)], world.calls
            assert not warnings_in(logged), logged
        with_world(world, refused)
        # A Falcon-named camera at 2: the dialog runs on it with the notice, and the
        # detection line comes back once the notice has had its time.
        world = World(opens={0, 2}, falcon=2, cams=[0, 1, 2], names=["Integrated Camera", "Falcon Camera"], frame=frame)
        hold = ui_dialogs.CameraCaptureDialog.NOTICE_HOLD_S
        ui_dialogs.CameraCaptureDialog.NOTICE_HOLD_S = 1.2

        def falcon_named(logged):
            early = capture_dialog(world, 700)
            assert early['cap'] is not None and early['index'] == 2, early
            assert early['status'] == "Camera #1 could not be opened; showing camera #2.", early['status']
            assert len(warnings_in(logged)) == 1, logged
            world.calls.clear()
            late = capture_dialog(world, 2000)
            assert late['index'] == 2 and late['status'].startswith("Detected: "), late
        try:
            with_world(world, falcon_named)
        finally:
            ui_dialogs.CameraCaptureDialog.NOTICE_HOLD_S = hold
        calibration_on_disk(False)
    check("CameraCaptureDialog, saved camera gone: with only a webcam the old error and no fallback; with a "
          "Falcon-named camera it runs there, shows the notice, then the detection line resumes",
          capture_dialog_refuses_a_webcam_and_takes_a_falcon_named_camera)

    # ---- the app's own live preview gives the camera up before a capture ----
    with open(os.path.join(_PROFILE, "app_settings.json"), "w", encoding="utf-8") as f:
        json.dump({"experimental_machine_menu": True, "camera_index_override": 1,
                   "seen_polygon_tutorial": True}, f)
    real_supported = main_mod.machine_supported
    main_mod.machine_supported = lambda: True
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()

    def the_app_releases_its_live_preview_before_get_from_camera():
        calibration_on_disk(True)
        world = World(opens={1}, falcon=None, cams=[0, 1], names=["Falcon Camera"], frame=frame)
        seen = {}

        def go(logged):
            # A Frame & Cut with "Live camera preview" ticked leaves this open on the saved camera.
            app._live_cam_window = run_in_loop(
                lambda: ui_dialogs.LiveCameraWindow(root, camera_index=1, settings=app.settings), 700)
            live = app._live_cam_window
            assert live._cap is not None and world.held == {1}, "fixture: the live preview holds camera 1"
            opens_before = len(world.calls)
            real_poly, real_capture = ui_dialogs.PolygonDrawWindow, ui_dialogs.CameraCaptureDialog
            real_capture_init = real_capture.__init__

            class Capture(real_capture):
                def __init__(self, parent, *a, **k):
                    seen['live_open_at_capture'] = bool(live.winfo_exists())
                    seen['held_at_capture'] = set(world.held)

                    def close():
                        seen['index'] = self._camera_index
                        seen['cap'] = self._cap
                        seen['status'] = self._status_var.get()
                        self._on_close()
                    root.after(1200, close)
                    real_capture_init(self, parent, *a, **k)

            class Driven(real_poly):
                def __init__(self, parent, **k):
                    parent.after(150, self.on_get_from_camera)     # the user clicks Get from camera
                    parent.after(3200, self.on_cancel)
                    real_poly.__init__(self, parent, **k)
            main_mod.PolygonDrawWindow, ui_dialogs.CameraCaptureDialog = Driven, Capture
            try:
                run_in_loop(app.on_draw_custom_shape, 4000)
            finally:
                main_mod.PolygonDrawWindow, ui_dialogs.CameraCaptureDialog = real_poly, real_capture
            assert 'index' in seen, "the capture dialog never opened"
            assert seen['live_open_at_capture'] is False and seen['held_at_capture'] == set(), \
                f"the live preview still held the camera when the capture dialog opened: {seen}"
            assert seen['index'] == 1 and seen['cap'] is not None, f"the capture dialog did not get the saved camera: {seen}"
            assert "could not be opened" not in seen['status'], seen['status']
            assert not warnings_in(logged), logged
            assert world.calls[opens_before:] == [("open", 1)], f"the capture dialog's opens: {world.calls[opens_before:]}"
            assert app._live_cam_window is None and not live.winfo_exists(), "the live preview was not closed"
            assert world.held == set(), world.held
        with_world(world, go)
        calibration_on_disk(False)
    check("the app's live preview is open on the saved camera: Draw / Capture Shape > Get from camera closes it first "
          "and gets that camera, no fallback, no warning", the_app_releases_its_live_preview_before_get_from_camera)

    def camera_calibration_releases_it_too():
        world = World(opens={1}, falcon=None, cams=[1], frame=frame)
        seen = {}

        def go(logged):
            app._live_cam_window = run_in_loop(
                lambda: ui_dialogs.LiveCameraWindow(root, camera_index=1, settings=app.settings), 700)
            live = app._live_cam_window
            assert live._cap is not None and world.held == {1}

            class FakeCalibration:
                def __init__(self, parent, **k):
                    seen['live_open'] = bool(live.winfo_exists())
                    seen['held'] = set(world.held)
                    seen['index'] = k.get('camera_index')
            real = ui_dialogs.CameraCalibrationDialog
            ui_dialogs.CameraCalibrationDialog = FakeCalibration
            try:
                app._open_camera_calibration()
            finally:
                ui_dialogs.CameraCalibrationDialog = real
            assert seen.get('live_open') is False and seen.get('held') == set() and seen.get('index') == 1, seen
            assert app._live_cam_window is None
        with_world(world, go)
        # With no live preview open the release is a no-op.
        app._release_live_camera()
        assert app._live_cam_window is None
    check("Camera Calibration closes the app's live preview before its dialog opens the camera; with none open the "
          "release is a no-op", camera_calibration_releases_it_too)

    main_mod.machine_supported = real_supported
    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
