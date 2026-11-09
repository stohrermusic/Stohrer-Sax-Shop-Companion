"""The camera falls back when its saved number cannot be opened.

Matt, 2026-11-09, asked about the laser being plugged in while the app runs
and then, of the camera, "same deal?": yes. Each camera window used to
resolve an index when it opened (the saved number, camera_index_override,
winning outright) and simply fail when that number could not be opened.

Pinned here, against fakes (no OpenCV needed, no camera touched):
  - the saved number fails (or opens and gives no frame) and another camera
    opens: the window runs on it, one warning is logged with the number that
    failed, the number in use and the camera names the system reports, and
    the saved setting is untouched
  - the saved number opens: exactly one open call, no enumeration, no name
    lookup
  - nothing opens: the first error, as before
  - a number that is NOT the saved one fails: no fallback (that is a
    request for one specific camera)
  - all four windows that open a camera go through the one helper
    (camera_capture.open_camera_or_fallback) and none calls open_camera
    itself
"""
import inspect
import logging
import os
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-camera-fallback-")

from i18n import init_translation  # noqa: E402
init_translation("en")

import camera_capture as cam  # noqa: E402

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


class FakeCap:
    def __init__(self, index, frames=True):
        self.index, self.frames, self.released = index, frames, False

    def read(self):
        return (True, f"frame-from-{self.index}") if self.frames else (False, None)

    def release(self):
        self.released = True


class World:
    """What the machine has: which indices open, which give frames, the
    Falcon-named index, the enumeration, the names. Records every call."""

    def __init__(self, opens=(), no_frame=(), falcon=None, cams=(), names=()):
        self.opens, self.no_frame, self.falcon, self.cams, self.names = set(opens), set(no_frame), falcon, list(cams), list(names)
        self.calls = []
        self.caps = []

    def open_camera(self, index, *a, **k):
        self.calls.append(("open", index))
        if index not in self.opens:
            raise RuntimeError(f"Could not open camera index {index}")
        cap = FakeCap(index, frames=index not in self.no_frame)
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


def saved_number_fails_another_opens():
    settings = {"camera_index_override": 1}
    world = World(opens={0, 2}, falcon=None, cams=[0, 2], names=["Integrated Camera", "Falcon Camera"])

    def go(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, settings["camera_index_override"], settings)
        assert used == 2 and frame == "frame-from-2" and cap.index == 2, (used, frame)
        assert world.calls == [("open", 1), ("falcon",), ("enumerate",), ("open", 2), ("names",)], world.calls
        assert settings == {"camera_index_override": 1}, "the saved number must not change"
        warnings = [m for lvl, m in logged if lvl == "WARNING"]
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


check("saved number fails, another camera opens: the window runs on it, one warning with both numbers and the "
      "names, the saved setting untouched", saved_number_fails_another_opens)


def saved_number_opens_but_gives_no_frame():
    world = World(opens={1, 2}, no_frame={1}, falcon=None, cams=[1, 2], names=["A", "B"])

    def go(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert used == 2 and world.caps[0].released, "the frameless capture must be released before falling back"
        assert [m for lvl, m in logged if lvl == "WARNING"], logged
    with_world(world, go)


check("saved number opens but returns no frame: released, and the fallback runs", saved_number_opens_but_gives_no_frame)


def saved_number_opens_nothing_else_happens():
    world = World(opens={1}, falcon=0, cams=[0, 1], names=["x"])

    def go(logged):
        cap, used, frame = cam.open_camera_or_fallback(1, 1, {})
        assert (used, frame) == (1, "frame-from-1")
        assert world.calls == [("open", 1)], f"an extra call when the saved camera opened: {world.calls}"
        assert not [m for lvl, m in logged if lvl == "WARNING"], logged
    with_world(world, go)


check("saved number opens: exactly one open call, no enumeration, no name lookup, no warning",
      saved_number_opens_nothing_else_happens)


def nothing_opens_is_the_old_error():
    world = World(opens=set(), falcon=0, cams=[0, 2], names=["x"])

    def go(logged):
        try:
            cam.open_camera_or_fallback(1, 1, {})
        except RuntimeError as e:
            assert "camera index 1" in str(e), f"the FIRST error is raised: {e}"
        else:
            raise AssertionError("nothing opened and no error was raised")
        assert ("open", 0) in world.calls and ("open", 2) in world.calls
        assert not [m for lvl, m in logged if lvl == "WARNING"], "nothing to say when nothing worked"
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


def every_camera_window_goes_through_the_helper():
    import ui_dialogs
    windows = [ui_dialogs.PolygonDrawWindow, ui_dialogs.CameraCalibrationDialog,
               ui_dialogs.CameraCaptureDialog, ui_dialogs.LiveCameraWindow]
    for cls in windows:
        src = inspect.getsource(cls)
        assert "open_camera_or_fallback(" in src, f"{cls.__name__} does not use the helper"
        assert ".open_camera(" not in src, f"{cls.__name__} still opens the camera itself"
        assert 'get("camera_index_override")' in src, f"{cls.__name__} does not pass the saved number"
    whole = inspect.getsource(ui_dialogs)
    assert whole.count("open_camera_or_fallback(") == len(windows), \
        f"{whole.count('open_camera_or_fallback(')} call sites for {len(windows)} windows"
    assert ".open_camera(" not in whole.replace("open_camera_or_fallback(", ""), "a camera is opened outside the helper"


check("all four camera windows open through the one helper with the saved number", every_camera_window_goes_through_the_helper)

print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
