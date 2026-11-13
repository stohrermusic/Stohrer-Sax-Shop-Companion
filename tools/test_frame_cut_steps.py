"""Frame & Cut's steps, each on its own.

on_frame_and_cut (machine_control.MachineControlMixin) became a sequence of
_frame_cut_* steps handing on one FrameCutRun on 2026-11-10. This suite drives
each step headless, on a stand-in app with a fake sender and fake run dialogs,
then the whole sequence, for the guarantee each step owns:

  - origin:    G92 X0 Y0 for a job not on the loaded shape, the anchor vertex for one on it
  - framing:   M4 at the MATERIAL's framing power, the polygon traced or the cut's bounding box
  - position:  the jog-only dialog opts into the lid question and streams nothing
  - pad job:   flip with the outline, frame the outline, cut on the inset; the camera
               extras (Auto Locate target, live camera) only with a calibration
  - tooling:   never on the shape; the tool's after_cut is told how the cut ended
  - connect:   a failed connect shows the box, asks for a rescan, leaves no sender to disconnect
  - cut:       kept in last_frame_cut.gcode, G92 first, before it streams; the scrap number
  - record:    Job History; in Scrap Mode the scrap is committed only on a complete cut
  - the whole: a cancelled position step disconnects, releases the sleep lock, records
               nothing and abandons the pending scrap; a complete run does it all in order

Headless: no Tk root, no window. Run:
    python tools/test_frame_cut_steps.py
"""
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Never the user's real settings folder: an isolated profile, set before any
# project module is imported (tools/test_config_isolation.py checks this).
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-frame-cut-steps-")

from i18n import init_translation  # noqa: E402
init_translation("en")

import tkinter.messagebox as mb  # noqa: E402
import camera_capture  # noqa: E402
import falcon_sender  # noqa: E402
import gcode_engine as ge  # noqa: E402
import main  # noqa: E402
import sleep_lock  # noqa: E402
import ui_dialogs  # noqa: E402
from config import DEFAULT_SETTINGS  # noqa: E402
from machine_control import FrameCutRun, MachineControlMixin  # noqa: E402

PROFILE = os.environ["SAXSHOP_CONFIG_DIR"]
KEPT = os.path.join(PROFILE, "last_frame_cut.gcode")
results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


# ---- what the steps reach for, stood in -------------------------------------------------------
events, boxes, runs = [], [], []
for _name in ("showerror", "showwarning", "showinfo"):
    setattr(mb, _name, lambda title, msg, _n=_name, **k: (boxes.append((_n, str(title))),
                                                          events.append(f"box:{title}"))[0])
sleep_lock.prevent_sleep = lambda *a, **k: events.append("sleep:prevent")
sleep_lock.allow_sleep = lambda *a, **k: events.append("sleep:allow")


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeSender:
    def __init__(self, port=None):
        events.append(f"sender:new:{port}")

    def connect(self):
        events.append("sender:connect")

    def unlock(self):
        events.append("sender:unlock")

    def disconnect(self):
        events.append("sender:disconnect")


class DeadSender(FakeSender):
    def connect(self):
        events.append("sender:connect-failed")
        raise OSError("port busy")


class FakeRun:
    """A run dialog: ends as the script says ("complete" once the script runs out),
    records what it was asked to stream and with which options, and whether the
    kept file already held those lines when it opened."""
    script = []

    def __init__(self, parent, sender, lines, **kw):
        how = FakeRun.script.pop(0) if FakeRun.script else "complete"
        if how == "raise":
            events.append("run:raised")
            raise RuntimeError("the dialog could not be built")
        kept = open(KEPT).read().splitlines() if os.path.exists(KEPT) else None
        runs.append({"lines": list(lines), "kw": dict(kw), "ends": how, "kept_held_these": kept == list(lines)})
        events.append(f"run:{how}")
        self._final_reason = how


class FakePreview:
    answer = "save"

    def __init__(self, parent, placements, w_mm, h_mm, **kw):
        events.append("preview")
        self.result = FakePreview.answer


class Ended:
    """A cut dialog that has ended."""

    def __init__(self, how):
        self._final_reason = how


# A 100 x 80 mm camera scrap in storage coordinates (Y down), its 5 mm inset, a tilted
# one whose anchor vertex is not at (0, 0), and where the anchor is on the machine.
OUTLINE = [(0.0, 0.0), (100.0, 0.0), (100.0, 80.0), (0.0, 80.0)]
INSET = [(5.0, 5.0), (95.0, 5.0), (95.0, 75.0), (5.0, 75.0)]
TILTED = [(20.0, 0.0), (100.0, 30.0), (80.0, 80.0), (0.0, 50.0)]
LB_MACHINE = (200.0, 100.0)
PADS = [{'size': 20.0, 'qty': 2}]


class App(MachineControlMixin):
    """The app as the steps see it: the Pad Maker's form answers canned (two 20 mm
    felt pads on a 200 x 150 sheet, or on the loaded shape), the bookkeeping
    recorded instead of done."""

    def __init__(self, **over):
        self.root = None
        self.settings = dict(DEFAULT_SETTINGS)
        self.settings["laser_framing_power_by_material"] = {"felt": 10, "leather": 25}
        self.settings["laser_framing_feed"] = 2000
        self.falcon_port = "TEST"
        self._falcon_detect_attempted = True
        self._falcon_homed_this_session = False
        self.custom_polygon = None
        self.custom_polygon_outline = None
        self._custom_polygon_lb_machine = None
        self.calibrated = False
        self.scrap_mode_var = FakeVar(False)
        self.live_camera_var = FakeVar(False)
        self.material_vars = {m: FakeVar(m == "felt") for m in ("felt", "card", "leather", "exact_size")}
        self.scrap_session = {"scrap_count": 0, "hole_dia": 0}
        self.recorded, self.advanced = [], []
        self.abandoned = self.detects = 0
        for k, v in over.items():
            setattr(self, k, v)

    def _prepare_generation(self):
        return {'pads': list(PADS), 'hole_dia': 0, 'base': 'pads',
                'width_mm': 200.0, 'height_mm': 150.0, 'card_paper_dims': None}

    def _get_material_dimensions(self, material, width_mm, height_mm, card_paper_dims):
        return width_mm, height_mm, self.custom_polygon

    def _camera_capture_ready(self):
        return self.calibrated

    def _detect_falcon_async(self):
        self.detects += 1

    def _resolve_camera_index(self):
        return None

    def _record_job(self, output, materials, pads, params, placed_count=None, scrap_num=None, status=None):
        self.recorded.append(("pad", output, materials, placed_count, scrap_num, status))

    def _record_tooling_job(self, job, output, status="complete", save_dir=None):
        self.recorded.append(("tool", job.get('label'), output, status))

    def _scrap_abandon_pending(self):
        self.abandoned += 1

    def _frame_cut_scrap_advance(self, placed_count, remaining):
        self.advanced.append((placed_count, remaining))


def machine(fn, script=(), sender=FakeSender):
    """fn() against the stand-ins, from a clean slate; what it returned (the dialogs it
    opened are in `runs`, everything else in `events`)."""
    del events[:], boxes[:], runs[:]
    if os.path.exists(KEPT):
        os.remove(KEPT)
    FakeRun.script = list(script)
    keep = (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main.NestingPreviewWindow)
    ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main.NestingPreviewWindow = FakeRun, sender, FakePreview
    try:
        return fn()
    finally:
        ui_dialogs.FalconRunDialog, falcon_sender.FalconSender, main.NestingPreviewWindow = keep
        FakeRun.script = []


print("Frame & Cut's steps")
print("=" * 60)


def the_run_object_refuses_a_misspelled_field():
    run = FrameCutRun()
    run.framing_lines = []
    try:
        run.framin_lines = []
    except AttributeError:
        return
    raise AssertionError("a misspelled field was silently created")
check("FrameCutRun refuses a field it does not have (slots)", the_run_object_refuses_a_misspelled_field)


def origin_off_the_shape():
    app = App(custom_polygon=INSET, custom_polygon_outline=TILTED)   # a shape IS loaded; the job is not on it
    run = FrameCutRun(on_shape=False, framing_lines=["M4 S10", "G1 X1 Y1"])
    app._frame_cut_origin(run)
    assert run.g92 == (0.0, 0.0), run.g92
    assert run.framing_lines == ["G92 X0.000 Y0.000", "M4 S10", "G1 X1 Y1"], run.framing_lines
    run = FrameCutRun(on_shape=False, framing_lines=[])                # nothing to frame: no prefix to add
    app._frame_cut_origin(run)
    assert run.g92 == (0.0, 0.0) and run.framing_lines == []
check("origin: a job not on the loaded shape gets G92 X0 Y0, whatever shape is loaded", origin_off_the_shape)


def origin_on_the_shape():
    app = App(custom_polygon=INSET, custom_polygon_outline=TILTED)
    run = FrameCutRun(on_shape=True, framing_lines=["M4 S10"])
    app._frame_cut_origin(run)
    ax, ay = app._frame_cut_anchor()
    assert (ax, ay) != (0.0, 0.0), "fixture: the tilted outline's anchor is off the bbox corner"
    assert run.g92 == (ax, ay), (run.g92, (ax, ay))
    assert run.framing_lines == [f"G92 X{ax:.3f} Y{ay:.3f}", "M4 S10"], run.framing_lines
check("origin: a job on the loaded shape gets G92 at the outline's anchor vertex", origin_on_the_shape)


def framing_lines_at_the_materials_power():
    app = App()
    run = FrameCutRun(material="leather", framing_polygon=OUTLINE, gcode_text="G1 X0 Y0")
    app._frame_cut_framing_lines(run)
    assert run.framing_lines == ge.generate_polygon_framing_gcode(OUTLINE, power_s=25, feed=2000), run.framing_lines
    assert "M4 S25" in run.framing_lines and not any("M3" in ln for ln in run.framing_lines)
    run = FrameCutRun(material="felt", framing_polygon=None, gcode_text="G0 X10 Y20\nG1 X70 Y60")
    app._frame_cut_framing_lines(run)
    assert run.framing_lines == ge.generate_framing_gcode(10.0, 20.0, 70.0, 60.0, power_s=10, feed=2000)
    assert "M4 S10" in run.framing_lines
    run = FrameCutRun(material="felt", framing_polygon=None, gcode_text="M5")
    app._frame_cut_framing_lines(run)
    assert run.framing_lines == [], run.framing_lines
check("framing: M4 at the material's own power, the polygon traced or the cut's bbox; nothing to frame, no lines",
      framing_lines_at_the_materials_power)


def position_opts_into_the_lid_question():
    app = App(_falcon_homed_this_session=True)
    run = FrameCutRun(sender=object(), auto_locate_target=LB_MACHINE)
    assert machine(lambda: app._frame_cut_position(run)) is True
    kw = runs[0]["kw"]
    assert runs[0]["lines"] == [], "the jog dialog streams nothing"
    assert kw["remind_lid_on_advance"] is True and "Start Frame" in kw["done_button_label"], kw
    assert kw["auto_locate_target"] == LB_MACHINE and kw["is_homed"] is True, kw
    assert kw["show_pause_resume"] is False and kw["show_cut_button"] is False and kw["stop_needs_confirm"] is False
    assert machine(lambda: app._frame_cut_position(run), script=["cancelled"]) is None
check("position: the jog-only dialog streams nothing, opts into the lid question, carries the Auto Locate target; "
      "cancelled stops the run", position_opts_into_the_lid_question)


def pad_job_on_a_camera_scrap():
    app = App(custom_polygon=INSET, custom_polygon_outline=OUTLINE, _custom_polygon_lb_machine=LB_MACHINE,
              live_camera_var=FakeVar(True))
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_prepare_pad_job(run)) is True and not boxes, boxes
    assert run.on_shape is True and run.flip_height_mm == 80.0, (run.on_shape, run.flip_height_mm)
    assert run.framing_polygon == OUTLINE and run.cut_polygon == INSET
    assert run.material == "felt" and run.piece_count == len(run.placed) == 2 and run.pads == PADS
    assert run.gcode_text and "G1" in run.gcode_text and run.job_label is None and run.scrap_mode is False
    # No calibration: no Auto Locate target and no live camera, whatever is known or ticked.
    assert run.auto_locate_target is None and run.show_live_camera is False
    app.calibrated = True
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_prepare_pad_job(run)) is True
    assert run.auto_locate_target == LB_MACHINE and run.show_live_camera is True
check("pad job on a camera scrap: flip with the outline's height, frame the outline, cut on the inset; "
      "Auto Locate and the live camera only with a calibration", pad_job_on_a_camera_scrap)


def pad_job_off_the_shape():
    # Card on the paper size: the material dimensions come back with no polygon although a shape is loaded.
    app = App(custom_polygon=INSET, custom_polygon_outline=OUTLINE, _custom_polygon_lb_machine=LB_MACHINE,
              calibrated=True)
    app._get_material_dimensions = lambda material, w, h, dims: (w, h, None)
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_prepare_pad_job(run)) is True and not boxes, boxes
    assert run.on_shape is False and run.flip_height_mm is None and run.framing_polygon is None
    assert run.cut_polygon is None and run.auto_locate_target is None
check("pad job not on the loaded shape (card on the paper size): no flip, no outline, no Auto Locate target",
      pad_job_off_the_shape)


def pad_job_refusals_stop_the_run():
    app = App()
    app.material_vars["card"].set(True)                       # two materials ticked
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_prepare_pad_job(run)) is None
    assert boxes == [("showerror", "Pick One Material")], boxes
    assert "preview" not in events and run.gcode_text is None
    app = App()
    FakePreview.answer = None                                 # Adjust on the nesting preview
    try:
        run = FrameCutRun()
        assert machine(lambda: app._frame_cut_prepare_pad_job(run)) is None
    finally:
        FakePreview.answer = "save"
    assert "preview" in events and not boxes and run.gcode_text is None
    app = App()
    app._prepare_generation = lambda: None                    # the form was refused (its own box)
    assert machine(lambda: app._frame_cut_prepare_pad_job(FrameCutRun())) is None and "preview" not in events
check("pad job: two materials, Adjust on the preview or a refused form stop the run before anything is generated",
      pad_job_refusals_stop_the_run)


def tooling_job_is_never_on_the_shape():
    app = App(custom_polygon=INSET, custom_polygon_outline=OUTLINE, _custom_polygon_lb_machine=LB_MACHINE,
              calibrated=True, live_camera_var=FakeVar(True))
    assert app._frame_cut_prepare_tooling_job(FrameCutRun(tooling_job=lambda: None)) is None
    job = {'gcode_text': "G1 X1 Y1", 'material': 'acrylic', 'label': 'die inserts', 'count': 3, 'scrap_num': 2}
    run = FrameCutRun(tooling_job=lambda: job)
    assert app._frame_cut_prepare_tooling_job(run) is True
    assert run.job is job and run.gcode_text == "G1 X1 Y1" and run.material == "acrylic"
    assert run.job_label == "die inserts" and run.piece_count == 3
    assert run.on_shape is False and run.framing_polygon is None and run.flip_height_mm is None
    assert run.cut_polygon is None and run.auto_locate_target is None
    assert run.show_live_camera is True                       # the tick, with a calibration
    app.calibrated = False
    run = FrameCutRun(tooling_job=lambda: job)
    app._frame_cut_prepare_tooling_job(run)
    assert run.show_live_camera is False
check("tooling job: the tool's own values, never the loaded shape; the live camera only with a calibration",
      tooling_job_is_never_on_the_shape)


def connect_failure_leaves_nothing_to_disconnect():
    app = App()
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_connect(run)) is True
    assert isinstance(run.sender, FakeSender) and events == ["sender:new:TEST", "sender:connect"], events
    run = FrameCutRun()
    assert machine(lambda: app._frame_cut_connect(run), sender=DeadSender) is None
    assert run.sender is None and boxes == [("showerror", "Connection Failed")] and app.detects == 1, (boxes, run.sender)
check("connect: the connected sender on the run; a failed connect shows the box, asks for a rescan, leaves no sender",
      connect_failure_leaves_nothing_to_disconnect)


def the_live_camera_step():
    app = App()
    built = []

    class FakeLive:
        def __init__(self, parent, camera_index=None, settings=None):
            built.append(camera_index)

    class Existing:
        def winfo_exists(self):
            return True

        def lift(self):
            lifted.append("lift")

        def focus_force(self):
            lifted.append("focus")

    lifted = []
    keep = ui_dialogs.LiveCameraWindow, camera_capture.HAS_OPENCV
    ui_dialogs.LiveCameraWindow, camera_capture.HAS_OPENCV = FakeLive, True
    try:
        app._frame_cut_live_camera(FrameCutRun(show_live_camera=False))
        assert built == []
        app._resolve_camera_index = lambda: 1
        app._frame_cut_live_camera(FrameCutRun(show_live_camera=True))
        assert built == [1] and isinstance(app._live_cam_window, FakeLive), built
        app._live_cam_window = Existing()                     # left from the last run: lifted, not doubled
        app._frame_cut_live_camera(FrameCutRun(show_live_camera=True))
        assert built == [1] and lifted == ["lift", "focus"], (built, lifted)
    finally:
        ui_dialogs.LiveCameraWindow, camera_capture.HAS_OPENCV = keep
check("live camera: nothing without the tick; opened once; a window left from the last run is lifted, not doubled",
      the_live_camera_step)


def the_cut_is_kept_before_it_streams():
    app = App()
    run = FrameCutRun(gcode_text="G0 X1 Y1\nG1 X2 Y2", material="felt", sender=object(), g92=(1.5, 2.0),
                      flip_height_mm=80.0, cut_polygon=INSET, piece_count=2)
    machine(lambda: app._frame_cut_cut(run))
    assert run.cut_lines == ["G92 X1.500 Y2.000", "G0 X1 Y1", "G1 X2 Y2"], run.cut_lines
    assert runs[0]["lines"] == run.cut_lines and runs[0]["kept_held_these"], runs
    assert open(KEPT).read().splitlines() == run.cut_lines
    assert run.cut_scrap_num is None and run.cut_dlg._final_reason == "complete"
    assert "Cutting" in runs[0]["kw"]["title"] and "felt" in runs[0]["kw"]["title"], runs[0]["kw"]
    # The scrap number: one past the session's count for a pad job in Scrap Mode, the job's own for a tool.
    app.scrap_session["scrap_count"] = 2
    run = FrameCutRun(gcode_text="G1 X1 Y1", material="felt", sender=object(), g92=(0.0, 0.0), scrap_mode=True)
    machine(lambda: app._frame_cut_cut(run))
    assert run.cut_scrap_num == 3, run.cut_scrap_num
    run = FrameCutRun(gcode_text="G1 X1 Y1", material="acrylic", sender=object(), g92=(0.0, 0.0),
                      tooling_job=lambda: None, job={'scrap_num': 7}, job_label="die inserts")
    machine(lambda: app._frame_cut_cut(run))
    assert run.cut_scrap_num == 7, run.cut_scrap_num
check("cut: kept in last_frame_cut.gcode, G92 first, before the dialog streams it; the scrap number per kind of job",
      the_cut_is_kept_before_it_streams)


def record_a_pad_job():
    app = App()
    run = FrameCutRun(material="felt", pads=PADS, params={'pads': PADS}, placed=[1, 2], cut_dlg=Ended("complete"))
    app._frame_cut_record(run)
    assert app.recorded == [("pad", "laser", ["felt"], 2, None, "complete")], app.recorded
    assert app.advanced == []                                 # not Scrap Mode
    app = App()
    run = FrameCutRun(material="felt", pads=PADS, params={}, placed=[1, 2], cut_dlg=Ended(None))
    app._frame_cut_record(run)
    assert app.recorded[0][-1] == "stopped", app.recorded    # a cut that never said how it ended
    # Scrap Mode: the scrap is committed only on a complete cut.
    app = App()
    app.scrap_session["scrap_count"] = 1
    remaining = [{'size': 20.0, 'qty': 5}]
    run = FrameCutRun(material="felt", pads=PADS, params={}, placed=[1, 2], scrap_mode=True,
                      scrap_remaining=remaining, cut_dlg=Ended("stopped"))
    app._frame_cut_record(run)
    assert app.recorded == [("pad", "laser", ["felt"], 2, 2, "stopped")] and app.advanced == [], app.recorded
    run.cut_dlg = Ended("complete")
    app._frame_cut_record(run)
    assert app.advanced == [(2, remaining)], app.advanced
check("record, pad job: Job History with how the cut ended; in Scrap Mode the scrap is committed only on a complete cut",
      record_a_pad_job)


def record_a_tooling_job():
    told = []
    job = {'gcode_text': "G1", 'material': 'acrylic', 'label': 'kerf test', 'count': 1, 'after_cut': told.append}
    app = App()
    run = FrameCutRun(tooling_job=lambda: job, job=job, material="acrylic", cut_dlg=Ended("stopped"))
    app._frame_cut_record(run)
    assert told == ["stopped"] and app.recorded == [], (told, app.recorded)     # no 'tool': History takes nothing
    job['tool'] = 'kerf_test'
    run.cut_dlg = Ended("complete")
    app._frame_cut_record(run)
    assert told == ["stopped", "complete"], told
    assert app.recorded == [("tool", "kerf test", "laser", "complete")], app.recorded
check("record, tooling job: after_cut is told how the cut ended; Job History only when the job names its tool",
      record_a_tooling_job)


def a_cancelled_position_step():
    app = App()
    machine(app.on_frame_and_cut, script=["cancelled"])
    assert events == ["sleep:prevent", "preview", "sender:new:TEST", "sender:connect", "sender:unlock",
                      "run:cancelled", "sender:disconnect", "sleep:allow"], events
    assert app.recorded == [] and not os.path.exists(KEPT), "nothing is recorded or kept"
    assert app.abandoned == 1 and not boxes, (app.abandoned, boxes)
check("the whole run, cancelled at the jog: disconnect, release the lock, abandon the pending scrap, record nothing",
      a_cancelled_position_step)


def a_complete_run():
    app = App()
    machine(app.on_frame_and_cut)
    assert events == ["sleep:prevent", "preview", "sender:new:TEST", "sender:connect", "sender:unlock",
                      "run:complete", "run:complete", "run:complete", "sender:disconnect", "sleep:allow"], events
    assert runs[0]["lines"] == [] and [r["lines"][0] for r in runs[1:]] == ["G92 X0.000 Y0.000"] * 2, runs
    assert runs[1]["kw"].get("loop") is True and runs[2]["kept_held_these"], runs
    assert app.recorded == [("pad", "laser", ["felt"], 2, None, "complete")] and app.abandoned == 1, app.recorded
check("the whole run: lock, nest, preview, connect, unlock, jog, framing loop, cut (kept), record, disconnect, release",
      a_complete_run)


def a_run_that_cannot_connect_or_has_no_laser():
    app = App()
    machine(app.on_frame_and_cut, sender=DeadSender)
    assert events == ["sleep:prevent", "preview", "sender:new:TEST", "sender:connect-failed",
                      "box:Connection Failed", "sleep:allow"], events
    assert app.detects == 1 and app.abandoned == 1 and app.recorded == []
    app = App(falcon_port=None)
    machine(app.on_frame_and_cut)
    assert events == ["box:Falcon Not Detected"] and app.detects == 1 and app.abandoned == 0, events
check("no laser, or a laser that will not connect: the box, a rescan asked for, nothing opened or held",
      a_run_that_cannot_connect_or_has_no_laser)


def a_step_that_raises():
    app = App()
    machine(app.on_frame_and_cut, script=["complete", "raise"])          # the framing dialog raises
    assert events[-4:] == ["run:raised", "sender:disconnect", "box:An Error Occurred", "sleep:allow"], events
    assert app.recorded == [] and app.abandoned == 1
check("a step that raises after the port is open: disconnect, the error box, release, abandon", a_step_that_raises)


print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
