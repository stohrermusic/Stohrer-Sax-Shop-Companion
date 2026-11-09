"""Tooling > Frame & Cut: every Tooling tool that writes G-code can send it
straight to the laser, with the Pad Maker's Frame & Cut sequence.

Die Inserts (and its Scrap Mode), Die Holders, Die Organizer, Kerf Test,
Speed & Power Test and Tonehole Crescent, each driven through its real
Frame & Cut button handler on the live tab against a mock laser: a fake sender and fake run dialogs that record
what they were handed. Nothing here touches a serial port.

What is pinned, per tool:
  - the cut that is streamed is the G-code the file export writes for the
    same form, line for line, behind one G92 X0 Y0 line; it is kept at
    last_frame_cut.gcode with one app.log line; no file lands in the user's
    folders (the Speed & Power legend excepted)
  - framing traces that G-code's bounding box at the material's framing
    power (acrylic for dies and holders, the selected material for the kerf
    and Speed & Power tests), with M4 and never M3
  - the lid question is asked once, at Start Frame (the jog dialog is the
    only one that opts in), and the jog dialog is worded as the Pad Maker's
  - cancelling at the jog step, stopping the framing, or stopping the cut
    consumes nothing
  - Speed & Power: the legend is written where the user chose before
    anything moves, and backing out of that prompt backs out of the run
  - Die Scrap Mode: one sheet per click, counted only when its cut finished
    as "complete"; a later file export in that session asks for its folder
  - the same refusals, in the same words, as Generate G-code
  - one gating rule for both tabs (Matt, 2026-11-09: Frame & Cut needs no
    camera calibration "on either tab"): the buttons show with machine
    integration on and a laser detected, never without pyserial; the
    Machine menu's laser items are enabled by the toggle alone, and only
    the camera-polygon inset waits for a calibration
  - with no calibration a Pad Maker run opens no camera and has no Auto
    Locate target, even with a stale tick and a stale position
  - a scrap shape loaded on the Pad Maker cannot bend a Tooling job
  - the Pad Maker's own Frame & Cut still streams the engine's output for
    a standard job, a hand-drawn polygon, a camera outline and scrap mode

Nothing is sampled on a timer. The run dialogs and previews are fakes that
answer at once; the one real window (the Speed & Power preview's button
label) is inspected from inside its own wait_window, scheduled before it
is built.

Isolated profile (SAXSHOP_CONFIG_DIR). Self-skips without a display.
"""
import json
import logging
import os
import re
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-tooling-frame-cut-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
# The profile of someone who has switched machine integration on, so the
# app builds its Options > Machine menu (it is only built at startup).
with open(os.path.join(_PROFILE, "app_settings.json"), "w") as _f:
    json.dump({"experimental_machine_menu": True}, _f)

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


# One framing power per material, all different, so a run that frames at
# another material's power (or at the default 10) cannot pass.
FRAMING_S = {"felt": 12, "card": 14, "leather": 27, "acrylic": 19, "basswood": 21, "cork": 23}
G92_ZERO = "G92 X0.000 Y0.000"


def own_extent(lines):
    """X and Y extent of a G-code listing, parsed here rather than with the
    engine's extract_gcode_bbox, so the framing check does not lean on the
    function the app frames with."""
    xs, ys = [], []
    for ln in lines:
        if ln.lstrip().startswith(";"):
            continue
        ln = ln.split(";")[0]
        xs += [float(v) for v in re.findall(r"X(-?\d+(?:\.\d+)?)", ln)]
        ys += [float(v) for v in re.findall(r"Y(-?\d+(?:\.\d+)?)", ln)]
    return min(xs), min(ys), max(xs), max(ys)


def main():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()
    # Check names and failure messages carry an arrow and a dash; a console
    # in a legacy code page must not turn a FAIL into a crash.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass

    print("Tooling: Frame & Cut")
    print("=" * 60)
    import builtins
    import inspect
    import config
    import falcon_sender
    import gcode_engine as ge
    import main as main_mod
    import sleep_lock
    import svg_engine as se
    import tooling_tab
    import ui_dialogs

    # pyserial is not in every build (the Intel Mac has none); the menu and
    # the buttons are what is under test here, not the library.
    real_machine_supported = main_mod.machine_supported
    main_mod.machine_supported = lambda: True
    app = main_mod.PadSVGGeneratorApp(root)
    assert app._machine_enabled() and getattr(app, "_machine_menu", None) is not None, \
        "fixture: the app should have built its Machine menu"
    app.notebook.select(app.tooling_tab_frame)
    root.update()
    app.settings["laser_framing_power_by_material"].update(FRAMING_S)
    feed = app.settings.get("laser_framing_feed", 2000)
    kept_path = os.path.join(_PROFILE, "last_frame_cut.gcode")

    # ---- the mock shop: dialogs, laser, log ------------------------------
    out = tempfile.mkdtemp(prefix="ssc-tooling-frame-cut-out-")
    events, runs, boxes, saved, folders, log_lines, previews, temp_files = [], [], [], [], [], [], [], []
    kerf_shown, camera_asks = [], []
    answer = {"save": "auto", "folder": out, "preview": "save"}

    def fake_save(**kw):
        events.append("save-dialog")
        if answer["save"] is None:
            saved.append(None)
            return ""
        name = kw.get("initialfile") if answer["save"] == "auto" else answer["save"]
        path = os.path.join(out, name)
        saved.append(path)
        return path

    def fake_folder(**kw):
        folders.append(str(kw.get("title")))
        return answer["folder"] or ""

    tooling_tab.filedialog.asksaveasfilename = fake_save
    tooling_tab.filedialog.askdirectory = fake_folder
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))

    class Grab(logging.Handler):
        def emit(self, record):
            msg = record.getMessage()
            if "frame&cut" in msg:
                log_lines.append(msg)
    logging.getLogger().addHandler(Grab())

    class FakeSender:
        def __init__(self, port=None):
            events.append("sender:new")

        def connect(self):
            events.append("sender:connect")

        def unlock(self):
            events.append("sender:unlock")

        def disconnect(self):
            events.append("sender:disconnect")

    class DeadSender(FakeSender):
        def connect(self):
            raise OSError("port busy")

    class FakeRun:
        script = []      # how each dialog in turn ends; "complete" once it runs out

        def __init__(self, parent, sender, lines, **kw):
            self._final_reason = FakeRun.script.pop(0) if FakeRun.script else "complete"
            runs.append({"lines": list(lines), "kw": dict(kw), "ends": self._final_reason,
                         "files_then": sorted(os.listdir(out))})
            events.append(f"run:{self._final_reason}")

    class FakeFSPreview:
        def __init__(self, parent, pieces, w_mm, h_mm, **kw):
            previews.append(dict(kw))
            events.append("preview")
            self.result = answer["preview"]

    class FakeNest:
        def __init__(self, parent, placements, w_mm, h_mm, **kw):
            previews.append({"placements": placements, **kw})
            self.result = "save"

    real_named_temp = tempfile.NamedTemporaryFile

    def recording_temp(*a, **k):
        f = real_named_temp(*a, **k)
        temp_files.append(f.name)
        return f

    app._show_kerf_instructions = lambda name: kerf_shown.append(name)
    app._resolve_camera_index = lambda: camera_asks.append(1)   # returns None: no window
    app._open_remaining_pads_window = lambda: None
    app._update_remaining_pads_window = lambda: None
    pad_continues = []
    app._show_scrap_continue_dialog = lambda *a: pad_continues.append(a)

    def machine(fn, script=(), sender=FakeSender):
        """One click of a Frame & Cut button, against the mock laser."""
        del events[:], runs[:], boxes[:], log_lines[:], previews[:], temp_files[:]
        del kerf_shown[:], camera_asks[:]
        FakeRun.script = list(script)
        keep = (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender,
                tooling_tab.SpeedPowerTestPreview, main_mod.NestingPreviewWindow,
                sleep_lock.prevent_sleep, sleep_lock.allow_sleep, tempfile.NamedTemporaryFile)
        ui_dialogs.FalconRunDialog = FakeRun
        falcon_sender.FalconSender = sender
        tooling_tab.SpeedPowerTestPreview = FakeFSPreview
        main_mod.NestingPreviewWindow = FakeNest
        sleep_lock.prevent_sleep = lambda *a, **k: events.append("sleep:prevent")
        sleep_lock.allow_sleep = lambda *a, **k: events.append("sleep:allow")
        tempfile.NamedTemporaryFile = recording_temp
        app.falcon_port = "TEST"
        try:
            fn()
            root.update()
        finally:
            (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender,
             tooling_tab.SpeedPowerTestPreview, main_mod.NestingPreviewWindow,
             sleep_lock.prevent_sleep, sleep_lock.allow_sleep, tempfile.NamedTemporaryFile) = keep
            FakeRun.script = []
        return [dict(r) for r in runs]

    def errors():
        return [b for b in boxes if b[0] == "showerror"]

    def read_lines(path):
        with open(path) as f:           # the encoding the app reads its own G-code with
            return f.read().splitlines()

    def kept_state():
        if not os.path.exists(kept_path):
            return None
        with open(kept_path) as f:
            return f.read()

    # ---- the four tools ----------------------------------------------------
    def setup_die():
        app.die_scrap_var.set(False)
        app._clear_die_scrap_session()
        app.die_units_var.set("mm")
        app.die_width_var.set("300")
        app.die_height_var.set("300")
        app.die_size_entry.delete("1.0", tk.END)
        app.die_size_entry.insert("1.0", "30.0, 35.5, 42.0")
        app.die_filename_var.set("dies")

    def setup_holder():
        app.holder_variant_var.set("small")
        app.holder_layer_count_var.set(5)
        app.holder_units_var.set("mm")
        app.holder_width_var.set("300")
        app.holder_height_var.set("300")
        app.holder_filename_var.set("holder")

    def setup_kerf():
        app.kerf_material_var.set("Felt")
        app.kerf_use_existing_var.set(True)

    def setup_fs():
        app.fs_material_var.set("Leather")
        app.fs_show_preview_var.set(False)
        app.fs_filename_var.set("sp")
        app.fs_diameter_var.set("20")
        app.fs_inner_diameter_var.set("0")

    def setup_crescent():
        app.crescent_list.clear()
        app._refresh_crescent_list()
        app.crescent_tonehole_var.set("43")
        app.crescent_percent_var.set("20")
        app.crescent_curve_var.set(75)
        app.crescent_qty_var.set("2")
        app.crescent_material_var.set("Cork")

    def setup_organizer():
        app.organizer_variant_var.set("lower")
        app.organizer_material_var.set("Basswood")
        app.organizer_units_var.set("mm")
        app.organizer_width_var.set("300")
        app.organizer_height_var.set("400")
        app.organizer_filename_var.set("org")

    TOOLS = [
        {"name": "die inserts", "setup": setup_die, "export": app._on_generate_die_gcode,
         "click": app._on_frame_cut_die, "material": "acrylic", "legend": False},
        {"name": "die holder", "setup": setup_holder, "export": app._on_generate_holder_gcode,
         "click": app._on_frame_cut_holder, "material": "acrylic", "legend": False},
        {"name": "kerf test", "setup": setup_kerf, "export": app._on_generate_kerf_gcode,
         "click": app._on_frame_cut_kerf, "material": "felt", "legend": False},
        {"name": "speed & power test", "setup": setup_fs, "export": app._on_generate_feeds_speeds_gcode,
         "click": app._on_frame_cut_feeds_speeds, "material": "leather", "legend": True},
        {"name": "tonehole crescent", "setup": setup_crescent, "export": app._on_generate_crescent_gcode,
         "click": app._on_frame_cut_crescent, "material": "cork", "legend": False},
        {"name": "die organizer lower", "setup": setup_organizer, "export": app._on_generate_organizer_gcode,
         "click": app._on_frame_cut_organizer, "material": "basswood", "legend": False},
    ]

    def exported(tool):
        """What Generate G-code writes for the tool's form, as lines."""
        tool["setup"]()
        answer["save"] = "auto"
        del boxes[:]
        before = len(saved)
        tool["export"]()
        root.update()
        assert not errors(), errors()
        assert len(saved) == before + 1 and saved[-1].endswith(".gcode"), saved[before:]
        return read_lines(saved[-1])

    def pad_job_history():
        return [j for j in config.load_job_history()]

    # ---- per tool: the cut, the frame, the lid, the cancels -----------------
    pad_jog = {}

    def pad_form(pads_text, material, w=200, h=150):
        app.settings["units"] = "mm"
        app.settings["show_engraving_warning"] = False
        app.settings["gcode_cut_grouping"] = "layer"
        app.preview_var.set(False)
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", pads_text)
        app.width_entry.delete(0, tk.END)
        app.width_entry.insert(0, str(w))
        app.height_entry.delete(0, tk.END)
        app.height_entry.insert(0, str(h))
        for m, v in app.material_vars.items():
            v.set(m == material)
        root.update()

    def pad_reset():
        app.custom_polygon = None
        app.custom_polygon_outline = None
        app._custom_polygon_lb_machine = None
        app._clear_scrap_session()
        app.scrap_mode_var.set(False)
        app.live_camera_var.set(False)

    def engine_pad_cut(placed, material, w, h, polygon=None, flip=None):
        """The engine's own G-code for a nest, with the hole size the form asks for."""
        hole_dia = app._prepare_generation()["hole_dia"]
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "want.gcode")
            ge.generate_gcode_from_placed(placed, material, w, h, path, hole_dia, app.settings,
                                          polygon=polygon, flip_height_mm=flip)
            return read_lines(path)

    def the_pad_maker_still_streams_the_engine_output():
        # Standard job on a rectangle: G92 at the corner, the bounding box framed.
        pad_reset()
        pad_form("20 x 4\n30.5 x 3", "felt")
        hist = len(pad_job_history())
        got = machine(app.on_frame_and_cut)
        assert not errors(), errors()
        assert len(got) == 3, f"jog, framing, cut expected; got {len(got)}"
        jog, framing, cut = got
        pad_jog.update(jog["kw"])
        placed = previews[0]["placements"]["felt"]
        assert cut["lines"] == [G92_ZERO] + engine_pad_cut(placed, "felt", 200.0, 150.0),             "the pad cut is not the engine's output behind G92 X0 Y0"
        x0, y0, x1, y1 = own_extent(cut["lines"][1:])
        assert framing["lines"] == [G92_ZERO] + ge.generate_framing_gcode(
            x0, y0, x1, y1, power_s=FRAMING_S["felt"], feed=feed)
        assert jog["kw"].get("remind_lid_on_advance") is True and jog["kw"].get("auto_locate_target") is None
        new = pad_job_history()
        assert len(new) == hist + 1 and new[0]["output"] == "laser" and new[0]["materials"] == ["felt"]
        assert len(log_lines) == 1 and log_lines[0].startswith("frame&cut felt scrap -: 7 discs, "), log_lines
    check("Pad Maker: a standard job still streams the engine's G-code behind G92, frames its box, logs the job",
          the_pad_maker_still_streams_the_engine_output)

    def the_pad_maker_polygons_and_scrap_mode_unchanged():
        feed_ = feed
        # Hand-drawn polygon: traced for framing, G92 at its anchor vertex.
        pad_reset()
        pad_form("12 x 4\n18.5 x 3", "felt")
        poly = [(12.0, 0.0), (180.0, 9.0), (171.0, 140.0), (95.0, 96.0), (0.0, 131.0)]
        app.custom_polygon = poly
        got = machine(app.on_frame_and_cut)
        assert not errors() and len(got) == 3, (errors(), len(got))
        jog, framing, cut = got
        ax, ay = app._frame_cut_anchor()
        g92 = f"G92 X{ax:.3f} Y{ay:.3f}"
        assert (ax, ay) != (0.0, 0.0), "fixture: the anchor should not be the bbox corner"
        assert framing["lines"] == [g92] + ge.generate_polygon_framing_gcode(
            poly, power_s=FRAMING_S["felt"], feed=feed_)
        placed = previews[0]["placements"]["felt"]
        assert cut["lines"] == [g92] + engine_pad_cut(placed, "felt", 200.0, 150.0, polygon=poly, flip=140.0)

        # Camera outline with an inset: frame the outline, flip the cut with its height.
        pad_reset()
        pad_form("12 x 4\n18.5 x 3", "leather")
        outline = [(8.0, 0.0), (190.0, 14.0), (182.0, 150.0), (0.0, 138.0)]
        inset = [(13.0, 5.5), (184.5, 18.7), (177.3, 144.6), (5.4, 133.2)]
        app.custom_polygon_outline, app.custom_polygon = outline, inset
        app._custom_polygon_lb_machine = (123.4, 56.7)
        # A camera outline comes from a calibrated camera; Auto Locate is one of its extras.
        keep_cal = app._camera_calibration_present
        app._camera_calibration_present = lambda: True
        try:
            got = machine(app.on_frame_and_cut)
        finally:
            app._camera_calibration_present = keep_cal
        assert not errors() and len(got) == 3, (errors(), len(got))
        jog, framing, cut = got
        ax, ay = app._frame_cut_anchor()
        g92 = f"G92 X{ax:.3f} Y{ay:.3f}"
        assert jog["kw"].get("auto_locate_target") == (123.4, 56.7)
        assert framing["lines"] == [g92] + ge.generate_polygon_framing_gcode(
            outline, power_s=FRAMING_S["leather"], feed=feed_)
        placed = previews[0]["placements"]["leather"]
        run_log = list(log_lines)
        assert cut["lines"] == [g92] + engine_pad_cut(placed, "leather", 200.0, 150.0, polygon=inset, flip=150.0),             "the camera-outline cut is not the engine's output flipped with the outline's height"
        assert cut["lines"] != [g92] + engine_pad_cut(placed, "leather", 200.0, 150.0, polygon=inset),             "fixture: the inset flip should differ from the outline flip"
        assert len(run_log) == 1 and "flip 150.0" in run_log[0], run_log

        # Scrap mode: a stopped cut is not counted, a finished one is.
        pad_reset()
        pad_form("12.5 x 30\n20.0 x 6", "felt", w=70, h=60)
        app.scrap_mode_var.set(True)
        try:
            machine(app.on_frame_and_cut, script=["complete", "complete", "stopped"])
            assert app.scrap_session["active"] and app.scrap_session["scrap_count"] == 0
            assert not app.scrap_session.get("pending")
            got = machine(app.on_frame_and_cut)
            assert len(got) == 3 and app.scrap_session["scrap_count"] == 1
            placed = previews[0]["placements"]["felt"]
            assert sum(app.scrap_session["done"].values()) == len(placed)
        finally:
            pad_reset()
            app._update_scrap_status_display()
    check("Pad Maker: polygon, camera outline and scrap-mode jobs are framed, anchored and counted as before",
          the_pad_maker_polygons_and_scrap_mode_unchanged)

    def the_pad_maker_runs_without_a_camera_calibration():
        # No calibration file exists in this profile. Frame & Cut used to be
        # hidden then; now it runs, and touches nothing of the camera's even
        # when a tick and a position are left over from an older calibration.
        assert not app._camera_calibration_present(), "fixture: this profile has no calibration"
        assert not os.path.exists(os.path.join(_PROFILE, "camera_calibration.json"))
        for label, polygon in (("standard sheet", None),
                               ("hand-drawn shape", [(12.0, 0.0), (180.0, 9.0), (171.0, 140.0),
                                                     (95.0, 96.0), (0.0, 131.0)])):
            pad_reset()
            pad_form("12 x 4\n18.5 x 3", "felt")
            app.custom_polygon = polygon
            app._custom_polygon_lb_machine = (123.4, 56.7)      # stale
            app.live_camera_var.set(True)                       # stale tick
            try:
                got = machine(app.on_frame_and_cut)
                ran_log = list(log_lines)
            finally:
                app.live_camera_var.set(False)
            assert not errors(), (label, errors())
            assert len(got) == 3, f"{label}: jog, framing and cut expected, got {len(got)}"
            jog, framing, cut = got
            assert camera_asks == [], f"{label}: a camera was opened with no calibration"
            assert jog["kw"].get("auto_locate_target") is None, f"{label}: Auto Locate offered with no calibration"
            assert jog["kw"].get("remind_lid_on_advance") is True
            placed = previews[0]["placements"]["felt"]
            ax, ay = app._frame_cut_anchor()
            g92 = f"G92 X{ax:.3f} Y{ay:.3f}"
            flip = max(y for _x, y in polygon) if polygon else None
            assert cut["lines"] == [g92] + engine_pad_cut(placed, "felt", 200.0, 150.0, polygon=polygon, flip=flip), \
                f"{label}: the cut is not the engine's output"
            if polygon:
                assert framing["lines"] == [g92] + ge.generate_polygon_framing_gcode(
                    polygon, power_s=FRAMING_S["felt"], feed=feed)
            assert len(ran_log) == 1
        pad_reset()
    check("Pad Maker: with NO camera calibration a full run works (sheet and hand-drawn shape), camera untouched",
          the_pad_maker_runs_without_a_camera_calibration)

    def with_a_calibration_the_camera_extras_come_back():
        pad_reset()
        pad_form("12 x 4\n18.5 x 3", "felt")
        # A machine position only ever exists together with the camera
        # shape it belongs to (every unload clears both), and a job that is
        # not on a shape is offered no Auto Locate target.
        app.custom_polygon_outline = [(8.0, 0.0), (190.0, 14.0), (182.0, 150.0), (0.0, 138.0)]
        app.custom_polygon = [(13.0, 5.5), (184.5, 18.7), (177.3, 144.6), (5.4, 133.2)]
        app._custom_polygon_lb_machine = (123.4, 56.7)
        app.live_camera_var.set(True)
        keep_cal = app._camera_calibration_present
        app._camera_calibration_present = lambda: True
        try:
            got = machine(app.on_frame_and_cut)
            asks = list(camera_asks)
        finally:
            app._camera_calibration_present = keep_cal
            pad_reset()
        assert got[0]["kw"].get("auto_locate_target") == (123.4, 56.7) and asks == [1], (got[0]["kw"], asks)
    check("Pad Maker: with a calibration the live camera tick and Try Auto Locate work as before",
          with_a_calibration_the_camera_extras_come_back)

    for tool in TOOLS:
        def streams_what_the_export_writes(tool=tool):
            want = exported(tool)
            files_before = set(os.listdir(out))
            dialogs_before = len(saved)
            hist = len(pad_job_history())
            got = machine(tool["click"])
            assert not errors(), errors()
            assert len(got) == 3, f"jog, framing, cut expected; got {[r['ends'] for r in got]}"
            jog, framing, cut = got
            assert jog["lines"] == [], "the jog dialog streams nothing"
            assert cut["lines"][0] == G92_ZERO, cut["lines"][0]
            assert cut["lines"][1:] == want, "Frame & Cut streamed something other than the exported file"
            assert cut["kw"].get("title") == f"Cutting — {tool['material']}", cut["kw"].get("title")
            # Kept for the audit, one log line, and no Pad Maker job history.
            assert read_lines(kept_path) == cut["lines"]
            assert len(log_lines) == 1, log_lines
            assert log_lines[0].startswith(f"frame&cut tooling {tool['name']} ({tool['material']}) scrap -: "), log_lines
            assert " pieces, " in log_lines[0] and "G92 X0.000 Y0.000" in log_lines[0], log_lines
            assert len(pad_job_history()) == hist, "Tooling runs stay out of Job History"
            # Nothing in the user's folders but the Speed & Power legend.
            new_files = set(os.listdir(out)) - files_before
            new_dialogs = saved[dialogs_before:]
            if tool["legend"]:
                assert new_dialogs == [os.path.join(out, "sp_leather_legend.txt")], new_dialogs
                assert new_files <= {"sp_leather_legend.txt"}, new_files
            else:
                assert new_dialogs == [] and new_files == set(), (new_dialogs, new_files)
            # The G-code went through a temp file that is gone again.
            assert temp_files and not [p for p in temp_files if os.path.exists(p)], temp_files
            # Around it: the sleep lock, one connection, closed at the end.
            assert events[0] == "sleep:prevent" and events[-1] == "sleep:allow", events
            assert events.count("sender:connect") == 1 and events.count("sender:disconnect") == 1, events
            assert events.index("sender:unlock") < events.index("run:complete") < events.index("sender:disconnect")
        check(f"{tool['name']}: the cut is the exported G-code behind G92 X0 Y0, kept and logged, no file written",
              streams_what_the_export_writes)

        def frames_the_bounding_box_at_the_materials_power(tool=tool):
            want = exported(tool)
            got = machine(tool["click"])
            _jog, framing, _cut = got
            power = FRAMING_S[tool["material"]]
            assert framing["kw"].get("loop") is True, "framing runs as a loop until Cut"
            x0, y0, x1, y1 = own_extent(want)
            assert framing["lines"] == [G92_ZERO] + ge.generate_framing_gcode(
                x0, y0, x1, y1, power_s=power, feed=feed), "framing is not the job's bounding box"
            fx0, fy0, fx1, fy1 = own_extent(framing["lines"][1:])
            assert max(abs(fx0 - x0), abs(fy0 - y0), abs(fx1 - x1), abs(fy1 - y1)) < 0.01, \
                f"the traced box {(fx0, fy0, fx1, fy1)} is not the cut's extent {(x0, y0, x1, y1)}"
            laser = [ln for ln in framing["lines"] if re.match(r"M[34]\b", ln)]
            assert laser and all(ln == f"M4 S{power}" for ln in laser), \
                f"framing must be M4 at {tool['material']}'s power S{power}: {laser}"
            assert not [ln for ln in framing["lines"] if ln.startswith("M3")], "framing must never use M3"
        check(f"{tool['name']}: framing is the G-code's bounding box at {tool['material']}'s framing power, M4 not M3",
              frames_the_bounding_box_at_the_materials_power)

        def the_lid_question_is_asked_once_at_start_frame(tool=tool):
            tool["setup"]()
            got = machine(tool["click"])
            asks = [bool(r["kw"].get("remind_lid_on_advance")) for r in got]
            assert asks == [True, False, False], f"only the jog dialog (Start Frame) asks about the lid: {asks}"
            jog = got[0]["kw"]
            assert pad_jog, "the Pad Maker run above recorded its jog dialog"
            assert jog == pad_jog, \
                f"the jog dialog differs from the Pad Maker's: {sorted(set(jog.items()) ^ set(pad_jog.items()))}"
            assert jog["done_button_label"] == "Start Frame →" and "bottom-left corner" in jog["title"]
            assert jog["auto_locate_target"] is None and jog["stop_needs_confirm"] is False
        check(f"{tool['name']}: one lid question, on Start Frame, in the Pad Maker's jog dialog",
              the_lid_question_is_asked_once_at_start_frame)

        def cancelling_consumes_nothing(tool=tool):
            tool["setup"]()
            answer["save"] = "auto"
            hist = len(pad_job_history())
            for script, n_dialogs, cut_streamed in ((["cancelled"], 1, False),
                                                    (["complete", "stopped"], 2, False),
                                                    (["complete", "complete", "stopped"], 3, True),
                                                    (["complete", "complete", "error"], 3, True)):
                kept_before = kept_state()
                files_before = set(os.listdir(out))
                got = machine(tool["click"], script=script)
                assert not errors(), errors()
                assert len(got) == n_dialogs, f"{script}: {len(got)} dialogs"
                assert events[-1] == "sleep:allow", events
                if cut_streamed:
                    assert read_lines(kept_path) == got[2]["lines"], "a cut that was streamed is kept, stopped or not"
                else:
                    assert kept_state() == kept_before and not log_lines, \
                        f"{script}: nothing was cut, so nothing may be kept or logged"
                new_files = set(os.listdir(out)) - files_before
                assert new_files <= ({"sp_leather_legend.txt"} if tool["legend"] else set()), new_files
                assert not kerf_shown, f"{script}: the kerf measuring steps belong after a finished cut"
                assert not [b for b in boxes if b[0] == "showinfo"], boxes
                assert not app.tooling_scrap_session["active"]
            assert len(pad_job_history()) == hist
        check(f"{tool['name']}: cancel at the jog step, stop while framing, stop or error in the cut: nothing consumed",
              cancelling_consumes_nothing)

    # ---- kerf test: the measuring steps ------------------------------------
    def kerf_steps_follow_a_finished_cut():
        setup_kerf()
        machine(app._on_frame_cut_kerf)
        assert kerf_shown == ["Felt"], kerf_shown
        setup_kerf()
        app.kerf_use_existing_var.set(False)
        app.kerf_speed_var.set("450")
        app.kerf_power_var.set("55")
        got = machine(app._on_frame_cut_kerf)
        text = "\n".join(got[2]["lines"])
        assert "F450" in text and "S550" in text, "the custom cut speed and power did not reach the laser"
        app.kerf_use_existing_var.set(True)
    check("kerf test: the measuring steps show after a finished cut; custom speed and power are what is streamed",
          kerf_steps_follow_a_finished_cut)

    # ---- speed & power: the legend and the preview ---------------------------
    def the_legend_is_written_before_anything_moves():
        want = exported(TOOLS[3])
        export_legend = read_lines(os.path.splitext(saved[-1])[0] + "_legend.txt")
        legend = os.path.join(out, "sp_leather_legend.txt")
        if os.path.exists(legend):
            os.remove(legend)
        got = machine(app._on_frame_cut_feeds_speeds)
        assert len(got) == 3 and got[2]["lines"][1:] == want
        assert "sp_leather_legend.txt" in got[0]["files_then"], \
            "the legend must be on disk before the jog dialog opens"
        assert events.index("save-dialog") < events.index("sender:new"), events
        mine = read_lines(legend)
        assert mine[0].endswith("sp_leather_legend.txt") and mine[0].startswith("Speed & Power Test"), mine[0]
        assert mine[1:] == export_legend[1:], "the legend is not the one the file export writes"
        assert any(ln.startswith("01 ") for ln in mine), "no disc rows in the legend"
    check("speed & power: the legend the export writes is saved where the user chose, before the head moves",
          the_legend_is_written_before_anything_moves)

    def backing_out_of_the_legend_backs_out_of_the_run():
        setup_fs()
        files_before = set(os.listdir(out))
        answer["save"] = None
        try:
            got = machine(app._on_frame_cut_feeds_speeds)
        finally:
            answer["save"] = "auto"
        assert got == [] and "sender:new" not in events, events
        assert set(os.listdir(out)) == files_before and not errors() and not log_lines
        assert events == ["sleep:prevent", "save-dialog", "sleep:allow"], events
    check("speed & power: cancelling the legend prompt runs nothing and writes nothing",
          backing_out_of_the_legend_backs_out_of_the_run)

    def the_legend_goes_exactly_where_the_user_said():
        setup_fs()
        for stray in ("my notes.txt", "my notes_legend.txt"):
            if os.path.exists(os.path.join(out, stray)):
                os.remove(os.path.join(out, stray))
        answer["save"] = "my notes.txt"
        try:
            got = machine(app._on_frame_cut_feeds_speeds)
        finally:
            answer["save"] = "auto"
        assert len(got) == 3
        assert os.path.isfile(os.path.join(out, "my notes.txt")), os.listdir(out)
        assert not os.path.exists(os.path.join(out, "my notes_legend.txt")), "the legend was renamed behind the user's back"
    check("speed & power: a legend name the user typed is the name that is written",
          the_legend_goes_exactly_where_the_user_said)

    def the_preview_comes_first_and_adjust_backs_out():
        setup_fs()
        app.fs_show_preview_var.set(True)
        try:
            got = machine(app._on_frame_cut_feeds_speeds)
            assert len(got) == 3 and len(previews) == 1
            assert previews[0].get("proceed_label") == "Continue →", previews[0]
            assert events.index("preview") < events.index("save-dialog") < events.index("sender:new"), events
            answer["preview"] = None
            got = machine(app._on_frame_cut_feeds_speeds)
            assert got == [] and events == ["sleep:prevent", "preview", "sleep:allow"], events
        finally:
            answer["preview"] = "save"
            app.fs_show_preview_var.set(False)
    check("speed & power: with 'Preview before saving' ticked the preview comes first; Adjust runs nothing",
          the_preview_comes_first_and_adjust_backs_out)

    def the_real_preview_says_continue():
        # The window blocks in wait_window() inside its constructor, so the
        # look is scheduled first and runs inside that wait; whatever
        # happens, the window is closed or the constructor never returns.
        seen = {}
        pieces = [{"id": "01", "cx": 20.0, "cy": 20.0, "diameter": 20.0, "inner_diameter": 0,
                   "speed": 300, "power": 50, "passes": 1, "air_assist": True}]

        def look(label):
            def go():
                wins = [c for c in root.winfo_children() if isinstance(c, ui_dialogs.SpeedPowerTestPreview)]
                try:
                    seen[label] = wins[0]._proceed_btn.cget("text")
                except Exception as e:
                    seen[label] = repr(e)
                finally:
                    for w in wins:
                        if w.winfo_exists():
                            w._on_adjust()
            root.after(0, go)
        look("machine")
        ui_dialogs.SpeedPowerTestPreview(root, pieces, 60, 60, eng_speed=1000, eng_power=10,
                                         material_display="Felt", proceed_label="Continue →")
        look("file")
        ui_dialogs.SpeedPowerTestPreview(root, pieces, 60, 60, eng_speed=1000, eng_power=10,
                                         material_display="Felt")
        assert seen == {"machine": "Continue →", "file": "Save G-code"}, seen
    check("speed & power: the real preview's button reads Continue for the laser and Save G-code for a file",
          the_real_preview_says_continue)

    # ---- die inserts: scrap mode ----------------------------------------------
    def die_sizes_in(lines_or_placed):
        return sorted(p[0] for p in lines_or_placed)

    def engine_die_cut(placed, w, h):
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "want.gcode")
            ge.generate_die_gcode_from_placed(placed, w, h, path, app._get_die_settings())
            return read_lines(path)

    def die_scrap_one_sheet_per_click_counted_when_cut():
        setup_die()
        app.die_size_entry.delete("1.0", tk.END)
        # Ten small dies (50 mm rings) and two large (70 mm). Nested from the
        # bottom-left corner a 130 x 130 mm sheet takes four small ones, so
        # two sheets leave two small and two large for the file exports.
        # (Seven dies were enough while the nest started from the middle of
        # the sheet and took fewer per sheet.)
        sizes = (20.0, 22.0, 24.0, 26.0, 28.0, 30.0, 32.0, 34.0, 36.0, 38.0, 42.0, 44.0)
        app.die_size_entry.insert("1.0", ", ".join(f"{s:g}" for s in sizes))
        app.die_width_var.set("130")
        app.die_height_var.set("130")
        app.die_scrap_var.set(True)
        del folders[:]
        cwd_before = set(os.listdir(os.getcwd()))
        try:
            total = len(sizes)
            sess = lambda: app.tooling_scrap_session   # noqa: E731

            def untouched(why):
                assert sess()["active"], why
                assert sess()["scrap_count"] == 0, f"{why}: the sheet was counted"
                assert app._count_die_remaining() == total, f"{why}: dies came off Remaining"
                assert not [b for b in boxes if b[0] == "showinfo"], f"{why}: {boxes}"

            # Nothing that stops short of a finished cut is counted.
            got = machine(app._on_frame_cut_die, script=["cancelled"])
            assert len(got) == 1
            untouched("cancelled at the jog step")
            if app.tooling_scrap_window is not None:
                app.tooling_scrap_window.withdraw()
            got = machine(app._on_frame_cut_die, script=["complete", "stopped"])
            assert len(got) == 2
            untouched("framing stopped")
            got = machine(app._on_frame_cut_die, script=["complete", "complete", "stopped"])
            assert len(got) == 3
            untouched("cut stopped")
            first_try = got[2]["lines"]
            got = machine(app._on_frame_cut_die, script=["complete", "complete", "error"])
            untouched("cut errored")
            got = machine(app._on_frame_cut_die, sender=DeadSender)
            assert got == [] and [b[1] for b in errors()] == ["Connection Failed"], boxes
            del boxes[:]
            untouched("the laser would not connect")
            assert folders == [], f"Frame & Cut writes no files and asks for no folder: {folders}"

            # A finished cut counts its sheet: the same dies the stopped try offered.
            got = machine(app._on_frame_cut_die)
            assert len(got) == 3 and not errors(), errors()
            cut1 = got[2]["lines"]
            assert cut1 == first_try, "the sheet a stopped cut left behind should be offered again unchanged"
            pads = [{"size": s, "qty": 1} for s in sizes]
            placed1, remaining1, _any = se.try_nest_partial(pads, "die_ring", 130.0, 130.0, app._get_die_settings())
            assert len(placed1) == 4, f"fixture: the first sheet takes four 50 mm dies from the corner, took {len(placed1)}"
            assert cut1 == [G92_ZERO] + engine_die_cut(placed1, 130.0, 130.0)
            assert sess()["scrap_count"] == 1 and app._count_die_remaining() == total - len(placed1)
            assert log_lines[0].startswith("frame&cut tooling die inserts (acrylic) scrap 1: "), log_lines
            info = [b for b in boxes if b[0] == "showinfo"]
            assert len(info) == 1 and info[0][1] == "Sheet Cut" and f"Cut {len(placed1)} dies on sheet #1" in info[0][2], info
            assert f"{len(placed1)} / {total} dies complete" == app.die_scrap_header.cget("text")

            # The next click takes what is left, not the list from the top.
            got = machine(app._on_frame_cut_die)
            placed2, remaining2, _any = se.try_nest_partial(remaining1, "die_ring", 130.0, 130.0, app._get_die_settings())
            assert got[2]["lines"] == [G92_ZERO] + engine_die_cut(placed2, 130.0, 130.0)
            assert not set(die_sizes_in(placed1)) & set(die_sizes_in(placed2)), "a die was cut twice"
            assert sess()["scrap_count"] == 2
            assert app._count_die_remaining() == total - len(placed1) - len(placed2)
            assert log_lines[0].startswith("frame&cut tooling die inserts (acrylic) scrap 2: "), log_lines

            assert app._count_die_remaining() > 0, "fixture: dies should be left for the file export below"
            # A file export joining a session Frame & Cut started has no
            # folder yet: it asks once, and writes there (not into the
            # working directory).
            del folders[:]
            app.die_width_var.set("60")
            app.die_height_var.set("60")           # one small die per sheet from here
            app._on_generate_die_svg()
            root.update()
            assert folders == ["Select Folder to Save SVGs"], folders
            assert sess()["save_dir"] == out
            assert os.path.isfile(os.path.join(out, "dies_scrap3.svg")), sorted(os.listdir(out))
            assert sess()["scrap_count"] == 3
            del folders[:]
            assert app._count_die_remaining() > 0, "fixture: dies should be left for a second file export"
            app._on_generate_die_gcode()
            root.update()
            assert folders == [], "the folder is asked for once per session"
            assert os.path.isfile(os.path.join(out, "dies_scrap4.gcode")), sorted(os.listdir(out))
            stray = [f for f in set(os.listdir(os.getcwd())) - cwd_before if "scrap" in f]
            assert not stray, f"a scrap file was written into the working directory: {stray}"
        finally:
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            app._update_die_scrap_status()
    check("die inserts, Scrap Mode: one sheet per click, counted only after a finished cut; no folder until a file export",
          die_scrap_one_sheet_per_click_counted_when_cut)

    def die_scrap_last_sheet_ends_the_session():
        setup_die()
        app.die_size_entry.delete("1.0", tk.END)
        app.die_size_entry.insert("1.0", "30, 32")
        app.die_scrap_var.set(True)
        try:
            got = machine(app._on_frame_cut_die)
            if app.tooling_scrap_window is not None:
                app.tooling_scrap_window.withdraw()
            assert len(got) == 3 and app._count_die_remaining() == 0
            info = [b for b in boxes if b[0] == "showinfo"]
            assert len(info) == 1 and info[0][1] == "Session Complete!" and "Cut 2 dies on sheet #1" in info[0][2], info
            got = machine(app._on_frame_cut_die)
            assert got == [] and "sender:new" not in events
            assert [b[1] for b in boxes if b[0] == "showinfo"] == ["Session Complete"], boxes
        finally:
            app._clear_die_scrap_session()
            app.die_scrap_var.set(False)
            app._update_die_scrap_status()
    check("die inserts, Scrap Mode: the last sheet ends the session and a further click cuts nothing",
          die_scrap_last_sheet_ends_the_session)

    def die_scrap_file_exports_still_ask_first_and_count_on_writing():
        # The two file exports share their front half with Frame & Cut now.
        # For them nothing changed: the folder is asked for when the session
        # starts, and a sheet is counted as soon as its file is written.
        for handler, title, ext in ((app._on_generate_die_gcode, "Select Folder to Save G-code", ".gcode"),
                                    (app._on_generate_die_svg, "Select Folder to Save SVGs", ".svg")):
            setup_die()
            app.die_size_entry.delete("1.0", tk.END)
            app.die_size_entry.insert("1.0", "30, 32, 34")
            app.die_width_var.set("60")
            app.die_height_var.set("60")
            app.die_filename_var.set("filed")
            app.die_scrap_var.set(True)
            for stale in ("filed_scrap1" + ext, "filed_scrap2" + ext):
                if os.path.exists(os.path.join(out, stale)):
                    os.remove(os.path.join(out, stale))
            try:
                del folders[:], boxes[:]
                answer["folder"] = ""
                handler()
                root.update()
                assert folders == [title] and not app.tooling_scrap_session["active"], \
                    "backing out of the folder prompt starts no session"
                answer["folder"] = out
                del folders[:]
                handler()
                root.update()
                if app.tooling_scrap_window is not None:
                    app.tooling_scrap_window.withdraw()
                assert folders == [title], folders
                s = app.tooling_scrap_session
                assert s["active"] and s["save_dir"] == out and s["scrap_count"] == 1
                assert app._count_die_remaining() == 2 and os.path.isfile(os.path.join(out, "filed_scrap1" + ext))
                handler()
                root.update()
                assert folders == [title] and app.tooling_scrap_session["scrap_count"] == 2
                assert os.path.isfile(os.path.join(out, "filed_scrap2" + ext))
                assert not errors(), errors()
            finally:
                answer["folder"] = out
                app._clear_die_scrap_session()
                app.die_scrap_var.set(False)
                app._update_die_scrap_status()
    check("die inserts, Scrap Mode: the file exports still ask for a folder first and count a sheet when it is written",
          die_scrap_file_exports_still_ask_first_and_count_on_writing)

    # ---- the same refusals as Generate G-code ------------------------------------
    def refusals_match_generate_gcode():
        def both(tool, spoil, mend):
            tool["setup"]()
            spoil()
            try:
                del boxes[:]
                before = len(saved)
                tool["export"]()
                root.update()
                from_export = list(boxes)
                assert len(saved) == before, f"{tool['name']}: a refused form reached the save dialog"
                got = machine(tool["click"])
                from_machine = list(boxes)
            finally:
                mend()
            assert from_export and from_export[0][0] == "showerror", f"{tool['name']}: fixture not refused: {from_export}"
            assert from_machine == from_export, f"{tool['name']}: {from_machine} != {from_export}"
            assert got == [] and "sender:new" not in events, f"{tool['name']}: a refused form reached the laser"
            assert events[-1] == "sleep:allow"

        def die_empty():
            app.die_size_entry.delete("1.0", tk.END)
        both(TOOLS[0], die_empty, setup_die)
        both(TOOLS[0], lambda: (app.die_width_var.set("60"), app.die_height_var.set("60")), setup_die)
        both(TOOLS[0], lambda: app.die_width_var.set("wide"), setup_die)
        both(TOOLS[1], lambda: (app.holder_variant_var.set("both"), app.holder_width_var.set("200")), setup_holder)

        def kerf_bad():
            app.kerf_use_existing_var.set(False)
            app.kerf_speed_var.set("fast")

        def kerf_mend():
            app.kerf_speed_var.set("180")
            app.kerf_use_existing_var.set(True)
        both(TOOLS[2], kerf_bad, kerf_mend)
        both(TOOLS[3], lambda: app.fs_diameter_var.set("0"), setup_fs)
        both(TOOLS[3], lambda: app.fs_inner_diameter_var.set("25"), setup_fs)
    check("every tool refuses a bad form in the same words as Generate G-code, and nothing reaches the laser",
          refusals_match_generate_gcode)

    # ---- the Pad Maker's state stays out of it --------------------------------------
    def a_shape_loaded_on_the_pad_maker_cannot_bend_a_tooling_job():
        want = exported(TOOLS[2])
        app.custom_polygon_outline = [(8.0, 0.0), (190.0, 14.0), (182.0, 150.0), (0.0, 138.0)]
        app.custom_polygon = [(13.0, 5.5), (184.5, 18.7), (177.3, 144.6), (5.4, 133.2)]
        app._custom_polygon_lb_machine = (123.4, 56.7)
        try:
            assert app._frame_cut_anchor() != (0.0, 0.0), "fixture: the pad shape's anchor is off the corner"
            got = machine(app._on_frame_cut_kerf)
        finally:
            pad_reset()
        jog, framing, cut = got
        assert cut["lines"] == [G92_ZERO] + want, "the pad shape's anchor moved a Tooling cut"
        x0, y0, x1, y1 = own_extent(want)
        assert framing["lines"] == [G92_ZERO] + ge.generate_framing_gcode(
            x0, y0, x1, y1, power_s=FRAMING_S["felt"], feed=feed), "framing traced the pad shape"
        assert jog["kw"].get("auto_locate_target") is None, "Try Auto Locate aims at the pad shape"
        assert "flip sheet" in log_lines[0] and "polygon bbox none" in log_lines[0], log_lines
    check("a scrap shape loaded on the Pad Maker changes nothing about a Tooling job (origin, frame, auto locate)",
          a_shape_loaded_on_the_pad_maker_cannot_bend_a_tooling_job)

    def live_camera_only_with_a_calibration():
        setup_kerf()
        keep = app._camera_calibration_present
        app.live_camera_var.set(True)
        try:
            app._camera_calibration_present = lambda: False
            machine(app._on_frame_cut_kerf)
            assert camera_asks == [], "a hidden tick opened the camera with no calibration on file"
            app._camera_calibration_present = lambda: True
            machine(app._on_frame_cut_kerf)
            assert camera_asks == [1], "with a calibration, the Pad Maker's live-camera tick applies here too"
        finally:
            app._camera_calibration_present = keep
            app.live_camera_var.set(False)
    check("the Pad Maker's 'Live camera preview' tick applies to Tooling only while a camera calibration exists",
          live_camera_only_with_a_calibration)

    def no_laser_no_run():
        setup_kerf()
        app._detect_falcon_async = lambda: events.append("detect")
        keep = ui_dialogs.FalconRunDialog, falcon_sender.FalconSender
        ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = FakeRun, FakeSender
        del events[:], runs[:], boxes[:]
        app.falcon_port = None
        try:
            app._on_frame_cut_kerf()
        finally:
            ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = keep
        assert runs == [] and "sender:new" not in events
        assert [b[1] for b in errors()] == ["Falcon Not Detected"], boxes
    check("with no laser detected a Tooling click says so and runs nothing", no_laser_no_run)

    # ---- the buttons ---------------------------------------------------------------------
    def shown():
        return [bool(b.winfo_manager()) for b, _opts in app._tooling_frame_cut_btns]

    LASER_ITEMS = ("home", "test", "clear", "reset", "framing_power")

    def menu_states():
        return {key: str(app._machine_menu.entrycget(index, "state"))
                for key, index in app._machine_menu_indices.items()}

    def gate(toggle, port, calibration, pyserial=True):
        """(Tooling buttons shown, Pad Maker button shown, live camera tick shown)."""
        keep = app._camera_calibration_present, main_mod.machine_supported
        app.settings["experimental_machine_menu"] = toggle
        app.falcon_port = port
        app._camera_calibration_present = lambda: calibration
        main_mod.machine_supported = lambda: pyserial
        try:
            app._refresh_machine_ui_state()
            root.update_idletasks()
            return (shown(), bool(app._frame_cut_btn.winfo_manager()),
                    bool(app._live_camera_chk.winfo_manager()))
        finally:
            app._camera_calibration_present, main_mod.machine_supported = keep

    def rest():
        app.settings["experimental_machine_menu"] = True
        app.falcon_port = None
        app._refresh_machine_ui_state()

    N = len(TOOLS)

    def one_rule_for_the_button_on_both_tabs():
        try:
            assert len(app._tooling_frame_cut_btns) == N
            hidden = ([False] * N, False)
            assert gate(False, None, False)[:2] == hidden, "machine integration off"
            assert gate(False, "COM7", True)[:2] == hidden, "machine integration off, laser and calibration there"
            assert gate(True, None, False)[:2] == hidden, "no laser detected"
            assert gate(True, None, True)[:2] == hidden, "no laser detected, calibration there"
            assert gate(True, "COM7", True, pyserial=False)[:2] == hidden, "no pyserial in this build"
            # No calibration needed, on either tab.
            assert gate(True, "COM7", False)[:2] == ([True] * N, True), \
                "toggle on + laser, no calibration: Frame & Cut shows on the Pad Maker and on Tooling"
            assert gate(True, "COM7", True)[:2] == ([True] * N, True), "with a calibration too"
            # And they go away again, live.
            assert gate(True, None, False)[:2] == hidden
            assert gate(True, "COM7", False)[:2] == ([True] * N, True)
            assert gate(False, "COM7", False)[:2] == hidden
            # The live camera tick is the camera's: it still needs the calibration.
            assert gate(True, "COM7", False)[2] is False and gate(True, None, True)[2] is True
            assert gate(False, "COM7", True)[2] is False
        finally:
            rest()
    check("buttons: Frame & Cut on both tabs needs the toggle and a laser, no calibration; the camera tick still does",
          one_rule_for_the_button_on_both_tabs)

    def machine_menu_laser_items_need_only_the_toggle():
        try:
            assert set(app._machine_menu_indices) == set(LASER_ITEMS) | {"camera_cal", "inset"}, \
                sorted(app._machine_menu_indices)
            assert set(main_mod.MACHINE_MENU_LASER_ITEMS) == set(LASER_ITEMS)
            assert tuple(main_mod.MACHINE_MENU_CAMERA_ITEMS) == ("inset",)
            for port in (None, "COM7"):
                gate(True, port, False)
                st = menu_states()
                for key in LASER_ITEMS:
                    assert st[key] == "normal", f"{key} is greyed without a calibration (laser {port})"
                assert st["camera_cal"] == "normal", "Camera Calibration is always available"
                assert st["inset"] == "disabled", "the camera-polygon inset means nothing without a calibration"
                gate(True, port, True)
                st = menu_states()
                assert all(v == "normal" for v in st.values()), st
            # Back to no calibration: the inset greys out again, nothing else does.
            gate(True, "COM7", False)
            st = menu_states()
            assert st["inset"] == "disabled" and all(st[k] == "normal" for k in LASER_ITEMS + ("camera_cal",)), st
        finally:
            rest()
    check("Machine menu: Home, Test, Clear, Reset and Framing Power need only the toggle; the inset needs a calibration",
          machine_menu_laser_items_need_only_the_toggle)

    def camera_features_stay_behind_the_calibration():
        keep = app._camera_calibration_present
        try:
            app._camera_calibration_present = lambda: False
            assert app._camera_capture_ready() is False, "Get from camera, the overlay, Re-capture: need a calibration"
            app._camera_calibration_present = lambda: True
            assert app._camera_capture_ready() is True
            app.settings["experimental_machine_menu"] = False
            assert app._camera_capture_ready() is False, "and the toggle"
        finally:
            app._camera_calibration_present = keep
            rest()
    check("camera features (Get from camera, overlay, Re-capture) still need the toggle and a calibration",
          camera_features_stay_behind_the_calibration)

    GCODE_SECTIONS = {"die_inserts", "die_holders", "die_organizer", "kerf_test",
                      "feeds_speeds_tester", "tonehole_crescent"}

    def one_button_per_gcode_section_beside_generate_gcode():
        def texts(frame):
            found = []
            for w in frame.winfo_children():
                if isinstance(w, tk.Button):
                    found.append(str(w.cget("text")))
                found += texts(w)
            return found
        with_gcode = {k for k, f in app._tooling_sections.items() if "Generate G-code" in texts(f)}
        with_frame_cut = {k for k, f in app._tooling_sections.items() if "Frame & Cut" in texts(f)}
        assert with_gcode == GCODE_SECTIONS, with_gcode
        assert set(app._tooling_sections) - with_gcode == {"pad_press_spacers"}, \
            "only the 3D-printed spacers have no G-code"
        assert with_frame_cut == with_gcode, \
            f"every section that writes G-code needs Frame & Cut, and only those: {with_frame_cut ^ with_gcode}"
        gate(True, "COM7", False)
        try:
            for btn, _opts in app._tooling_frame_cut_btns:
                row = btn.master
                section = [k for k, f in app._tooling_sections.items() if str(btn).startswith(str(f) + ".")]
                assert len(section) == 1, section
                order = [str(w.cget("text")) for w in row.pack_slaves() if isinstance(w, (tk.Button, tk.Checkbutton))]
                at = order.index("Frame & Cut")
                assert at > 0 and order[at - 1] == "Generate G-code", f"{section[0]}: {order}"
                assert btn.tooltip.text.startswith("Send this job straight to the laser"), btn.tooltip.text
        finally:
            rest()
    check("buttons: one in each section that writes G-code, right after Generate G-code, with a tooltip",
          one_button_per_gcode_section_beside_generate_gcode)

    def a_long_translated_label_is_not_clipped():
        # A tk button's width is in characters and does not grow with its
        # text. "Frame & Cut" is 20 characters in German.
        real = builtins._
        builtins._ = lambda s: "Rahmen und schneiden" if s == "Frame & Cut" else s
        holder = tk.Frame(root)
        try:
            btn = app._add_tooling_frame_cut_button(holder, lambda: None, side="left")
        finally:
            builtins._ = real
            app._tooling_frame_cut_btns.pop()
        try:
            assert str(btn.cget("text")) == "Rahmen und schneiden"
            assert int(btn.cget("width")) >= len("Rahmen und schneiden"), btn.cget("width")
            for b, _opts in app._tooling_frame_cut_btns:
                assert int(b.cget("width")) == 15, "in English it matches the Generate buttons beside it"
        finally:
            holder.destroy()
    check("buttons: as wide as the Generate buttons, and never narrower than its label in a longer language",
          a_long_translated_label_is_not_clipped)

    def each_button_runs_its_own_tool():
        seen = {}
        for tool in TOOLS:
            tool["setup"]()
        answer["save"] = "auto"
        for btn, _opts in app._tooling_frame_cut_btns:
            section = [k for k, f in app._tooling_sections.items() if str(btn).startswith(str(f) + ".")][0]
            got = machine(btn.invoke)
            assert len(got) == 3 and not errors(), (section, errors())
            seen[section] = log_lines[0].split(" scrap ")[0]
        assert seen == {
            "die_inserts": "frame&cut tooling die inserts (acrylic)",
            "die_holders": "frame&cut tooling die holder (acrylic)",
            "die_organizer": "frame&cut tooling die organizer lower (basswood)",
            "kerf_test": "frame&cut tooling kerf test (felt)",
            "feeds_speeds_tester": "frame&cut tooling speed & power test (leather)",
            "tonehole_crescent": "frame&cut tooling tonehole crescent (cork)",
        }, seen
    check("buttons: clicking each one runs that section's own job", each_button_runs_its_own_tool)

    # ---- material names in a translated app ------------------------------------------------
    def material_key_survives_translation():
        key = app._tooling_material_key
        for english in ("Felt", "Card", "Leather", "Acrylic", "Basswood"):
            assert key(english) == english.lower()
        german = {"Felt": "Filz", "Card": "Karton", "Leather": "Leder", "Acrylic": "Acryl", "Basswood": "Lindenholz"}
        real = builtins._
        builtins._ = lambda s: german.get(s, s)
        try:
            for english, shown_as in german.items():
                assert key(shown_as) == english.lower(), f"{shown_as} -> {key(shown_as)}"
            assert key("Acrylic") == "acrylic", "the dropdown's untranslated default must still resolve"
            # The point of it: the user's own felt settings reach the kerf test.
            app.settings["gcode_settings"]["felt"]["cut_speed"] = 777
            app.settings["gcode_settings"]["felt"]["cut_power"] = 33
            app.kerf_material_var.set("Filz")
            app.kerf_use_existing_var.set(True)
            job = app._kerf_gcode_job()
            assert (job["mat_key"], job["cut_speed"], job["cut_power"]) == ("felt", 777, 33), job
            app.fs_material_var.set("Leder")
            assert app._get_feeds_speeds_material_key() == "leather"
        finally:
            builtins._ = real
            setup_kerf()
            setup_fs()
        assert key("Mystery") == "mystery" and key("") == ""
    check("a material picked in a translated dropdown still finds its own settings and framing power",
          material_key_survives_translation)

    # ---- wiring --------------------------------------------------------------------------------
    def one_routine_for_both_tabs():
        sig = inspect.signature(main_mod.PadSVGGeneratorApp.on_frame_and_cut)
        assert list(sig.parameters) == ["self", "tooling_job"] and sig.parameters["tooling_job"].default is None
        src = inspect.getsource(tooling_tab.ToolingTabMixin)
        for word in ("FalconRunDialog", "FalconSender", "generate_framing_gcode", "G92"):
            assert word not in src, f"Tooling must go through on_frame_and_cut, not its own {word}"
        assert src.count("self.on_frame_and_cut(tooling_job=") == 7, "six tools, and Die Scrap Mode"
    check("one machine routine: Tooling hands its job to on_frame_and_cut and drives no dialog or sender itself",
          one_routine_for_both_tabs)

    main_mod.machine_supported = real_machine_supported
    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
