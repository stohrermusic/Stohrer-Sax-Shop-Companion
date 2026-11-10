"""The machine half of the app: the Machine menu, Frame & Cut, the laser watch and the camera plumbing.

MachineControlMixin is mixed into main.PadSVGGeneratorApp beside the tab mixins. Every method here moved
out of main.py on 2026-11-10 byte for byte (structural step 2); the three constants moved with their
comments and main re-imports them, so `main.MACHINE_MENU_LASER_ITEMS`, `main.MACHINE_MENU_CAMERA_ITEMS`
and `main.LASER_WATCH_PERIOD_MS` still work. Not a tab, hence not a *_tab module: it adds menu items,
buttons and background work to tabs that exist.

Two names the moved code reads are ones the suites and the recordings stand in for on `main`
(main.NestingPreviewWindow = Fake, main.machine_supported = lambda: True). They are reached through
_FromMain at the call, so the stand-in is what the moved code gets and the method bodies stayed
untouched; every other name is imported here as main imported it. History: docs/decisions/machine-integration.md.
"""
import logging
import os
import sys
import time
import tkinter as tk
from tkinter import messagebox

from threads import start_worker
from svg_engine import nest_with_zones
from gcode_engine import generate_gcode_from_placed


def _main_module():
    """The live main module: `main` when it was imported (the suites, the recorders), else the
    running `__main__` (the app itself; importing `main` there would run main.py a second time)."""
    mod = sys.modules.get("main")
    if mod is None or not hasattr(mod, "PadSVGGeneratorApp"):
        mod = sys.modules["__main__"]
    return mod


class _FromMain:
    """A name of the main module, read at the call: `NestingPreviewWindow(...)` and
    `machine_supported()` in the moved methods build or ask whatever main holds right now,
    a stand-in included."""

    def __init__(self, name):
        self._name = name

    def __call__(self, *args, **kwargs):
        return getattr(_main_module(), self._name)(*args, **kwargs)


NestingPreviewWindow = _FromMain("NestingPreviewWindow")
machine_supported = _FromMain("machine_supported")



# Options > Machine, by what each item needs (see _refresh_machine_ui_state).
# The laser items work with the machine-integration toggle alone; the camera
# item also needs a camera calibration on disk. Camera Calibration itself is
# always enabled.
MACHINE_MENU_LASER_ITEMS = ('framing_power', 'home', 'test', 'clear', 'reset')
MACHINE_MENU_CAMERA_ITEMS = ('inset',)
# How often the laser watch lists the serial ports (enumeration only; a port
# is opened only when LaserWatch says so). See _start_laser_watch.
LASER_WATCH_PERIOD_MS = 2000


class MachineControlMixin:
    """The Machine menu, Frame & Cut, the laser watch and the camera: see the module docstring."""

    def _release_live_camera(self):
        """Close the app's own live camera preview (LiveCameraWindow, left
        open after a Frame & Cut with "Live camera preview" ticked so the
        user can see the result) before a dialog opens a camera of its
        own. Two handles on one DirectShow device fail or starve each
        other, and with the fallback in open_camera_or_fallback the second
        open would otherwise move to another camera. The next Frame & Cut
        opens the preview again. Called before the shape dialog (its
        overlay and "Get from camera") and before Camera Calibration."""
        existing = getattr(self, '_live_cam_window', None)
        self._live_cam_window = None
        if existing is None:
            return
        try:
            if existing.winfo_exists():
                existing.close()
        except Exception:
            pass

    def _camera_capture_ready(self):
        """True iff the experimental toggle is on AND OpenCV is installed
        AND a camera calibration file exists. Toggle-off short-circuits
        so the 'Re-capture from camera' button in the scrap-continue
        dialog (and any other caller using this helper) hides when the
        user has opted out of machine integration or is on a platform
        we don't support."""
        return bool(self._machine_enabled() and self._camera_calibration_present())

    def _frame_cut_anchor(self):
        """The G92 X/Y for Frame & Cut: the outline's anchor vertex
        (svg_engine.polygon_anchor_vertex) in framing's Y-up frame,
        relative to the outline's bounding box. (0, 0) without a
        polygon. Try Auto Locate drives the head to the same vertex."""
        from svg_engine import polygon_anchor_vertex
        offset_source = (self.custom_polygon_outline
                         or self.custom_polygon)
        if not offset_source:
            return 0.0, 0.0
        y_max_storage = max(p[1] for p in offset_source)
        flipped = [(x, y_max_storage - y) for (x, y) in offset_source]
        x, y = polygon_anchor_vertex(flipped)
        return float(x), float(y)

    def _adopt_camera_polygon(self, machine_polygon):
        """Adopt an ABSOLUTE machine-Y-up polygon (from the camera's
        pixel→machine homography) as the active scrap shape.

        Stores BOTH:
          - self.custom_polygon: the INSET shape (safety-shrunk),
            used for pad nesting so cuts land safely inside the
            actual scrap edges.
          - self.custom_polygon_outline: the ORIGINAL un-insetted
            shape, used for Frame & Cut's framing trace so the trace
            visually matches the scrap edge.

        Both are normalized to the OUTLINE's bbox-BL so they share a
        coordinate origin — pad placements (in custom_polygon coords)
        land at the right machine positions during cuts even though
        the framing trace uses the outline.

        Also captures the polygon's LB-vertex MACHINE COORDINATE on
        self._custom_polygon_lb_machine for Frame & Cut's "Try Auto
        Locate" button.

        Used by:
          - File menu / standalone "Get from camera" flow
          - The polygon dialog's "Capture from camera" sub-action
          - The polygon dialog's overlay-traced path (via
            get_machine_polygon_mm)
        """
        if not machine_polygon:
            return
        # Capture the LB-vertex BEFORE inset shrinks the polygon — the
        # user is going to jog to the visible bottom-left corner of
        # their material, not the inset corner.
        from svg_engine import polygon_anchor_vertex
        self._custom_polygon_lb_machine = tuple(
            polygon_anchor_vertex(machine_polygon))
        outline = list(machine_polygon)
        inset_polygon = outline
        try:
            import camera_capture
            inset_mm = float(
                self.settings.get("camera_polygon_inset_mm", 3.0))
            if inset_mm > 0:
                try:
                    inset_polygon = camera_capture.inset_polygon_mm(
                        outline, inset_mm)
                except Exception as e:  # fall back to uninsetted on failure
                    logging.getLogger("ssc.machine").warning(
                        "camera shape: the %s mm safety inset failed, so pads will nest on the uninset outline: %s",
                        inset_mm, e)
        except ImportError as e:
            logging.getLogger("ssc.machine").warning(
                "camera shape: could not load camera_capture, so the %s mm safety inset was not applied: %s",
                self.settings.get("camera_polygon_inset_mm", 3.0), e)
        # Pass BOTH to the shared adopter so the inset polygon (for
        # nesting) and the outline (for framing) share an origin.
        self._set_custom_polygon_from_y_up(
            inset_polygon, outline_y_up_mm=outline)

    def _set_custom_polygon_from_y_up(self, points_y_up_mm,
                                        outline_y_up_mm=None):
        """Single boundary crossing into custom_polygon storage.

        Input:
          - points_y_up_mm: the polygon used for NESTING (the inset
            shape when there is one).
          - outline_y_up_mm (optional): the polygon used for FRAMING
            (the un-inset original). When None, frame and nest use
            the same shape — appropriate for plain hand-drawn
            polygons where no inset was applied.

        Both polygons are normalized to the OUTLINE's bbox-BL so they
        share a coordinate origin. The framing trace then matches the
        scrap edge while pad placements still land inside the inset
        boundary.

        Output:
          - self.custom_polygon: Y-DOWN, nesting boundary
          - self.custom_polygon_outline: Y-DOWN, framing trace
            (= self.custom_polygon when no separate outline)
        """
        if not points_y_up_mm:
            return
        outline = (outline_y_up_mm if outline_y_up_mm is not None
                    else points_y_up_mm)
        min_x = min(p[0] for p in outline)
        max_y = max(p[1] for p in outline)
        # Both polygons get the same translation + Y-flip so they
        # share an origin.
        self.custom_polygon = [(x - min_x, max_y - y)
                               for (x, y) in points_y_up_mm]
        self.custom_polygon_outline = [(x - min_x, max_y - y)
                                        for (x, y) in outline]
        self._update_shape_status()

    def _frame_cut_scrap_advance(self, placed_count, remaining):
        """Post-cut scrap bookkeeping for Frame & Cut, mirroring the tail
        of _generate_gcode_scrap_mode: bump the scrap count, record what's
        left for the next scrap, refresh the status UI, and either
        announce session completion or show the continue/recapture dialog.

        Called only after a cut has streamed to completion — if the user
        stopped the cut or it errored, the session is left untouched so
        the scrap can be re-cut.
        """
        self.scrap_session['scrap_count'] += 1
        scrap_num = self.scrap_session['scrap_count']
        self._scrap_record_result(remaining)
        self._update_scrap_status_display()
        self._update_remaining_pads_window()

        remaining_count = self._count_remaining_pads()
        if remaining_count == 0:
            messagebox.showinfo(_("Session Complete!"),
                                _("All pads have been placed!"))
        else:
            self._show_scrap_continue_dialog(
                placed_count, scrap_num, remaining_count)

    def _keep_last_cut_gcode(self, lines, material, scrap_num, g92, flip_height,
                             polygon, disc_count, job_label=None):
        """Keep a copy of the G-code Frame & Cut actually streamed, plus one
        log line describing its frame.

        The temp file the cut is generated into is deleted as soon as it's
        read, which left nothing to audit when a user reported labels
        landing outside their discs (2026-09-20). This is app.log for the
        machine: the exact lines the Falcon received, overwritten on every
        cut, at <config dir>/last_frame_cut.gcode. Swallows every error;
        bookkeeping must never stop a cut.

        job_label is None for a pad job, whose log line is unchanged, and
        the tool's name for a Tooling job ("tooling kerf test (felt)",
        counted in pieces).
        """
        import logging
        try:
            from config import get_config_dir
            path = os.path.join(get_config_dir(), "last_frame_cut.gcode")
            with open(path, "w") as f:
                f.write("\n".join(lines))
            if polygon:
                xs = [p[0] for p in polygon]
                ys = [p[1] for p in polygon]
                bbox = f"{max(xs) - min(xs):.1f}x{max(ys) - min(ys):.1f}"
            else:
                bbox = "none"
            if job_label:
                subject, unit = f"tooling {job_label} ({material})", "pieces"
            else:
                subject, unit = material, "discs"
            logging.warning(
                "frame&cut %s scrap %s: %d %s, %d lines, G92 X%.3f Y%.3f, "
                "flip %s, polygon bbox %s, kept at %s",
                subject, scrap_num if scrap_num else "-", disc_count, unit, len(lines),
                g92[0], g92[1],
                f"{flip_height:.1f}" if flip_height is not None else "sheet",
                bbox, path)
        except Exception:
            logging.exception("could not keep the last cut G-code")

    # ---- the laser watch -------------------------------------------------------
    # Matt, 2026-11-09: "is there a way for the ssc to be ok with a user
    # plugging in the laser when the ssc is already running? I feel like now
    # it has to be plugged in first." Until then the app probed every serial
    # port every 10 s until a laser answered, then never looked again: a
    # laser unplugged, power-cycled or moved to another socket left a stale
    # port (Frame & Cut failed to connect until a restart, the buttons stayed
    # up) and its homed flag, so Try Auto Locate stayed enabled after the
    # machine had lost its zero. Now falcon_sender.LaserWatch decides what to
    # probe from the port list every LASER_WATCH_PERIOD_MS, on a worker
    # thread (one at a time), and every state change happens here, on the Tk
    # thread. Started from the __main__ block only: building the app starts
    # no watch, no thread and no timer.

    def _start_laser_watch(self):
        """Begin watching the serial ports for the laser. Launch path only."""
        try:
            import falcon_sender
        except ImportError:
            return
        if not falcon_sender.HAS_PYSERIAL or not self._machine_enabled():
            return
        if getattr(self, '_laser_watch', None) is not None:
            return
        self._laser_watch = falcon_sender.LaserWatch()
        self._laser_watch_busy = False
        logging.getLogger("ssc.machine").info("laser watch started (every %d ms)", LASER_WATCH_PERIOD_MS)
        self._laser_watch_tick()

    def _laser_watch_tick(self):
        """One look at the ports, on a worker; then the next tick."""
        if getattr(self, '_laser_watch', None) is None:
            return
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        if not self._laser_watch_busy:
            self._laser_watch_busy = True
            start_worker(self._laser_watch_worker, name='laser-watch')
        self.root.after(LASER_WATCH_PERIOD_MS, self._laser_watch_tick)

    def _laser_watch_worker(self):
        """Worker thread: list, decide, probe; hand the events to the Tk
        thread. Must not raise, least of all after the window is gone."""
        import time as _time
        try:
            import falcon_sender
            events = self._laser_watch_pass(falcon_sender.list_serial_ports,
                                            falcon_sender.probe_grbl_port, _time.monotonic)
        except Exception as e:
            events = [('error', str(e))]
        finally:
            self._laser_watch_busy = False
        try:
            self.root.after(0, self._laser_watch_apply, events)
        except (RuntimeError, tk.TclError):
            pass                                # the window is gone; nothing to tell

    def _laser_watch_pass(self, list_ports, probe, now):
        """One pass of the watch, synchronously: what the worker does, and
        what the tests call with a simulated port list, a fake probe and a
        fake clock. Returns the watch's events."""
        watch = self._laser_watch
        ports = list_ports()
        for p in watch.look(ports, now()):
            watch.probed(p['device'], probe(p['device']), now())
            if watch.known_port:
                break               # the laser answered: no other port is opened
        watch.finish()
        return watch.pop_events()

    def _laser_watch_apply(self, events):
        """Tk thread: act on what a pass found."""
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        machine_log = logging.getLogger("ssc.machine")
        for event in events:
            kind = event[0]
            if kind == 'found':
                _k, device, how, info = event
                machine_log.info("laser found on %s (%s, %s:%s) by %s", device,
                                 info.get('description', ''),
                                 f"{info['vid']:04X}" if info.get('vid') else "-",
                                 f"{info['pid']:04X}" if info.get('pid') else "-",
                                 "its status report" if how == 'status' else "the boot banner after a reset")
                self._on_falcon_detected(device)
            elif kind == 'lost':
                self._on_falcon_lost(event[1])
            elif kind == 'silent':
                _k, device, info = event
                machine_log.info("port %s (%s) did not answer the probe", device, info.get('description', ''))
            elif kind == 'rescan':
                _k, probed, found = event
                machine_log.info("rescan asked for: %d port(s) probed, %s", probed,
                                 f"laser on {found}" if found else "no laser")
            elif kind == 'error':
                # A pass that raised (port enumeration failing, say) would
                # otherwise say so every 2 s: once per distinct message.
                if event[1] != getattr(self, '_laser_watch_last_error', None):
                    self._laser_watch_last_error = event[1]
                    machine_log.warning("laser watch: %s", event[1])

    def _on_falcon_detected(self, port):
        """Tk thread: a laser answered on `port`."""
        if port:
            self.falcon_port = port
            # Frame & Cut (both tabs) shows once a Falcon is detected;
            # _refresh_machine_ui_state holds the rule.
            self._refresh_machine_ui_state()

    def _on_falcon_lost(self, port):
        """Tk thread: the known port is no longer listed. The machine may
        have been power-cycled, so its zero is gone with it: the homed flag
        goes too, and Try Auto Locate (a G53 move) is not offered until the
        laser is homed again."""
        logging.getLogger("ssc.machine").info("laser gone from %s", port)
        self.falcon_port = None
        self._falcon_homed_this_session = False
        self._refresh_machine_ui_state()

    def _detect_falcon_async(self):
        """The user asked for the laser (a Machine menu item with none
        detected, Frame & Cut with none, a connect that failed): a full
        rescan, every port present, Bluetooth last. Only while the watch
        is running, which is every real launch with machine integration on;
        in a suite or a tour nothing is probed. The name is kept because
        suites and outside scripts stand in for it."""
        watch = getattr(self, '_laser_watch', None)
        if watch is None:
            return
        watch.request_rescan()
        if not self._laser_watch_busy:
            self._laser_watch_busy = True
            start_worker(self._laser_watch_worker, name='laser-watch')

    def _machine_enabled(self):
        """True iff the user has opted into machine integration.

        The underlying Falcon serial path (pyserial + Grbl 1.1
        character-counting protocol) is cross-platform — Linux and
        macOS users with a Falcon on USB can use it via `/dev/ttyUSB*`
        or `/dev/tty.usbserial-*`. The Falcon-camera auto-detect
        heuristic is per-platform (PowerShell on Windows,
        system_profiler on macOS, /sys/class/video4linux on Linux);
        wherever it fails, _resolve_camera_index falls through to the
        last-enumerated camera and the user can Switch camera in any
        dialog if that picks wrong.

        Single source of truth used by every gating site. Tested on
        Windows; the macOS/Linux paths are best-effort and the user
        is opting in via the experimental toggle."""
        return (machine_supported()
                and bool(self.settings.get("experimental_machine_menu", False)))

    def _camera_calibration_present(self):
        """True iff a usable (non-legacy) camera calibration is on disk."""
        try:
            import camera_capture
        except ImportError:
            return False
        if not camera_capture.HAS_OPENCV:
            return False
        return camera_capture.load_calibration(
            camera_capture.default_calibration_path()) is not None

    def _refresh_machine_ui_state(self):
        """Single source of truth for which machine UI is enabled.

        Since 2026-11-09 the laser and the camera are gated separately.
        Matt was asked "Should Frame & Cut require a camera calibration?"
        and answered "No, on either tab".
          - experimental_machine_menu OFF → nothing machine-related
            visible (Machine menu wasn't built at all)
          - experimental_machine_menu ON → the Machine menu's laser
            items are enabled (Home Laser, Test Connection, Clear Errors,
            Reset Falcon, Framing Power, and Camera Calibration as
            always), and Frame & Cut shows on the Pad Maker and on the
            Tooling tab once a Falcon is detected
          - ... + a camera calibration on disk → everything that uses
            the camera: Camera-Polygon Inset Margin here, the live
            camera tick below, and through _camera_capture_ready() "Get
            from camera", the live overlay, Try Auto Locate and
            "Re-capture from camera"

        Until then the calibration gated all of it (commit 021d8df,
        2026-05-26, "Two-level machine gating"). That was a rule, not a
        dependency: typed dimensions or a hand-drawn shape, jog to the
        corner, G92, framing loop, cut never read the calibration.

        Safe to call any time; idempotent. Wired into:
          - app startup (after menu construction)
          - _on_falcon_detected (Falcon comes online)
          - after CameraCalibrationDialog returns saved (live un-grey)
        """
        toggle_on = self._machine_enabled()
        has_cal = self._camera_calibration_present()

        # Machine menu item states. Only meaningful if the cascade was
        # actually built (toggle was on at startup).
        if (toggle_on and hasattr(self, '_machine_menu')
                and self._machine_menu_indices):
            # The laser items need only the toggle. Camera Calibration is
            # the camera's gateway and is always enabled. The inset margin
            # only means something for a camera-captured shape, so it alone
            # still waits for a calibration.
            menu = self._machine_menu
            indices = self._machine_menu_indices
            try:
                menu.entryconfig(indices['camera_cal'], state="normal")
                for key in MACHINE_MENU_LASER_ITEMS:
                    if key in indices:
                        menu.entryconfig(indices[key], state="normal")
                for key in MACHINE_MENU_CAMERA_ITEMS:
                    if key in indices:
                        menu.entryconfig(indices[key],
                                         state="normal" if has_cal else "disabled")
            except tk.TclError:
                pass

        # Frame & Cut button visibility: toggle ON and a Falcon detected
        # on USB. No calibration (see the docstring). The packing happens
        # here so both gates resolve through one helper.
        btn = getattr(self, '_frame_cut_btn', None)
        if btn is not None:
            should_show = bool(toggle_on and self.falcon_port)
            try:
                is_packed = bool(btn.winfo_manager())
            except tk.TclError:
                is_packed = False
            if should_show and not is_packed:
                try:
                    btn.pack(side="left", padx=5)
                except tk.TclError:
                    pass
            elif not should_show and is_packed:
                try:
                    btn.pack_forget()
                except tk.TclError:
                    pass

        # Tooling's Frame & Cut buttons (one per section that writes
        # G-code): the same rule as the Pad Maker's button above, toggle
        # ON and a Falcon detected. (machine_supported(), i.e. pyserial,
        # is part of _machine_enabled().)
        show_tooling = bool(toggle_on and self.falcon_port)
        for t_btn, pack_options in getattr(self, '_tooling_frame_cut_btns', ()):
            try:
                is_packed = bool(t_btn.winfo_manager())
                if show_tooling and not is_packed:
                    t_btn.pack(**pack_options)
                elif not show_tooling and is_packed:
                    t_btn.pack_forget()
            except tk.TclError:
                pass

        # Live camera preview checkbox: same gate as the camera-capture
        # features (toggle + calibration). Falcon doesn't need to be
        # detected — the checkbox is a setting consumed by Frame & Cut
        # only when that button is also packed (which DOES need Falcon),
        # but the user can preset it any time post-calibration.
        chk = getattr(self, '_live_camera_chk', None)
        if chk is not None:
            should_show_chk = bool(toggle_on and has_cal)
            try:
                is_packed = bool(chk.winfo_manager())
            except tk.TclError:
                is_packed = False
            if should_show_chk and not is_packed:
                try:
                    chk.pack(side="left", padx=(0, 15))
                except tk.TclError:
                    pass
            elif not should_show_chk and is_packed:
                try:
                    chk.pack_forget()
                except tk.TclError:
                    pass

    # ==================================================================
    # Machine menu handlers (Pad Maker > Machine cascade)
    # ==================================================================

    def _machine_require_falcon(self):
        """Common pre-flight: ensure Falcon is detected. Returns True
        if good, False (with an error popup) if not."""
        if not self.falcon_port:
            messagebox.showerror(
                _("Falcon Not Detected"),
                _("No Grbl controller detected. Plug in the Falcon "
                  "over USB and try again."))
            self._falcon_detect_attempted = False
            self._detect_falcon_async()
            return False
        return True

    def _resolve_camera_index(self):
        """Pick a camera index for the camera dialogs.

        Order: persisted override → Falcon-name heuristic → highest
        enumerated index. Returns None if no camera is available or
        OpenCV isn't installed.

        Persist a working index after a successful calibration save (via
        _persist_camera_index) so subsequent runs skip the guess.
        """
        try:
            import camera_capture
        except ImportError:
            return None
        if not camera_capture.HAS_OPENCV:
            return None
        override = self.settings.get("camera_index_override")
        if override is not None:
            try:
                return int(override)
            except (TypeError, ValueError):
                pass
        idx = camera_capture.find_falcon_camera_index()
        if idx is not None:
            return idx
        cams = camera_capture.enumerate_cameras()
        if cams:
            return cams[-1]['index']
        return None

    def _persist_camera_index(self, index):
        """Save a known-working camera index to settings so future
        camera dialogs open it directly without re-guessing. Called
        from the calibration dialog after a successful save."""
        if index is None:
            return
        try:
            current = self.settings.get("camera_index_override")
            if current == int(index):
                return
            self.settings["camera_index_override"] = int(index)
            from config import save_settings
            save_settings(self.settings)
        except Exception as e:  # never block on settings save
            logging.getLogger("ssc.machine").warning(
                "could not save the camera number %s to the settings: %s", index, e)

    def _run_blocking_home(self, sender):
        """Send $H on a worker thread, show the homing-status modal,
        block the caller until homing completes (or fails). Returns
        True on success, False on failure (with an error popup
        already shown).

        Uses an after()-poll loop on the Tk thread instead of the
        older `while t.is_alive(): root.update()` pattern — root.update()
        pumps the full Tk event loop including WM_DELETE_WINDOW, which
        could destroy root mid-loop during the 60s home and produce a
        TclError on the next iteration.
        """
        result = {}

        def _worker():
            try:
                ok, m = sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = m
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        self._show_homing_status()
        t = start_worker(_worker, name='home-laser')
        # after()-poll until the worker finishes. wait_variable lets
        # the Tk event loop keep running (status modal stays painted,
        # other dialogs respond) without burning CPU in a tight loop.
        done_var = tk.BooleanVar(value=False)

        def _check():
            if t.is_alive():
                self.root.after(50, _check)
            else:
                done_var.set(True)

        self.root.after(50, _check)
        self.root.wait_variable(done_var)
        self._hide_homing_status()
        if result.get('ok'):
            self._falcon_homed_this_session = True
            return True
        messagebox.showerror(
            _("Homing Failed"),
            _("Falcon returned: {m}").format(
                m=result.get('msg', 'no response')))
        return False

    def _on_machine_home_falcon(self):
        """Standalone $H. Same flow as Frame & Cut's home prompt."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Home needs pyserial: {e}").format(e=e))
            return
        if not messagebox.askyesno(
                _("Home Laser?"),
                _("Send the laser home? The head will travel to the "
                  "home corner at full speed. Make sure nothing is "
                  "in its path.")):
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  _("Couldn't connect: {e}").format(e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            self._detect_falcon_async()         # the port may be stale: look again
            return
        success = self._run_blocking_home(sender)
        try:
            sender.disconnect()
        except Exception:
            pass
        if success:
            messagebox.showinfo(_("Homed"),
                                 _("Laser homed. Head is parked at "
                                   "the home corner."))

    def _on_machine_test_connection(self):
        """Quick Falcon ping: connect, read status, disconnect, report."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Test needs pyserial: {e}").format(e=e))
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            status = sender.get_status(timeout=0.5)
        except Exception as e:
            messagebox.showerror(
                _("Connection Failed"),
                _("Couldn't connect to the Falcon at {p}: {e}").format(
                    p=self.falcon_port, e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            self._detect_falcon_async()         # the port may be stale: look again
            return
        finally:
            pass
        try:
            sender.disconnect()
        except Exception:
            pass
        logging.getLogger("ssc.machine").info(
            "Machine > Test Connection on %s: %s", self.falcon_port,
            status.get('raw') if status else "no status reply")
        if not status:
            messagebox.showwarning(
                _("Connected, no response"),
                _("Port opened at {p}, but Grbl didn't reply to a "
                  "status query. Try the Reset Falcon menu item, "
                  "then test again.").format(p=self.falcon_port))
            return
        mpos = status.get('mpos') or [0, 0, 0]
        messagebox.showinfo(
            _("Connection OK"),
            _("Falcon connected on {p}.\n\n"
              "State: {s}\n"
              "MPos: X={x:.2f}  Y={y:.2f}  Z={z:.2f}").format(
                p=self.falcon_port, s=status.get('state', '?'),
                x=mpos[0], y=mpos[1], z=mpos[2]))

    def _on_machine_clear_errors(self):
        """Send $X to unlock motion after an alarm. Doesn't reset
        position info, just clears the alarm state."""
        if not self._machine_require_falcon():
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Clear needs pyserial: {e}").format(e=e))
            return
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  _("Couldn't connect: {e}").format(e=e))
            try:
                sender.disconnect()
            except Exception:
                pass
            self._detect_falcon_async()         # the port may be stale: look again
            return
        finally:
            pass
        try:
            sender.disconnect()
        except Exception:
            pass
        messagebox.showinfo(
            _("Errors Cleared"),
            _("Sent $X unlock. Motion is allowed again. If the "
              "Falcon was in an alarm state, it should now be Idle "
              "— check with Test Connection."))

    def _on_machine_reset_falcon(self):
        """Send Ctrl-X (RT_SOFT_RESET) to reboot the controller's
        firmware state. Stops any in-progress motion immediately,
        clears alarms, resets the planner buffer."""
        if not self._machine_require_falcon():
            return
        if not messagebox.askyesno(
                _("Reset Falcon?"),
                _("Send a soft-reset (Ctrl-X) to the Falcon. This:\n"
                  "  • Immediately stops any motion in progress\n"
                  "  • Clears alarm + error states\n"
                  "  • Resets the planner buffer\n"
                  "  • Loses the machine's known position (you'll "
                  "need to home again)\n\n"
                  "Only use this when the Falcon is stuck in a bad "
                  "state and Clear Errors didn't help.\n\n"
                  "Continue?")):
            return
        try:
            import falcon_sender
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Reset needs pyserial: {e}").format(e=e))
            return
        logging.getLogger("ssc.machine").info(
            "Machine > Reset Falcon: soft-reset on %s", self.falcon_port)
        sender = falcon_sender.FalconSender(port=self.falcon_port)
        # One outcome, one box (2026-11-10): a failed reset used to show
        # "Reset Failed" and then the success box, and cleared the homed
        # flag although the machine had not been reset and still knew
        # its position.
        try:
            sender.connect()
            if not sender.send_realtime(falcon_sender.RT_SOFT_RESET):
                raise RuntimeError(_("the reset byte could not be sent"))
            time.sleep(1.0)  # let firmware reboot
        except Exception as e:
            logging.getLogger("ssc.machine").warning(
                "Machine > Reset Falcon failed on %s: %s", self.falcon_port, e)
            messagebox.showerror(_("Reset Failed"),
                                  _("Couldn't reset: {e}").format(e=e))
            return
        finally:
            try:
                sender.disconnect()
            except Exception:
                pass
        # Soft-reset loses machine position — clear the homed flag so
        # the next Frame & Cut auto-homes again.
        self._falcon_homed_this_session = False
        messagebox.showinfo(
            _("Falcon Reset"),
            _("Soft-reset sent. The Falcon is now in alarm state — "
              "use Home Laser to recover, then jog as needed."))

    def _on_machine_recalibrate(self):
        """Open Camera Calibration with a strong warning since the
        engrave + capture flow takes the better part of an hour."""
        try:
            import camera_capture
        except ImportError:
            messagebox.showerror(
                _("OpenCV Required"),
                _("Camera calibration needs OpenCV:\n\n"
                  "    pip install opencv-python Pillow"))
            return
        cal_path = camera_capture.default_calibration_path()
        existing = camera_capture.load_calibration(cal_path) is not None
        if existing:
            if not messagebox.askyesno(
                    _("Recalibrate Camera?"),
                    _("⚠ This will REPLACE your existing calibration. "
                      "⚠\n\n"
                      "The full calibration takes ~60 minutes — the "
                      "calibration card has to be engraved on fresh "
                      "basswood, then you take 12+ photos.\n\n"
                      "Only recalibrate if:\n"
                      "  • The camera was bumped / re-mounted\n"
                      "  • Cuts are landing in the wrong spot\n"
                      "  • You're sure the existing one is wrong\n\n"
                      "Proceed?")):
                return
        # Delegate to the existing handler (inherited from
        # ToolingTabMixin) — same dialog, same flow.
        self._open_camera_calibration()

    def _on_machine_framing_power(self):
        """Per-material framing power (Grbl S, 0-1000 scale).

        Framing traces the outline so you can see where the head will go
        without marking anything. How low "visible but harmless" is
        depends on the material — 1% reads fine on card and is invisible
        on dark leather. Framing always uses M4 dynamic power whatever
        the value, so the beam still drops to zero when the head stalls
        at a vertex.
        """
        from config import (GCODE_PRESET_MATERIALS, FRAMING_POWER_S_MAX,
                            get_framing_power_s, save_settings)

        dlg = tk.Toplevel(self.root)
        dlg.title(_("Framing Power"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        bg = self._get_theme_color() if hasattr(
            self, '_get_theme_color') else None
        if bg:
            dlg.configure(bg=bg)

        max_pct = FRAMING_POWER_S_MAX / 10.0
        tk.Label(dlg, justify='left', wraplength=440, bg=bg,
                 text=_("Laser power used for the framing pass, per "
                        "material. Framing shows you where the cut will "
                        "land — it should be bright enough to follow by "
                        "eye without marking the material.\n\n"
                        "Dark materials like leather need more than card "
                        "does. Raise it a little at a time: the power "
                        "that becomes visible is usually close to the "
                        "power that starts leaving a mark.")
                 ).pack(padx=20, pady=(15, 10))

        grid = tk.Frame(dlg, bg=bg)
        grid.pack(padx=20, pady=(0, 10))
        vars_by_mat = {}
        # Exposed for tests: which materials this dialog is showing.
        dlg.vars_by_mat = vars_by_mat
        # Adhesive rings get a row only while the feature is on.
        shown = [m for m in GCODE_PRESET_MATERIALS
                 if m != 'adhesive_ring' or self._adhesive_ring_enabled()]
        for row, mat in enumerate(shown):
            row_text = (_("Adhesive rings") if mat == 'adhesive_ring'
                        else mat.replace('_', ' ').title()) + ":"
            tk.Label(grid, text=row_text,
                     bg=bg, anchor='e', width=max(12, len(row_text))).grid(
                row=row, column=0, sticky='e', pady=2)
            # Shown and entered as a percentage, like every other power
            # field in the app. Stored as a Grbl S value (percent x 10) —
            # see _ok() for why the stored unit doesn't change.
            # One decimal always shown ("1.0", not "1") so it's obvious a
            # fractional percent is accepted — the useful range is narrow
            # and S values are integers, so 0.1% is exactly the precision
            # available.
            pct_now = get_framing_power_s(mat, self.settings) / 10.0
            var = tk.StringVar(value=f"{pct_now:.1f}")
            vars_by_mat[mat] = var
            tk.Entry(grid, textvariable=var, width=8,
                     font=("Helvetica", 11)).grid(
                row=row, column=1, sticky='w', padx=6, pady=2)
            tk.Label(grid, text="%", bg=bg).grid(row=row, column=2,
                                                 sticky='w')

        tk.Label(dlg, bg=bg, fg="#666666", font=("Helvetica", 8),
                 text=_("Range: 0 to {max}%. Cutting power is typically "
                        "35-100%, so framing stays well below it.").format(
                            max=f"{max_pct:.1f}")
                 ).pack(padx=20, pady=(0, 8))

        def _ok():
            parsed = {}
            for mat, var in vars_by_mat.items():
                try:
                    pct = float(var.get())
                    if not (0 <= pct <= max_pct):
                        raise ValueError("out of range")
                except (ValueError, TypeError):
                    messagebox.showerror(
                        _("Invalid"),
                        _("{mat}: framing power must be a number from "
                          "0 to {max}%.").format(
                              mat=mat, max=f"{max_pct:.1f}"),
                        parent=dlg)
                    return
                # Store as Grbl S (percent x 10). Kept in S rather than
                # migrating the setting to percent because the existing
                # key already holds S values: reinterpreting a stored 10
                # as 10% would silently raise everyone's framing power
                # tenfold, and a botched migration here burns material.
                parsed[mat] = int(round(pct * 10))
            # Merge rather than replace, so a material with no row here
            # (adhesive rings while that feature is off) keeps its value.
            by_mat = self.settings.get("laser_framing_power_by_material")
            merged = dict(by_mat) if isinstance(by_mat, dict) else {}
            merged.update(parsed)
            self.settings["laser_framing_power_by_material"] = merged
            try:
                save_settings(self.settings)
            except Exception as e:
                messagebox.showerror(_("Save Failed"), str(e), parent=dlg)
                return
            dlg.destroy()

        btn_row = tk.Frame(dlg, bg=bg)
        btn_row.pack(pady=(0, 15))
        tk.Button(btn_row, text=_("OK"), command=_ok, width=10).pack(
            side='left', padx=5)
        tk.Button(btn_row, text=_("Cancel"), command=dlg.destroy,
                  width=10).pack(side='left', padx=5)

    def _on_machine_inset_settings(self):
        """Small dialog: edit camera_polygon_inset_mm."""
        dlg = tk.Toplevel(self.root)
        dlg.title(_("Camera-Polygon Inset Margin"))
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        bg = self._get_theme_color() if hasattr(
            self, '_get_theme_color') else None
        if bg:
            dlg.configure(bg=bg)

        tk.Label(dlg, justify='left', wraplength=420,
                  text=_("Camera-captured polygons are shrunk by this "
                          "many millimeters on every edge before "
                          "nesting. This safety margin compensates "
                          "for camera measurement error near the "
                          "edges of its view — keeps pads from "
                          "landing where a chunk of leather might "
                          "actually be missing.\n\n"
                          "Typical: 2-5mm. Set to 0 to disable.")
                  ).pack(padx=20, pady=(15, 10))

        row = tk.Frame(dlg)
        row.pack(pady=(0, 15))
        tk.Label(row, text=_("Inset margin (mm):")).pack(side='left',
                                                            padx=5)
        var = tk.StringVar(
            value=str(self.settings.get("camera_polygon_inset_mm", 3.0)))
        entry = tk.Entry(row, textvariable=var, width=8,
                          font=("Helvetica", 11))
        entry.pack(side='left', padx=5)
        entry.focus_set()

        def _ok():
            try:
                v = float(var.get())
                if v < 0 or v > 50:
                    raise ValueError("out of range")
            except (ValueError, TypeError):
                messagebox.showerror(_("Invalid"),
                                      _("Inset must be a number 0-50."),
                                      parent=dlg)
                return
            self.settings["camera_polygon_inset_mm"] = v
            from config import save_settings
            try:
                save_settings(self.settings)
            except Exception as e:
                messagebox.showerror(_("Save Failed"),
                                      str(e), parent=dlg)
                return
            dlg.destroy()

        btn_row = tk.Frame(dlg)
        btn_row.pack(pady=(0, 15))
        tk.Button(btn_row, text=_("Save"), command=_ok,
                   width=10, font=("Helvetica", 10, "bold")
                   ).pack(side='left', padx=5)
        tk.Button(btn_row, text=_("Cancel"), command=dlg.destroy,
                   width=10).pack(side='left', padx=5)

    def _show_homing_status(self):
        """Modal 'Homing...' window held open while $H blocks.

        Created as a Toplevel rather than a messagebox so we can dismiss
        it programmatically when homing returns. Lives in
        self._homing_status_win.
        """
        win = tk.Toplevel(self.root)
        win.title(_("Homing"))
        win.transient(self.root)
        win.resizable(False, False)
        try:
            win.grab_set()
        except tk.TclError:
            pass
        tk.Label(win, text=_("Homing the laser...\n\n"
                              "Head is moving to the home corner."),
                 padx=30, pady=20,
                 font=("Helvetica", 11), justify='left').pack()
        win.update_idletasks()
        # Center on root window
        try:
            x = (self.root.winfo_x()
                 + (self.root.winfo_width() // 2)
                 - (win.winfo_width() // 2))
            y = (self.root.winfo_y()
                 + (self.root.winfo_height() // 2)
                 - (win.winfo_height() // 2))
            win.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass
        self._homing_status_win = win
        win.update()

    def _hide_homing_status(self):
        win = getattr(self, '_homing_status_win', None)
        if win is not None:
            try:
                win.destroy()
            except tk.TclError:
                pass
            self._homing_status_win = None

    def on_frame_and_cut(self, tooling_job=None):
        """Generate G-code in memory, frame the bbox, then stream to Falcon.

        One routine for both tabs. The Pad Maker's button calls it bare and
        the pad job is validated, nested, previewed and generated in the
        first half below. The Tooling tab passes `tooling_job`, a function
        that does the same for one of its tools and returns what the
        machine half needs (the _on_frame_cut_* handlers in tooling_tab.py
        build it), or None when the user backed out or the input was
        refused.

        From the framing lines on, both run the same code: connect, jog to
        the corner (the lid question on Start Frame), G92, the framing
        loop, the cut, last_frame_cut.gcode. The machine half reads the
        job's own values only, never the Pad Maker's shape or session, so
        a scrap outline loaded on the Pad Maker cannot bend a Tooling cut.
        """
        try:
            import falcon_sender
            from gcode_engine import (
                extract_gcode_bbox,
                generate_framing_gcode,
                generate_polygon_framing_gcode,
            )
            from ui_dialogs import FalconRunDialog
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Frame & Cut needs pyserial:\n\n{e}").format(e=e))
            return

        if not self.falcon_port:
            messagebox.showerror(_("Falcon Not Detected"),
                                  _("No Grbl controller detected. Plug in "
                                    "the Falcon over USB and try again."))
            self._falcon_detect_attempted = False
            self._detect_falcon_async()
            return

        # No "would you like to recapture from camera?" prompt here —
        # the Draw / Capture Shape dialog already offered both paths
        # before the user picked one. If they drew a polygon, that's
        # their choice; proceeding in manual mode is the right
        # default. (An earlier two-button layout had Draw and Capture
        # as separate top-level buttons, and this prompt was the
        # bridge between them; with the consolidated dialog it just
        # makes users repeat work they already chose against.)

        # Long cuts can run for tens of minutes; block sleep so the
        # USB serial doesn't get killed mid-job. Released in the
        # finally block below, so nothing may return between here and
        # the try: the no-laser return used to sit after this and left
        # the lock held.
        try:
            import sleep_lock
            sleep_lock.prevent_sleep()
        except Exception:
            logging.getLogger("ssc.machine").exception(
                "Frame & Cut: could not take the sleep lock; the computer may sleep during the cut")

        try:
            if tooling_job is None:
                # --- Pad Maker job: input validation up through nesting
                # (mirrors on_generate_gcode) ---
                params = self._prepare_generation()
                if not params:
                    return
                pads, hole_dia = params['pads'], params['hole_dia']
                width_mm, height_mm = params['width_mm'], params['height_mm']
                card_paper_dims = params['card_paper_dims']

                supported_materials = [m for m, v in self.material_vars.items()
                                         if v.get() and m != "exact_size"]
                if len(supported_materials) != 1:
                    messagebox.showerror(
                        _("Pick One Material"),
                        _("Frame & Cut works with one material at a time. "
                          "Please select exactly one material (not Exact Size)."))
                    return
                material = supported_materials[0]
                mat_w, mat_h, mat_polygon = self._get_material_dimensions(
                    material, width_mm, height_mm, card_paper_dims)

                # Scrap mode places a partial batch on this scrap (one scrap
                # per click, mirroring _generate_gcode_scrap_mode); standard
                # mode nests the whole batch and requires it all to fit.
                # scrap_remaining is handed to _frame_cut_scrap_advance once
                # the cut completes; it stays None in standard mode.
                scrap_mode = self.scrap_mode_var.get()
                scrap_remaining = None
                if scrap_mode:
                    result = self._scrap_begin_partial(
                        pads, hole_dia, material, mat_w, mat_h, mat_polygon,
                        ask_save_dir=False)
                    if result is None:
                        return
                    placed, scrap_remaining = result
                    hole_dia = self.scrap_session['hole_dia']
                    # Labeled zones stand down in scrap mode: a zone is sized for
                    # a whole size's count, but a scrap only takes part of it, so
                    # the block would have to be re-sized per scrap.
                    zones = []
                else:
                    # One nest; its fixed counts are the fit check.
                    placed, zones, fixed_placed, fixed_total = nest_with_zones(
                        pads, material, mat_w, mat_h, self.settings, polygon=mat_polygon)
                    if fixed_placed != fixed_total:
                        messagebox.showerror(
                            _("Nesting Error"),
                            _("Could not fit all '{m}' pieces in the available "
                              "area.").format(m=material.replace('_', ' ')))
                        return

                # Show the nesting layout before anything runs on the machine —
                # same as the file-export Generate flow, so the user can see what
                # will be cut and back out if it's not what they want. In scrap
                # mode this is the subset that fits on the current scrap; in
                # standard mode it's the whole batch. Cancelling (Adjust) backs out
                # cleanly: the scrap is only counted after a completed cut (in
                # _frame_cut_scrap_advance), so nothing is consumed here.
                preview = NestingPreviewWindow(
                    self.root, {material: placed}, mat_w, mat_h,
                    polygon=mat_polygon, proceed_label=_("Continue →"), zones=zones,
                    settings=self.settings)
                if preview.result != "save":
                    return

                # Generate G-code to a temp file then read it back into memory.
                import tempfile
                import os as _os
                with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False,
                                                  mode='w') as tmp:
                    tmp_path = tmp.name
                # The cut streams into the SAME work frame as the framing pass,
                # and framing flips the scrap OUTLINE with the outline's height.
                # A camera-captured polygon is inset from that outline, so the
                # generator's default (flip with the polygon it's given, i.e.
                # the inset) put every disc `inset` mm toward the machine front
                # of where the frame showed it. Flip with the outline instead.
                #
                # A job that was not nested on the loaded shape (card on
                # the paper size, the one case where mat_polygon is None
                # with a shape loaded) takes nothing from the shape, here
                # or in the machine half below: no flip override, no
                # outline to frame, no anchor for the G92, no Auto Locate
                # target. It is a plain sheet jogged to by its corner,
                # which is how the exported file has always treated it.
                on_shape = mat_polygon is not None
                _flip_source = ((self.custom_polygon_outline or self.custom_polygon)
                                if on_shape else None)
                flip_height_mm = (max(p[1] for p in _flip_source)
                                  if _flip_source and len(_flip_source) >= 3
                                  else None)
                try:
                    generate_gcode_from_placed(
                        placed, material, mat_w, mat_h, tmp_path, hole_dia,
                        self.settings, polygon=mat_polygon, zones=zones,
                        flip_height_mm=flip_height_mm)
                    with open(tmp_path, 'r') as f:
                        gcode_text = f.read()
                finally:
                    try:
                        _os.remove(tmp_path)
                    except OSError:
                        pass

                # What the machine half below takes from a pad job.
                # Frame the OUTLINE polygon (un-insetted) so the trace
                # matches the actual scrap edge. Cuts still use the inset
                # polygon for placement — both share a coordinate origin
                # so the G92 math below works either way. Fall back to
                # custom_polygon for plain hand-drawn polygons that don't
                # have a separate outline.
                framing_polygon = ((self.custom_polygon_outline
                                    or self.custom_polygon)
                                   if on_shape else None)
                # If the polygon has a known machine LB-vertex, offer
                # "Try Auto Locate" — drives the head directly there so
                # the user doesn't have to jog manually. Only enabled once
                # the laser is homed (otherwise MPos drift could put the
                # auto-drive somewhere wrong).
                #
                # Both camera extras need a calibration (the laser itself
                # does not, since 2026-11-09): without one Frame & Cut
                # runs with no Auto Locate target and opens no camera,
                # even if a tick or a position is left over from a time
                # when a calibration existed.
                camera_ready = self._camera_capture_ready()
                auto_locate_target = (self._custom_polygon_lb_machine
                                      if camera_ready and on_shape else None)
                show_live_camera = bool(self.live_camera_var.get() and camera_ready)
                cut_polygon, piece_count, job_label = mat_polygon, len(placed), None
            else:
                # --- Tooling job: the tool's own validation, preview and
                # generation, done by the function Tooling handed in. ---
                job = tooling_job()
                if not job:
                    return
                gcode_text = job['gcode_text']
                material = job['material']
                job_label = job['label']
                piece_count = job['count']
                # A Tooling sheet is a plain rectangle jogged to by hand:
                # no outline to trace, no camera position to drive to, no
                # flip override, whatever the Pad Maker has loaded.
                on_shape = False
                framing_polygon = None
                auto_locate_target = None
                cut_polygon = None
                flip_height_mm = None
                # The "Live camera preview" tick lives on the Pad Maker and
                # is only on screen with a camera calibration; honour it
                # here on the same condition, so a stale hidden tick can't
                # open a camera nobody set up.
                show_live_camera = bool(self.live_camera_var.get()
                                        and self._camera_capture_ready())

            # Framing pass. If a custom polygon is loaded (drawn or camera-
            # captured), trace the polygon outline — on irregular scrap a
            # bbox rectangle overhangs every concave edge and tells the user
            # nothing about whether the cuts land on material. Fall back to
            # the bbox rectangle for rectangular jobs (no polygon loaded).
            from config import get_framing_power_s
            power_s = get_framing_power_s(material, self.settings)
            feed = self.settings.get("laser_framing_feed", 2000)
            framing_lines = []
            if framing_polygon and len(framing_polygon) >= 3:
                framing_lines = generate_polygon_framing_gcode(
                    framing_polygon, power_s=power_s, feed=feed,
                )
            else:
                bbox = extract_gcode_bbox(gcode_text)
                if bbox:
                    xmin, xmax, ymin, ymax = bbox
                    framing_lines = generate_framing_gcode(
                        xmin, ymin, xmax, ymax, power_s=power_s, feed=feed,
                    )

            sender = falcon_sender.FalconSender(port=self.falcon_port)
            try:
                sender.connect()
            except Exception as e:
                messagebox.showerror(_("Connection Failed"),
                                      _("Could not connect to the Falcon at "
                                        "{p}: {e}").format(p=self.falcon_port, e=e))
                self._detect_falcon_async()     # the port may be stale: look again
                return

            # From here on the port is open: every way out (jog cancelled,
            # framing or cut stopped, an exception) goes through the finally
            # that disconnects. The try used to start at the framing loop,
            # so cancelling at the jog step left the port open until the
            # sender was garbage-collected, and a quick second Frame & Cut
            # could fail to open it.
            try:
                # Optional live camera window — opens alongside the run dialog
                # so the user can watch the laser head while it works, and stays
                # open after the cut so the user can see the finished result.
                # Stored on self so the reference persists past this function.
                # Only ONE window at a time: if a previous Frame & Cut left
                # a live cam open (it intentionally outlives the cut), focus
                # that one instead of opening a duplicate.
                if show_live_camera:
                    try:
                        import camera_capture
                        from ui_dialogs import LiveCameraWindow
                        existing = getattr(self, '_live_cam_window', None)
                        if existing is not None and existing.winfo_exists():
                            try:
                                existing.lift()
                                existing.focus_force()
                            except tk.TclError:
                                pass
                        elif camera_capture.HAS_OPENCV:
                            cam_idx = self._resolve_camera_index()
                            if cam_idx is not None:
                                self._live_cam_window = LiveCameraWindow(
                                    self.root, camera_index=cam_idx,
                                    settings=self.settings)
                    except Exception:  # never block the cut on preview failure
                        logging.getLogger("ssc.machine").exception(
                            "Frame & Cut: could not open the live camera preview")

                # After our connect+soft-reset, Grbl is in alarm state
                # (homing required). $X clears that so motion is allowed.
                # We don't $H here — that would fly the head to the homing
                # switches and override whatever position the user manually
                # set as their material origin.
                sender.unlock()

                # Manual positioning step: jog dialog stays open so the
                # user can fine-tune with the arrow buttons. They can ALSO
                # have positioned the head physically (open lid → push head
                # to material's lower-left → close lid → click Start Frame),
                # so the Start Frame button is enabled immediately — no
                # forced jog click. If the head's already where they want,
                # they just click Start Frame straight through.
                # auto_locate_target (pad jobs on a camera polygon only) adds
                # "Try Auto Locate" to this dialog.
                jog_dlg = FalconRunDialog(
                    self.root, sender, [],  # nothing to stream
                    title=_("Position the head at your material's "
                             "bottom-left corner (jog here OR open the "
                             "lid and physically move it), then click "
                             "Start Frame"),
                    show_pause_resume=False, show_cut_button=False,
                    stop_needs_confirm=False,
                    done_button_label=_("Start Frame →"),
                    remind_lid_on_advance=True,
                    auto_locate_target=auto_locate_target,
                    is_homed=self._falcon_homed_this_session)
                if jog_dlg._final_reason != "complete":
                    return

                # Compute the polygon's leftmost-lowest vertex in framing's
                # Y-UP frame. This is the corner the user intuitively jogs
                # to ("BL of the scrap"). Computed from the OUTLINE (not
                # the inset) so the LB-vertex sits on an actual scrap
                # edge — for axis-aligned scraps this is bbox-(0,0); for
                # tilted scraps it's the visible bottom-left corner.
                #
                # The inset polygon shares the same coordinate origin
                # (see _set_custom_polygon_from_y_up), so cut placements
                # remain correctly anchored even though the G92 offset
                # comes from the outline.
                #
                # A job that is not on the loaded shape (every Tooling job, and
                # a pad job for card on the paper size) has no polygon: its
                # sheet's bottom-left corner is where the head was jogged to,
                # so the origin is (0, 0) there, and _frame_cut_anchor (which
                # reads the Pad Maker's loaded shape) is not asked.
                if on_shape:
                    g92_x, g92_y = self._frame_cut_anchor()
                else:
                    g92_x, g92_y = 0.0, 0.0

                # Both modes share the same framing-loop recipe:
                # G92 X{lb_x} Y{lb_y} sets work coords such that the
                # polygon's bbox-BL ends up at work (0, 0) when the head
                # is physically at the polygon's LB vertex — what the
                # user thinks of as "BL of the material."
                prefix = [f'G92 X{g92_x:.3f} Y{g92_y:.3f}']
                if framing_lines:
                    framing_lines = prefix + framing_lines

                if framing_lines:
                    # Framing always runs when there's something to
                    # frame — the user's escape hatch is clicking
                    # "Looks Good — Cut!" immediately in the loop
                    # dialog to skip to the cut. Loop mode lets them
                    # take their time verifying alignment and nudge
                    # the head with the jog buttons between/during
                    # passes.
                    fdlg = FalconRunDialog(
                        self.root, sender, framing_lines,
                        title=_("Framing — verify cut area, jog if "
                                 "needed, click Cut when ready"),
                        loop=True)
                    if fdlg._final_reason != "complete":
                        return  # user stopped or error during framing
                    # No "Frame Looks Good?" prompt — the user already
                    # confirmed by clicking Cut Now.

                # Cut uses the same G92 offset as framing so the cut
                # placements align with the LB-vertex jog convention.
                cut_lines = [f'G92 X{g92_x:.3f} Y{g92_y:.3f}'] + gcode_text.splitlines()
                if tooling_job is None:
                    cut_scrap_num = ((self.scrap_session['scrap_count'] + 1)
                                     if scrap_mode else None)
                else:
                    cut_scrap_num = job.get('scrap_num')
                self._keep_last_cut_gcode(
                    cut_lines, material, cut_scrap_num,
                    (g92_x, g92_y), flip_height_mm, cut_polygon, piece_count,
                    job_label=job_label)
                cut_dlg = FalconRunDialog(
                    self.root, sender, cut_lines,
                    title=_("Cutting — {m}").format(m=material.replace('_', ' ')))

                if tooling_job is not None:
                    # Logged like a pad run, stopped and errored cuts too
                    # (the material went under the laser); the Die Scrap
                    # sheet number travels with the job. Since 2026-11-09;
                    # before, Job History took no Tooling runs.
                    if job.get('tool'):
                        self._record_tooling_job(job, "laser",
                                                 status=cut_dlg._final_reason or "stopped")
                    # Tooling's own bookkeeping (its scrap session, the
                    # kerf test's measuring steps). It is told how the cut
                    # ended and decides for itself what "complete" earns.
                    after_cut = job.get('after_cut')
                    if after_cut is not None:
                        after_cut(cut_dlg._final_reason)
                    return

                # Log it before the scrap bookkeeping below, which can
                # sit on a modal dialog. Stopped and errored cuts are
                # logged too, with their status — the material still
                # went under the laser, and knowing a job was cut short
                # is exactly what you want when pads come up missing.
                # scrap_count is bumped by _frame_cut_scrap_advance, so
                # the number this cut belongs to is one past it.
                self._record_job(
                    "laser", [material], pads, params,
                    placed_count=len(placed),
                    scrap_num=(self.scrap_session['scrap_count'] + 1
                               if scrap_mode else None),
                    status=cut_dlg._final_reason or "stopped")

                # Scrap mode: commit this scrap (decrement remaining,
                # offer continue/recapture) only once the cut has
                # streamed to completion. A stopped or errored cut
                # leaves the session untouched so it can be re-cut.
                if scrap_mode and cut_dlg._final_reason == "complete":
                    self._frame_cut_scrap_advance(len(placed), scrap_remaining)
            finally:
                # Intentionally don't auto-close live_cam — users want
                # to see the finished result. They close it manually,
                # or it dies with the app when SSC exits.
                try:
                    sender.disconnect()
                except Exception:
                    pass

        except Exception as e:
            messagebox.showerror(_("An Error Occurred"),
                                  _("Something went wrong during Frame & Cut:\n\n"
                                    "{error}").format(error=e))
        finally:
            try:
                import sleep_lock
                sleep_lock.allow_sleep()
            except Exception:
                pass
            # Every early return above (preview Adjust, jog or framing
            # cancelled, stopped or errored cut, exception) lands here
            # with the scrap still pending; a completed cut has already
            # committed it, so this is a no-op then.
            self._scrap_abandon_pending()
