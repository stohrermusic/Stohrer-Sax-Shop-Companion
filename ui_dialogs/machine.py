"""The laser and the camera: the shape dialog with its overlay, camera calibration and capture, the lid question, the run dialog, the live camera window.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import messagebox
import functools
import math
import os
import time

from threads import GcHeldRLock, start_worker

from svg_engine import (
    polygon_anchor_vertex,
)

from .common import DIALOG_BG, mlog
import ui_dialogs as _pkg   # the dialogs suites stand in for are looked up through the package at call time


def _holds_port(method):
    """CameraCalibrationDialog: run ``method`` holding the dialog's serial
    port lock. The head-position poll reads the laser from a worker thread
    and skips a tick whenever the lock is taken, so a button that opens
    the port never collides with a poll (Windows serial ports are
    exclusive), and the poll never blocks the Tk thread."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        with self._port_lock:
            return method(self, *args, **kwargs)
    return wrapper


class PolygonDrawWindow(tk.Toplevel):
    """
    A dialog for drawing a polygon shape on a grid.
    Used for defining irregular leather skin shapes.
    """
    # 32 is comfortable for the canvas + handles camera-captured contours
    # well; previous 8-cap was a leftover from when nesting was slow and
    # high-vertex polygons made the per-pad point-in-polygon checks crawl.
    # Vectorized nesting (v2.40) handles 30+ vertices easily.
    MAX_POINTS = 32
    CANVAS_PX = 450  # Canvas size in pixels
    POINT_RADIUS = 6  # Radius of drawn points in pixels
    CLOSE_THRESHOLD = 15  # Pixels - how close to first point to auto-close
    OVERLAY_REFRESH_MS = 200  # 5 fps for the live camera underlay

    def __init__(self, parent, unit="in", settings=None):
        super().__init__(parent)
        self.unit = unit
        self.polygon_closed = False
        self.result = None  # Will hold the final polygon or None if cancelled

        # Grid size — pick the larger of the user's setting and what's
        # needed to cover the laser bed at 1:1 scale. Absolute-position
        # polygons (camera capture + hand-trace with overlay) extend
        # across the bed, so a too-small grid would clip them. The bed-
        # derived floor auto-upgrades users whose setting predates this
        # change without forcing a migration write.
        self._settings = settings or {}
        if self.unit == "in":
            try:
                self.grid_size = int(
                    self._settings.get("polygon_draw_grid_size_in", 17))
            except (TypeError, ValueError):
                self.grid_size = 17
            mm_per_unit_for_grid = 25.4
        else:
            try:
                self.grid_size = int(
                    self._settings.get("polygon_draw_grid_size_cm", 43))
            except (TypeError, ValueError):
                self.grid_size = 43
            mm_per_unit_for_grid = 10.0
        try:
            bed_max_mm = max(
                float(self._settings.get("laser_bed_x_max", 400)),
                float(self._settings.get("laser_bed_y_max", 415)))
        except (TypeError, ValueError):
            bed_max_mm = 415.0
        self.grid_size = max(
            self.grid_size,
            int(math.ceil(bed_max_mm / mm_per_unit_for_grid)))
        self.grid_size = max(2, self.grid_size)  # sanity floor

        self.points = []  # List of (x, y) in grid units, 0..grid_size each
        # Set when the user captures from the camera. Anchors the grid's
        # local (0, 0) to an absolute machine-mm point so the polygon
        # can be reconstructed in machine coords (for nesting that knows
        # the bed position) AND so a live camera overlay can render at
        # the same machine→grid scale the polygon was captured at.
        # Preserved across vertex edits — the grid frame doesn't move
        # when you tweak a polygon point. Cleared only on Clear or
        # Cancel.
        self._camera_anchor_mm = None
        # Sticky "this polygon has a meaningful machine-coord anchor"
        # flag. Set the first time the camera overlay turns on (overlay
        # always anchors at machine (0, 0), so any clicks the user
        # makes correspond to absolute machine positions in grid units).
        # Once True, never cleared mid-session — even if the user
        # disables the overlay, the polygon they've already traced
        # still carries machine meaning. Used by Frame & Cut to enable
        # the "Try Auto Locate" button (drives the head to the
        # polygon's LB vertex). Plain hand-drawn polygons (never enabled
        # overlay, never captured) keep this False — the button hides.
        self._overlay_used = False

        # Camera-overlay state (populated when the user toggles "Show
        # live camera"). None until enabled.
        self._show_camera_var = None     # BooleanVar — created in _create_widgets
        self._show_camera_chk = None     # Checkbutton widget reference
        self._cap = None
        self._cam_mod = None
        self._calibration = None
        self._PIL_Image = None
        self._PIL_ImageTk = None
        self._overlay_after_id = None
        self._overlay_errors = 0          # frames that failed to draw
        self._overlay_photo = None        # PhotoImage ref (prevent GC)

        # Title mirrors what the dialog can actually do: shows
        # "Draw / Capture Shape" when camera capture is available,
        # plain "Draw Shape" when only the draw path is wired up.
        # _camera_capture_available() checks both the experimental
        # toggle and calibration presence.
        self.title(_("Draw / Capture Shape")
                    if self._camera_capture_available()
                    else _("Draw Shape"))
        self.geometry("520x680")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        # Calculate pixels per grid unit (1 inch or 1 cm per square)
        self.px_per_unit = self.CANVAS_PX / self.grid_size

        self._create_widgets()
        self._draw_grid()

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window(self)

    def _create_widgets(self):
        # Instructions
        unit_label = "inches" if self.unit == "in" else "cm"
        instr_text = (
            f"Click anywhere to add a point (max {self.MAX_POINTS} points).\n"
            f"Click near the first point to close. Click a point to remove it."
        )
        tk.Label(self, text=instr_text, bg=DIALOG_BG, justify="center").pack(pady=(10, 5))

        # Grid info - each square = 1 unit (inch or cm)
        tk.Label(self, text=_("Grid: {size}x{size} {unit_label} (1 square = 1 {unit})").format(size=self.grid_size, unit_label=unit_label, unit=self.unit),
                 bg=DIALOG_BG, font=("Helvetica", 9)).pack(pady=(0, 5))

        # Canvas frame
        canvas_frame = tk.Frame(self, bg=DIALOG_BG)
        canvas_frame.pack(padx=10, pady=5)

        self.canvas = tk.Canvas(canvas_frame, width=self.CANVAS_PX, height=self.CANVAS_PX,
                                 bg="white", highlightthickness=1, highlightbackground="gray")
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        # Status label
        self.status_var = tk.StringVar(value="Click to add points...")
        tk.Label(self, textvariable=self.status_var, bg=DIALOG_BG, font=("Helvetica", 10)).pack(pady=5)

        # Live-camera overlay toggle. Available as soon as a camera
        # calibration is on disk — the user can either capture a
        # polygon first and then refine its vertices against the live
        # image, OR turn the overlay on with no capture and trace the
        # scrap shape by hand. When there's no capture, the overlay
        # uses the camera's image-bbox machine coords as a fallback
        # anchor so the trace lands at correct grid scale.
        overlay_row = tk.Frame(self, bg=DIALOG_BG)
        overlay_row.pack(pady=(0, 5))
        if self._camera_capture_available():
            self._show_camera_var = tk.BooleanVar(value=False)
            self._show_camera_chk = tk.Checkbutton(
                overlay_row,
                text=_("Show live camera underneath"),
                variable=self._show_camera_var, bg=DIALOG_BG,
                font=("Helvetica", 9),
                command=self._on_camera_overlay_toggle)
            self._show_camera_chk.pack(side="left")

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text=_("Clear"), command=self.on_clear, width=10).pack(side="left", padx=5)
        # "Get from camera" — only shown when OpenCV is available AND a
        # camera calibration file exists on disk.
        if self._camera_capture_available():
            tk.Button(btn_frame, text=_("Get from camera"),
                      command=self.on_get_from_camera,
                      width=15).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Submit"), command=self.on_submit, width=10,
                  font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Cancel"), command=self.on_cancel, width=10).pack(side="left", padx=5)

    def _camera_capture_available(self):
        """Return True if camera capture is wired up AND calibrated AND
        the user has opted into machine integration AND the platform
        supports it. Gating mirrors the Pad Maker tab: 'Get from camera'
        hides entirely when the experimental machine toggle is off,
        even if a stale calibration file is on disk, and the whole
        machine integration is Windows-only (see PadSVGGeneratorApp.
        _machine_enabled)."""
        import sys
        if sys.platform != 'win32':
            return False
        if not self._settings or not self._settings.get(
                "experimental_machine_menu", False):
            return False
        try:
            import camera_capture
        except ImportError:
            return False
        if not camera_capture.HAS_OPENCV:
            return False
        return camera_capture.load_calibration(
            camera_capture.default_calibration_path()) is not None

    def on_get_from_camera(self):
        """Open CameraCaptureDialog and adopt the returned polygon."""
        try:
            import camera_capture
        except ImportError:
            return
        cam_idx = self._resolve_camera_index()
        if cam_idx is None:
            messagebox.showerror(_("No Camera"),
                                  _("No cameras detected."), parent=self)
            return

        dlg = _pkg.CameraCaptureDialog(
            self, camera_index=cam_idx,
            calibration_path=camera_capture.default_calibration_path(),
            settings=self._settings)
        if not dlg.result_polygon_mm:
            return

        # CameraCaptureDialog returns the polygon in ABSOLUTE machine-mm
        # (Y-up). Store the polygon at absolute machine grid coords so
        # it appears in the grid view at the same place it sits on the
        # bed (matching what the user saw on the camera screen, and
        # matching where a hand-traced polygon with the live overlay
        # would land). The overlay anchors at machine (0, 0) — grid
        # origin = bed origin — so machine-mm divided by mm_per_unit
        # lands directly on the grid.
        machine_polygon = list(dlg.result_polygon_mm)
        self._camera_anchor_mm = (0.0, 0.0)
        scale_to_unit = 1.0 / 25.4 if self.unit == "in" else 1.0 / 10.0
        captured = [(x * scale_to_unit, y * scale_to_unit)
                     for (x, y) in machine_polygon]
        # Downsample to MAX_POINTS if the contour is denser. Picks every
        # k-th vertex (uniform sampling along the contour).
        if len(captured) > self.MAX_POINTS:
            step = len(captured) / self.MAX_POINTS
            captured = [captured[int(i * step)] for i in range(self.MAX_POINTS)]
        self.points = [(max(0, min(self.grid_size, x)),
                        max(0, min(self.grid_size, y)))
                        for (x, y) in captured]
        self.polygon_closed = True
        self._redraw_polygon()
        # Checkbox is already enabled (overlay works pre- or post-
        # capture); no state change needed here.
        self.status_var.set(
            _("Captured {n} points from camera").format(n=len(self.points)))

    def _resolve_camera_index(self):
        """Same resolver pattern as the dialogs in main.py: persisted
        override first, then Falcon-name heuristic, then last enumerated."""
        try:
            import camera_capture
        except ImportError:
            return None
        if not camera_capture.HAS_OPENCV:
            return None
        if self._settings:
            override = self._settings.get("camera_index_override")
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

    # ------------------------------------------------------------------
    # Live camera overlay
    # ------------------------------------------------------------------

    def _on_camera_overlay_toggle(self):
        """Checkbox handler — start/stop the live underlay."""
        if self._show_camera_var.get():
            self._start_camera_overlay()
        else:
            self._stop_camera_overlay()

    def _start_camera_overlay(self):
        # No anchor required up front — when there isn't a capture
        # anchor yet, _render_overlay falls back to the camera image's
        # own machine-coord bounding box as an anchor (so the trace
        # lands at correct grid scale even pre-capture).
        if self._cap is not None:
            return  # already running
        # Mark the polygon as machine-coord-anchored. The overlay
        # always renders at machine (0, 0), so any clicks the user
        # makes correspond to absolute machine positions. Setting
        # _camera_anchor_mm here (in addition to _overlay_used)
        # routes the polygon through main.py's _adopt_camera_polygon
        # path on submit — same path camera-captured polygons take —
        # so the camera-polygon safety inset is applied to traced
        # polygons too. A skip-the-anchor implementation would mean
        # only camera captures got the inset; hand-traced shapes
        # would cut right up to the user's click points, which is
        # less safe for irregular leather edges.
        self._overlay_used = True
        if self._camera_anchor_mm is None:
            self._camera_anchor_mm = (0.0, 0.0)
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
            self._calibration = camera_capture.load_calibration(
                camera_capture.default_calibration_path())
            if not self._calibration:
                raise RuntimeError(_("No camera calibration on disk."))
        except (ImportError, RuntimeError) as e:
            messagebox.showerror(_("Camera Overlay Unavailable"),
                                  str(e), parent=self)
            self._show_camera_var.set(False)
            return
        cam_idx = self._resolve_camera_index()
        if cam_idx is None:
            messagebox.showerror(_("No Camera"),
                                  _("No cameras detected."), parent=self)
            self._show_camera_var.set(False)
            return
        self._camera_index = cam_idx
        self.status_var.set(_("Opening camera for overlay..."))

        def _worker():
            try:
                # calibration_present=True: this overlay warps the picture
                # through the calibration it just loaded, so only the
                # Falcon-named camera may stand in for the saved one.
                cap, used, _frame = camera_capture.open_camera_or_fallback(
                    cam_idx, (self._settings or {}).get("camera_index_override"),
                    calibration_present=True)
                self._camera_index = used
                notice = camera_capture.fallback_notice(cam_idx, used) if used != cam_idx else None
                self.after(0, self._camera_overlay_ready, cap, notice)
            except Exception as e:
                self.after(0, self._camera_overlay_failed, e)

        start_worker(_worker, name='polydraw-overlay')

    def _camera_overlay_ready(self, cap, notice=None):
        # User may have toggled off / cancelled / closed in the meantime.
        # (Cancel and Submit don't clear the toggle, so a closed window
        # has to be checked separately or the handle is never released.)
        if (not self.winfo_exists() or not self._show_camera_var
                or not self._show_camera_var.get()):
            try:
                cap.release()
            except Exception:
                pass
            return
        self._cap = cap
        # `notice`: the overlay runs on another camera than the saved one
        # (camera_capture.fallback_notice); never silently.
        self.status_var.set(notice or _("Live camera underlay on."))
        self._overlay_refresh_loop()

    def _camera_overlay_failed(self, err):
        self.status_var.set(_("Camera open failed: {e}").format(e=err))
        if self._show_camera_var is not None:
            self._show_camera_var.set(False)

    def _overlay_refresh_loop(self):
        if (self._cap is None
                or not self.winfo_exists()
                or self._show_camera_var is None
                or not self._show_camera_var.get()):
            return
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok = False
        if ok and frame is not None:
            try:
                self._render_overlay(frame)
            except Exception:
                # Don't kill the loop on a bad frame, and don't hide it:
                # the first failure logs its traceback, the rest are
                # counted and reported when the overlay stops.
                self._overlay_errors += 1
                if self._overlay_errors == 1:
                    mlog.exception(
                        "camera overlay: drawing a frame failed (further "
                        "failures are counted and reported when the "
                        "overlay stops)")
        self._overlay_after_id = self.after(
            self.OVERLAY_REFRESH_MS, self._overlay_refresh_loop)

    def _render_overlay(self, frame):
        """Warp the camera frame into canvas-pixel space using the same
        pixel→machine homography the capture used, plus the grid's
        machine anchor + canvas px-per-mm scale. Result lands UNDER
        the grid + polygon items via tag_lower."""
        import cv2
        import numpy as np
        mismatch = self._cam_mod.calibration_size_mismatch(
            self._calibration, frame)
        if mismatch:
            cw, ch, fw, fh = mismatch
            if not getattr(self, '_size_mismatch_logged', False):
                self._size_mismatch_logged = True
                mlog.warning("camera overlay: camera sends %dx%d, "
                             "calibrated at %dx%d", fw, fh, cw, ch)
            self.status_var.set(_(
                "This camera is sending {fw}×{fh} images, but the "
                "calibration was made at {cw}×{ch}. Recalibrate the "
                "camera (Options > Machine > Camera Calibration).").format(
                    fw=fw, fh=fh, cw=cw, ch=ch))
            return
        undist = self._cam_mod.undistort_frame(frame, self._calibration)
        h, w = undist.shape[:2]
        # Map four image corners through pixels_to_mm (which handles
        # schema-1 Y-flip compat internally), then through our
        # machine→canvas-pixel affine, to get the destination corners
        # for the perspective warp.
        src_corners = [(0.0, 0.0), (float(w), 0.0),
                       (float(w), float(h)), (0.0, float(h))]
        machine_corners = self._cam_mod.pixels_to_mm(
            src_corners, self._calibration)
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        scale = self.px_per_unit / mm_per_unit  # canvas-px per machine-mm
        # Anchor priority: explicit capture anchor wins; otherwise
        # fall back to machine (0, 0) — the bed front-left corner.
        # Canvas BL then shows the bed origin, matching the captured-
        # mode convention of "polygon BL at canvas BL" so the camera
        # view is positioned consistently regardless of whether a
        # polygon was captured first.
        #
        # An earlier fallback anchored at the camera image's own bbox
        # bottom-left, which on Matt's calibrated rig (camera sees
        # past the bed edges to ~-133mm in machine X) put the canvas
        # BL way off the bed and made the visible bed content shift
        # into the right half of the canvas — read as "scale off"
        # since the bed appeared smaller than expected.
        if self._camera_anchor_mm is not None:
            ax, ay = self._camera_anchor_mm
        else:
            ax, ay = 0.0, 0.0
        dst_corners = []
        for mx, my in machine_corners:
            cx = (mx - ax) * scale
            # Canvas Y is image-down; grid Y is up — invert here.
            cy = self.CANVAS_PX - (my - ay) * scale
            dst_corners.append((cx, cy))
        src_arr = np.array(src_corners, dtype=np.float32)
        dst_arr = np.array(dst_corners, dtype=np.float32)
        try:
            homography = cv2.getPerspectiveTransform(src_arr, dst_arr)
        except cv2.error:
            return
        warped = cv2.warpPerspective(
            undist, homography,
            (self.CANVAS_PX, self.CANVAS_PX),
            borderMode=cv2.BORDER_CONSTANT,
            borderValue=(255, 255, 255))
        rgb = cv2.cvtColor(warped, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        self._overlay_photo = self._PIL_ImageTk.PhotoImage(img)
        try:
            self.canvas.delete('camera_overlay')
            self.canvas.create_image(
                0, 0, image=self._overlay_photo, anchor='nw',
                tags='camera_overlay')
            # Send to the bottom so grid + polygon items stay visible.
            self.canvas.tag_lower('camera_overlay')
        except tk.TclError:
            pass

    def _stop_camera_overlay(self):
        if self._overlay_after_id is not None:
            try:
                self.after_cancel(self._overlay_after_id)
            except (tk.TclError, ValueError):
                pass
            self._overlay_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        try:
            self.canvas.delete('camera_overlay')
        except tk.TclError:
            pass
        if self._overlay_errors:
            mlog.warning("camera overlay: %d frames failed to draw",
                         self._overlay_errors)
            self._overlay_errors = 0
        self._overlay_photo = None

    def _draw_grid(self):
        """Draw the grid lines on the canvas."""
        self.canvas.delete("grid")

        # Draw grid lines (1 line per unit)
        # Highlight every 5th line for easier orientation (especially useful for cm grid)
        for i in range(self.grid_size + 1):
            px = i * self.px_per_unit
            # Use reddish color for every 5th line (0, 5, 10, 15, ...)
            if i % 5 == 0:
                line_color = "#CC9999"
            else:
                line_color = "#CCCCCC"
            # Vertical lines
            self.canvas.create_line(px, 0, px, self.CANVAS_PX, fill=line_color, tags="grid")
            # Horizontal lines
            self.canvas.create_line(0, px, self.CANVAS_PX, px, fill=line_color, tags="grid")

        # Draw axis labels - every 5 for inches, every 10 for cm (to avoid crowding)
        label_step = 5 if self.unit == "in" else 10
        for i in range(0, self.grid_size + 1, label_step):
            px = i * self.px_per_unit
            # X-axis labels (bottom) - label value = grid coordinate (1 square = 1 unit)
            self.canvas.create_text(px, self.CANVAS_PX - 5, text=str(i),
                                     font=("Helvetica", 8), anchor="s", tags="grid")
            # Y-axis labels (left) - invert Y so 0 is at bottom
            self.canvas.create_text(5, self.CANVAS_PX - px, text=str(i),
                                     font=("Helvetica", 8), anchor="w", tags="grid")

    def _grid_to_canvas(self, gx, gy):
        """Convert grid coordinates to canvas pixels. Y is inverted (0 at bottom)."""
        cx = gx * self.px_per_unit
        cy = self.CANVAS_PX - (gy * self.px_per_unit)
        return cx, cy

    def _canvas_to_grid(self, cx, cy):
        """Convert canvas pixels to grid coordinates.

        Returns float grid units (no longer snapped to integer
        intersections — a 1-inch grid spacing is too coarse for
        accurate scrap tracing). Camera-captured polygons also produce
        floats, so downstream nesting + machine-coord round-trip
        already handle non-integer points.

        Bounds-clamped to [0, grid_size].
        """
        gx = cx / self.px_per_unit
        gy = (self.CANVAS_PX - cy) / self.px_per_unit
        gx = max(0.0, min(float(self.grid_size), gx))
        gy = max(0.0, min(float(self.grid_size), gy))
        return gx, gy

    def _redraw_polygon(self):
        """Redraw all points and lines."""
        self.canvas.delete("polygon")

        if not self.points:
            return

        # Draw lines between points
        if len(self.points) > 1:
            canvas_points = [self._grid_to_canvas(p[0], p[1]) for p in self.points]
            for i in range(len(canvas_points) - 1):
                x1, y1 = canvas_points[i]
                x2, y2 = canvas_points[i + 1]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

            # If closed, draw line from last to first
            if self.polygon_closed:
                x1, y1 = canvas_points[-1]
                x2, y2 = canvas_points[0]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

        # Draw points
        for i, (gx, gy) in enumerate(self.points):
            cx, cy = self._grid_to_canvas(gx, gy)
            color = "green" if i == 0 else "red"
            self.canvas.create_oval(
                cx - self.POINT_RADIUS, cy - self.POINT_RADIUS,
                cx + self.POINT_RADIUS, cy + self.POINT_RADIUS,
                fill=color, outline="black", tags="polygon"
            )

    def _update_status(self):
        """Update the status label."""
        if self.polygon_closed:
            self.status_var.set(f"Shape closed ({len(self.points)} points). Click Submit or adjust points.")
        elif len(self.points) >= self.MAX_POINTS:
            self.status_var.set(f"Max points reached ({self.MAX_POINTS}). Click near first point to close.")
        else:
            remaining = self.MAX_POINTS - len(self.points)
            self.status_var.set(f"{len(self.points)} points. {remaining} remaining. Click near first point to close.")

    def on_canvas_click(self, event):
        """Handle click on canvas. The grid's machine anchor (if set
        via camera capture) is preserved across vertex edits — moving
        a point in grid coords still has a valid machine-coord
        translation, so the polygon round-trips to machine on submit."""
        if self.polygon_closed:
            # If already closed, check if clicking on a point to remove it
            clicked_idx = self._get_clicked_point_index(event.x, event.y)
            if clicked_idx is not None:
                self.points.pop(clicked_idx)
                self.polygon_closed = False
                self._redraw_polygon()
                self._update_status()
            return

        # Check if clicking near first point to close the polygon (before point removal)
        if len(self.points) >= 3:
            first_cx, first_cy = self._grid_to_canvas(self.points[0][0], self.points[0][1])
            dist = ((event.x - first_cx) ** 2 + (event.y - first_cy) ** 2) ** 0.5
            if dist < self.CLOSE_THRESHOLD:
                self.polygon_closed = True
                self._redraw_polygon()
                self._update_status()
                return

        # Check if clicking on an existing point (to remove it)
        clicked_idx = self._get_clicked_point_index(event.x, event.y)
        if clicked_idx is not None:
            self.points.pop(clicked_idx)
            self._redraw_polygon()
            self._update_status()
            return

        # Get grid coordinates
        gx, gy = self._canvas_to_grid(event.x, event.y)

        # Check if we can add more points
        if len(self.points) >= self.MAX_POINTS:
            return

        # Check if point already exists at this location
        if (gx, gy) in self.points:
            return

        # Add the point
        self.points.append((gx, gy))
        self._redraw_polygon()
        self._update_status()

    def _get_clicked_point_index(self, cx, cy):
        """Return index of point near click, or None."""
        for i, (gx, gy) in enumerate(self.points):
            pcx, pcy = self._grid_to_canvas(gx, gy)
            dist = ((cx - pcx) ** 2 + (cy - pcy) ** 2) ** 0.5
            if dist <= self.POINT_RADIUS + 4:  # Small tolerance
                return i
        return None

    def on_clear(self):
        """Clear all points AND (if overlay is OFF) the machine
        anchor. If overlay is currently running, preserve the anchor
        so a retrace still has machine reference — the user is
        clearing the polygon but the same camera context still
        applies to the next set of clicks."""
        self.points = []
        self.polygon_closed = False
        if self._cap is None:
            # Overlay isn't running — no machine context applies to
            # the next polygon. Clear both signals.
            self._camera_anchor_mm = None
            self._overlay_used = False
        # Else: overlay is still showing camera content, leave
        # _camera_anchor_mm and _overlay_used as they are.
        self._redraw_polygon()
        self._update_status()

    def get_machine_polygon_mm(self):
        """Return absolute machine-Y-up polygon, reconstructed from the
        current grid points + the camera-set anchor. None if no anchor
        (purely drawn polygon — no machine reference)."""
        if self._camera_anchor_mm is None or not self.points:
            return None
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        ax, ay = self._camera_anchor_mm
        return [(gx * mm_per_unit + ax, gy * mm_per_unit + ay)
                for (gx, gy) in self.points]

    def get_polygon_lb_machine_mm(self):
        """Return the polygon's leftmost-lowest vertex in absolute
        machine coords (Y-up), or None if no machine reference exists.

        Used by Frame & Cut's "Try Auto Locate" — drives the laser
        head to this exact point so the user doesn't have to jog
        manually. Available whenever the polygon was either captured
        from the camera (anchor set) or hand-traced with the live
        overlay on (overlay_used set; the overlay anchors at machine
        (0, 0) so grid coords map 1:1 to machine coords). For purely
        hand-drawn polygons (no overlay, no capture) the grid coords
        are arbitrary — no machine reference — so None is returned
        and the auto-locate button hides.
        """
        if not self.points:
            return None
        if self._camera_anchor_mm is None and not self._overlay_used:
            return None
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        ax, ay = self._camera_anchor_mm or (0.0, 0.0)
        machine = [(gx * mm_per_unit + ax, gy * mm_per_unit + ay)
                    for (gx, gy) in self.points]
        return tuple(polygon_anchor_vertex(machine))

    def on_submit(self):
        """Submit the polygon."""
        if len(self.points) < 3:
            from tkinter import messagebox
            messagebox.showwarning(_("Not Enough Points"),
                                   _("Please draw at least 3 points to create a shape."),
                                   parent=self)
            return

        if not self.polygon_closed:
            from tkinter import messagebox
            messagebox.showwarning(_("Shape Not Closed"),
                                   _("Please close the shape by clicking near the first point."),
                                   parent=self)
            return

        # Release camera + cancel refresh timer before destroying the
        # dialog. get_machine_polygon_mm() is reconstructed on demand
        # from self.points + self._camera_anchor_mm, both of which
        # survive past destroy(), so the caller can still read the
        # absolute polygon after we close.
        self._stop_camera_overlay()
        # Return points as list of (x, y) tuples in grid units
        self.result = list(self.points)
        self.destroy()

    def on_cancel(self):
        """Cancel and close."""
        # Stop the camera overlay first — releases the cv2 cap handle
        # and cancels the refresh timer. Calling destroy() without
        # this would leak the camera.
        self._stop_camera_overlay()
        self.result = None
        # Drop the anchor so get_machine_polygon_mm doesn't return a
        # stale polygon after the user explicitly cancelled.
        self._camera_anchor_mm = None
        self.destroy()

    def get_polygon(self):
        """Return the polygon points or None if cancelled."""
        return self.result


class CameraCalibrationDialog(tk.Toplevel):
    """One-time camera calibration wizard.

    Shows the live camera feed at ~10 fps with detected ChArUco corners
    overlaid in green. User moves the calibration card to several
    positions on the bed; each frame where the board is detected gets
    captured. Once enough good frames are collected, the user clicks
    Calibrate; the dialog runs the calibration math and saves the result.

    Result:
        self.result = "saved" if calibration was successfully saved, else None.
    """

    LIVE_REFRESH_MS = 100   # ~10 fps preview
    TARGET_FRAMES = 12      # frames to collect for a robust calibration
    PREVIEW_W = 640
    PREVIEW_H = 480
    # Diagnostics (see _note_frame): one detection summary per window
    # instead of a line per frame, and the raw frame saved to the config
    # folder -- a few seconds after the camera goes live, then again
    # whenever the board hasn't been found for a while.
    SUMMARY_INTERVAL_S = 5.0
    SNAPSHOT_FIRST_S = 3.0
    SNAPSHOT_NO_DETECT_S = 10.0
    SNAPSHOT_MIN_GAP_S = 60.0
    SNAPSHOT_NAME = "last_calibration_frame.png"
    # The preview rewrites its progress line ~10x a second; any other
    # message gets this long on screen first (see _may_write_progress).
    STATUS_HOLD_S = 6.0
    # Calibrate needs this many captures with the card MOVED, after the
    # references. Fewer leaves lens distortion unconstrained: in a
    # simulated Falcon camera, 12 untouched captures mapped the bed edge
    # 21 mm wrong while reporting a 0.28 px RMS.
    MIN_MOVED_CAPTURES = 5

    def __init__(self, parent, camera_index, calibration_path,
                  falcon_port=None, settings=None):
        super().__init__(parent)
        self.title(_("Camera Calibration"))
        self.configure(bg=DIALOG_BG)
        # NOT transient(parent) — on Windows that strips the maximize
        # button. The dialog has a lot of content (instructions +
        # preview + jog cluster + buttons) that benefits from being
        # maximizable on small screens.
        self.grab_set()
        self.resizable(True, True)
        self.result = None

        self._camera_index = camera_index
        self._calibration_path = calibration_path
        self._falcon_port = falcon_port
        self._settings = settings or {}
        self._cap = None
        self._board = None
        self._detector_imports_ok = False
        self._captures = []       # list of (charuco_corners, charuco_ids)
        self._last_frame_shape = None
        self._image_size = (640, 480)
        self._latest_photo = None  # keep a reference so PIL doesn't GC it
        self._refresh_after_id = None
        self._cam_opening = False
        self._cam_open_reason = "open"
        # Live-preview diagnostics: see _note_frame / _maybe_snapshot.
        self._live_stats = None
        self._live_since = None
        self._last_good_t = None
        self._last_snapshot_t = None
        self._snapshot_this_open = False
        self._loop_errors = 0
        self._failed_reads = 0
        self._last_progress_status = None
        self._foreign_status = None
        self._foreign_status_t = 0.0
        self._port_lock = GcHeldRLock()   # a holder has the collector off: see threads.py
        self._win_size = (self.PREVIEW_W + 380, max(self.PREVIEW_H, 600))
        # Phase state: 'engrave' or 'capture'. Decided after the
        # persisted-offset check below — if SSC saved an offset from
        # a previous (possibly 2-hour) engrave, default to capture
        # phase so the user isn't asked to re-engrave a card that's
        # already on the bed.
        self._phase = None  # set below after offset restore
        # Capture-phase sub-state: 'reference' (card untouched, take
        # 1+ reference captures) → 'intrinsics' (card moved for
        # different poses). Tracked so calibrate_from_frames can pool
        # all reference captures for the machine-mm homography.
        self._capture_substate = 'reference'
        self._reference_count = 0
        # Has the laser been homed this dialog session? Required
        # before Engrave so MPos is meaningful — otherwise the
        # calibration's machine reference is whatever bogus value
        # Grbl booted with.
        self._engrave_homed = False
        # Linear-flow flags driving button enablement: Frame →
        # _framing_done → enables Engrave; Engrave → engrave offset
        # set → enables Next; Next → capture phase.
        self._framing_done = False
        # Engrave centres the card on the head's MPos when it's clicked.
        # A framing pass ended with Done finishes its loop and parks the
        # head back at the framed centre; one ended with the window's X
        # soft-resets mid-trace, which loses position, so
        # _on_frame_clicked then clears _engrave_homed and _framing_done.

        # Where the card was engraved on the bed, in machine-mm.
        # Either captured this session (set by _on_engrave_clicked
        # after a successful engrave) or restored from a previously-
        # persisted value (below).
        self._captured_engrave_offset = None

        # If the previous engrave's offset was saved (persisted to
        # settings on engrave success), restore it. Then the dialog
        # opens straight in capture mode — the user shouldn't have
        # to remember a "Skip — already engraved" step if the system
        # already knows the card was engraved at a specific offset.
        saved = self._settings.get(
            'camera_calibration_engrave_offset_mm')
        if (saved and isinstance(saved, (list, tuple))
                and len(saved) == 2):
            try:
                self._captured_engrave_offset = (
                    float(saved[0]), float(saved[1]))
            except (TypeError, ValueError):
                pass

        # Phase pick: capture if we have a saved offset (most users
        # will, after their first engrave), else engrave. Without a
        # Falcon connected, can't engrave anyway → capture.
        if self._captured_engrave_offset is not None or not falcon_port:
            self._phase = 'capture'
        else:
            self._phase = 'engrave'
        mlog.info("calibration dialog opened: phase=%s, saved engrave "
                  "offset=%s, camera #%s, laser port=%s, calibration "
                  "file=%s", self._phase, self._captured_engrave_offset,
                  camera_index, falcon_port, calibration_path)

        # Block system/display sleep for the lifetime of the dialog —
        # the engrave can run 45-60 minutes and Windows would
        # otherwise sleep the system and kill the USB serial mid-job.
        # Released in _on_close.
        try:
            import sleep_lock
            sleep_lock.prevent_sleep()
        except Exception as e:
            mlog.warning("calibration: could not keep the computer awake "
                         "during the calibration: %s", e)

        # Try to import deps. If they're missing we'll show an error and close.
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV is not available")
            self._board = camera_capture.make_charuco_board()
            self._detector_imports_ok = True
        except ImportError as e:
            tk.Label(self, text=_("Camera calibration requires opencv-python "
                                    "and Pillow:\n\n    pip install opencv-python Pillow\n\n"
                                    "Error: {e}").format(e=e),
                     bg=DIALOG_BG, fg="#a30000", padx=20, pady=20).pack()
            tk.Button(self, text=_("Close"),
                      command=self.destroy).pack(pady=(0, 15))
            return

        # ---- UI ----
        # Horizontal split layout: title bar at top, status at
        # bottom, then a body with the camera preview on the LEFT
        # (expands to fill) and the controls (instructions + buttons
        # + engrave-jog cluster) on the RIGHT (fixed width). Makes
        # better use of widescreen monitors and keeps everything
        # visible without scrolling.
        tk.Label(self, text=_("Camera Calibration"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold")
                 ).pack(side='top', pady=(10, 2))

        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(side='bottom', pady=(0, 8))

        body = tk.Frame(self, bg=DIALOG_BG)
        body.pack(side='top', fill='both', expand=True,
                   padx=10, pady=(0, 8))

        # RIGHT: control panel (instructions on top, jog/buttons
        # below), fixed width. Packed BEFORE the preview: the packer
        # serves widgets in packing order, so the panel always gets its
        # 340 px and the preview takes what's left. In the old order a
        # preview label showing no image (camera failed, or "Camera
        # opening...") asked for its 640x480 size in TEXT CHARACTERS,
        # about 4500x7200 px, and squeezed the panel to 1 px: no
        # Cancel, Switch camera or Reconnect when they were needed most.
        right_panel = tk.Frame(body, bg=DIALOG_BG, width=340)
        right_panel.pack(side='right', fill='y')
        right_panel.pack_propagate(False)  # honor the width=340

        # LEFT: camera preview, expands to fill available space.
        preview_frame = tk.Frame(body, bg='black')
        preview_frame.pack(side='left', fill='both', expand=True,
                            padx=(0, 8))
        self._preview_label = tk.Label(preview_frame, bg="#000000")
        self._preview_label.pack(fill='both', expand=True)

        # Instructions live in the right panel, narrower wraplength.
        self._instructions_var = tk.StringVar(value="")
        self._instructions_label = tk.Label(
            right_panel, textvariable=self._instructions_var,
            bg=DIALOG_BG, wraplength=325, justify="left",
            font=("Helvetica", 9))
        self._instructions_label.pack(side='top', anchor='w',
                                        padx=4, pady=(0, 8))

        # Bottom-row buttons in the right panel. Engrave jog cluster
        # also lives here (added by _enter_engrave_phase).
        self._btn_frame = tk.Frame(right_panel, bg=DIALOG_BG)
        self._btn_frame.pack(side='bottom', pady=(0, 8), fill='x')

        # Three-button linear flow: Frame → Engrave → Next. Each
        # enables the next as its prerequisite completes. Cancel is
        # always available.
        self._frame_btn = tk.Button(
            self._btn_frame, text=_("Frame"),
            command=self._on_frame_clicked,
            font=("Helvetica", 10), width=24)
        self._engrave_btn = tk.Button(
            self._btn_frame, text=_("Engrave"),
            command=self._on_engrave_clicked,
            font=("Helvetica", 10), width=24,
            state="disabled")  # enabled after framing done
        self._next_btn = tk.Button(
            self._btn_frame, text=_("Next →"),
            command=self._on_next_clicked,
            font=("Helvetica", 10, "bold"), width=24,
            bg="#4caf50", fg="white", activebackground="#388e3c",
            state="disabled")  # enabled after engrave done
        # "Engrave new card" — shown in CAPTURE phase to let the user
        # start over with a fresh piece of basswood (can't re-engrave
        # the same wood — each engrave is permanent, so a new card
        # means new material). Skip — already engraved button was
        # removed — it always used the (50,50) default which is wrong
        # for users who jogged before engraving.
        self._reengrave_btn = tk.Button(
            self._btn_frame, text=_("Engrave new card"),
            command=self._on_reengrave_clicked,
            font=("Helvetica", 9), width=16)

        # Capture-phase buttons. Label updates with substate.
        self._capture_btn = tk.Button(
            self._btn_frame, text=_("Capture reference"),
            command=self._on_capture_clicked,
            font=("Helvetica", 10, "bold"), width=20, state="disabled")
        # "Done with references — move card now" advances from the
        # reference substate (multiple captures of the untouched
        # card, pooled for the machine-mm homography) to the
        # intrinsics substate (card moved, used for distortion).
        self._done_refs_btn = tk.Button(
            self._btn_frame, text=_("Done with references"),
            command=self._on_done_refs_clicked,
            font=("Helvetica", 9), width=18, state="disabled")
        # Retake-last drops the most recent capture so the user can
        # redo a bad one (detection sketchy, lighting bad, etc.)
        # without losing everything else.
        self._retake_ref_btn = tk.Button(
            self._btn_frame, text=_("Retake last"),
            command=self._on_retake_last_clicked,
            font=("Helvetica", 9), width=12, state="disabled")
        self._reset_btn = tk.Button(
            self._btn_frame, text=_("Reset all"),
            command=self._on_reset_captures_clicked,
            font=("Helvetica", 9), width=10, state="disabled")
        self._calibrate_btn = tk.Button(
            self._btn_frame, text=_("Calibrate & Save"),
            command=self._on_calibrate_clicked,
            font=("Helvetica", 10), width=18, state="disabled")

        # Reconnect Camera — recovery if the USB camera dies mid-session
        # (the long engrave saturates USB and the cam can stall or drop;
        # cv2.VideoCapture has a stale handle after that, so we need
        # an explicit release+reopen to come back).
        self._reconnect_btn = tk.Button(
            self._btn_frame, text=_("Reconnect camera"),
            command=self._on_reconnect_camera_clicked,
            font=("Helvetica", 9), width=18)
        # Switch camera — fallback for when the auto-resolver picked
        # the wrong device (laptop webcam vs overhead Falcon cam).
        # Persists the chosen index so other camera-using dialogs also
        # default correctly afterward.
        self._switch_cam_btn = tk.Button(
            self._btn_frame, text=_("Switch camera"),
            command=self._on_switch_camera_clicked,
            font=("Helvetica", 9), width=18)

        # Manual-offset entry — escape hatch when an engrave succeeded
        # but the persisted offset was lost (e.g. older settings file
        # missing the DEFAULT_SETTINGS key, manual config edit, etc.)
        # and the card is still sitting on the bed at known coords.
        # Avoids requiring another 2-hour re-engrave to recover.
        self._manual_offset_btn = tk.Button(
            self._btn_frame, text=_("Enter offset manually..."),
            command=self._on_manual_offset_clicked,
            font=("Helvetica", 9), width=22)

        # Always-visible Cancel
        self._cancel_btn = tk.Button(
            self._btn_frame, text=_("Cancel"),
            command=self._on_close, width=10)

        # Wide default: preview width + right panel + padding.
        # Tall enough for the right panel's stacked controls.
        self.geometry("{}x{}".format(*self._win_size))
        self.minsize(900, 600)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Enter the initial phase. If no Falcon connected, skip
        # straight to capture phase (user runs engrave externally).
        if self._phase == 'engrave':
            self._enter_engrave_phase()
        else:
            self._enter_capture_phase()
        self.wait_window(self)

    def _enter_engrave_phase(self):
        """Show the engrave-phase UI: live camera preview + jog buttons
        + MPos display + "Engrave Centered Here" button. User homes the
        laser, jogs into the camera view, then clicks Engrave — card
        engraves around the head's current position."""
        self._phase = 'engrave'
        self._instructions_var.set(_(
            "Step 1 of 2 — engrave the calibration card.\n\n"
            "(Dimensions below are for the Falcon2 Pro 40W. Adjust "
            "for your machine. The engrave uses your basswood G-code "
            "preset — tune it under Tooling > Options.)\n\n"
            "1. Place a 12×12-inch basswood blank near the center "
            "of the bed and secure it. It cannot move between "
            "engraving and the Step 2 captures.\n"
            "2. Click Home Laser — head parks at the home corner.\n"
            "3. JOG THE HEAD TO ROUGHLY BED CENTER (X=200, Y=200, "
            "or X=~8\", Y=~8\"). The card is ~11\" square and "
            "engraves CENTERED on the head, so at the home corner "
            "the trace runs off the bed. Use 50mm step for fast "
            "travel, then 10mm or 1mm to fine-tune.\n"
            "4. Click Frame — low-power outline traces where the "
            "card will engrave (~11×11 inches). The outline follows "
            "your jog buttons in real time. Adjust until the trace "
            "sits comfortably on the basswood (~½\" margin all "
            "around). Click Stop on the framing window.\n"
            "5. Click Engrave to commit. (Takes a while.)"))
        self._status_var.set(_("Ready. Home the laser and jog into position."))
        self._preview_label.config(image="", text=_("Camera opening..."),
                                    fg="#888888", bg="#000000")
        for w in (self._capture_btn, self._done_refs_btn,
                   self._retake_ref_btn, self._reset_btn,
                   self._reengrave_btn, self._calibrate_btn,
                   self._next_btn):
            w.pack_forget()
        self._engrave_btn.config(text=_("Engrave"))
        # Build jog cluster + home + MPos display (once)
        if not hasattr(self, '_engrave_jog_built'):
            self._build_engrave_controls()
            self._engrave_jog_built = True
        # side='bottom' so the jog cluster pins above the bottom
        # button row (status_var) regardless of preview size. Without
        # this, on small screens the jog cluster falls off-screen
        # below the preview.
        self._engrave_jog_frame.pack(side='bottom', pady=(0, 8))
        # Stack vertically in the narrow (340px) right panel.
        self._frame_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._engrave_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._next_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reconnect_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._switch_cam_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._manual_offset_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._cancel_btn.pack(side="top", fill='x', padx=5, pady=2)
        # Reset to disabled, then enable based on current flow flags.
        # Without the reset, going from capture phase BACK to engrave
        # phase via Engrave-new-card would leave stale enabled state.
        # Frame is disabled after a completed engrave; coming back here
        # via "Engrave new card" must turn it on again or the new card
        # can never be framed (and so never engraved).
        self._frame_btn.config(state="normal")
        self._engrave_btn.config(state="disabled")
        self._next_btn.config(state="disabled")
        if self._captured_engrave_offset is not None:
            self._engrave_btn.config(state="normal")
            self._next_btn.config(state="normal")
        elif getattr(self, '_framing_done', False):
            self._engrave_btn.config(state="normal")
        # Open the camera early so the user can see the head as they jog.
        mlog.info("calibration: engrave phase")
        if self._cap is None and self._detector_imports_ok:
            self.after(50, self._open_camera_and_start, "engrave phase")
        self._start_mpos_polling()
        self._fit_height()

    def _build_engrave_controls(self):
        """Build the home/jog/MPos cluster used during engrave-phase
        positioning. Parented to the right-panel control area so it
        sits with the other controls, not floating in the dialog."""
        self._engrave_jog_frame = tk.LabelFrame(
            self._btn_frame.master, text=_("Position the head"),
            bg=DIALOG_BG, padx=8, pady=4)

        row = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        row.pack(anchor='w')
        tk.Button(row, text=_("Home Laser ($H)"),
                   command=self._on_engrave_home).pack(side='left', padx=4)
        self._mpos_var = tk.StringVar(value=_("MPos: --"))
        tk.Label(row, textvariable=self._mpos_var, bg=DIALOG_BG,
                  font=("Courier", 10)).pack(side='left', padx=10)

        step_row = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        step_row.pack(anchor='w', pady=(4, 2))
        tk.Label(step_row, text=_("Step (mm):"), bg=DIALOG_BG,
                  font=("Helvetica", 9)).pack(side='left')
        self._engrave_jog_step = tk.DoubleVar(value=10.0)
        for step in (1.0, 10.0, 50.0):
            tk.Radiobutton(
                step_row, text=str(step),
                variable=self._engrave_jog_step, value=step,
                bg=DIALOG_BG, font=("Helvetica", 9),
                highlightthickness=0).pack(side='left')

        btn_grid = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        btn_grid.pack(pady=2)
        tk.Button(btn_grid, text="↑", width=3,
                   command=lambda: self._engrave_jog(0, 1)
                   ).grid(row=0, column=1, padx=2, pady=1)
        tk.Button(btn_grid, text="←", width=3,
                   command=lambda: self._engrave_jog(-1, 0)
                   ).grid(row=1, column=0, padx=2, pady=1)
        tk.Button(btn_grid, text="→", width=3,
                   command=lambda: self._engrave_jog(1, 0)
                   ).grid(row=1, column=2, padx=2, pady=1)
        tk.Button(btn_grid, text="↓", width=3,
                   command=lambda: self._engrave_jog(0, -1)
                   ).grid(row=2, column=1, padx=2, pady=1)

    def _start_mpos_polling(self):
        """Poll Falcon's MPos every ~300ms while in engrave phase."""
        if self._phase != 'engrave' or not self._falcon_port:
            return
        self._mpos_after_id = self.after(300, self._poll_mpos)

    def _poll_mpos(self):
        """Read the head position for the MPos label, on a worker thread.

        Quick connect/disconnect per read: a persistent sender would tie
        up the port that Home, Jog, Frame and Engrave each open. The read
        used to run on the Tk thread, and FalconSender.connect() alone
        sleeps 0.3 s, so the preview and the jog buttons froze for about
        a third of every second while positioning the head. The worker
        skips its read whenever a button holds the port (_holds_port)."""
        if self._phase != 'engrave' or not self._falcon_port:
            return

        def _worker():
            st = None
            if self._port_lock.acquire(blocking=False):
                try:
                    import falcon_sender
                    s = falcon_sender.FalconSender(port=self._falcon_port)
                    try:
                        s.connect()
                        st = s.get_status(timeout=0.2)
                    finally:
                        s.disconnect()
                except Exception:
                    st = None
                finally:
                    self._port_lock.release()
            try:
                self.after(0, self._mpos_polled, st)
            except (RuntimeError, tk.TclError):
                pass  # app shutting down

        start_worker(_worker, name='camcal-mpos')

    def _mpos_polled(self, st):
        if not self.winfo_exists():
            return
        if st and st.get('mpos'):
            self._mpos_var.set(
                _("MPos: X={x:7.2f}  Y={y:7.2f}").format(
                    x=st['mpos'][0], y=st['mpos'][1]))
        if self._phase == 'engrave':
            self._mpos_after_id = self.after(800, self._poll_mpos)

    @_holds_port
    def _on_engrave_home(self):
        """Send $H standalone, with a small "Homing..." indicator."""
        if not self._falcon_port:
            return
        try:
            import falcon_sender
        except ImportError:
            return
        if not messagebox.askyesno(
                _("Home Laser?"),
                _("Send the laser home? The head will travel to the "
                  "home corner at full speed."), parent=self):
            return
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        self._status_var.set(_("Homing..."))
        self.update_idletasks()
        result = {}

        def _worker():
            try:
                ok, m = sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = m
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        t = start_worker(_worker, name='camcal-home')
        # after()-poll until done. wait_variable pumps the Tk event
        # loop without the WM_DELETE_WINDOW destroy-root-mid-loop
        # hazard of `while t.is_alive(): self.update()`.
        done_var = tk.BooleanVar(value=False)

        def _check():
            if t.is_alive():
                self.after(50, _check)
            else:
                done_var.set(True)

        self.after(50, _check)
        self.wait_variable(done_var)
        try:
            sender.disconnect()
        except Exception:
            pass
        if result.get('ok'):
            self._engrave_homed = True
            self._status_var.set(_("Homed. Now jog into the camera view."))
        else:
            self._status_var.set(
                _("Home failed: {m}").format(m=result.get('msg', '?')))

    @_holds_port
    def _engrave_jog(self, dx_sign, dy_sign):
        """Send a relative jog via Grbl $J=."""
        if not self._falcon_port:
            return
        try:
            import falcon_sender
        except ImportError as e:
            mlog.warning("calibration: jog failed: %s", e)
            return
        s = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            s.connect()
            try:
                step = float(self._engrave_jog_step.get())
            except Exception:
                step = 10.0
            s.jog(x=dx_sign * step, y=dy_sign * step,
                   feed=3000, relative=True)
        except Exception as e:
            mlog.warning("calibration: jog failed: %s", e)
        finally:
            s.disconnect()   # a jog that raised used to leave the port open (2026-11-10)

    def _enter_capture_phase(self):
        """Switch to live-preview + ChArUco capture mode."""
        self._phase = 'capture'
        self._instructions_var.set(_(
            "Step 2 of 2 — calibration captures.\n\n"
            "FIRST capture(s): board must stay at its engraved "
            "position. This anchors camera to machine coords; any "
            "shift here offsets every future scrap capture by that "
            "amount.\n\n"
            "Click 'Done with references' after 1+ high quality "
            "reference shots, then move the card to different spots "
            "between the remaining captures so the math can solve "
            "lens distortion. Watch the green-corner overlay; click "
            "Capture when the board is fully or mostly detected. "
            "Play with lighting if necessary.\n\n"
            "Target: {n} good captures total."
        ).format(n=self.TARGET_FRAMES))
        # Stop MPos polling and hide engrave-phase controls.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None
        for w in (self._frame_btn, self._engrave_btn, self._next_btn,
                   self._reconnect_btn, self._manual_offset_btn):
            w.pack_forget()
        if hasattr(self, '_engrave_jog_frame'):
            self._engrave_jog_frame.pack_forget()
        self._status_var.set(_("Opening camera ..."))
        self._capture_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._done_refs_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._retake_ref_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reset_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reengrave_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._calibrate_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reconnect_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._switch_cam_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._cancel_btn.pack(side="top", fill='x', padx=5, pady=2)
        # Always-visible Reconnect — the camera can die during the long
        # engrave or at dialog-transition. Without an explicit reopen
        # path the user has to cancel and lose any in-memory captures.
        # Kick off camera open + live refresh (may already be open
        # from engrave phase — _open_camera_and_start handles that: it does nothing when the camera is already open).
        mlog.info("calibration: capture phase (%d captures so far)",
                  len(self._captures))
        self.after(50, self._open_camera_and_start, "capture phase")
        self._fit_height()

    def _fit_height(self):
        """Grow the window (never shrink it) until the right panel's
        controls fit. Engrave phase stacks the instructions, the jog
        cluster and six buttons; at the fixed 600 px the jog cluster got
        1 px and Cancel fell off the bottom, and display scaling above
        100 % makes everything taller still. Capped at the screen."""
        try:
            self.update_idletasks()
            panel = self._btn_frame.master
            need = sum(w.winfo_reqheight() for w in panel.pack_slaves())
            need += 12 * len(panel.pack_slaves())  # pack padding
            need += sum(w.winfo_reqheight() for w in self.pack_slaves()
                        if w is not panel.master)  # title + status
            need += 40
            screen = self.winfo_screenheight() - 80  # taskbar, title bar
            target = min(need, screen)
            w, h = self._win_size
            if target > h:
                self._win_size = (w, target)
                self.geometry(f"{w}x{target}")
            self.minsize(900, min(max(600, target), screen))
        except tk.TclError:
            pass

    def _regrab(self):
        """FalconRunDialog takes the input grab while it runs, and Tk
        doesn't hand it back when that window closes: after Frame or
        Engrave this dialog stopped being modal and the main window
        took clicks in the middle of a calibration."""
        try:
            if self.winfo_exists():
                self.grab_set()
        except tk.TclError:
            pass

    def _engrave_geometry(self):
        """Return (cols, rows, sq, board_w, board_h, border, label_h)
        — the card-layout constants the engrave + framing flows both
        need."""
        import camera_capture as _cam
        cols, rows = _cam.CHARUCO_COLS, _cam.CHARUCO_ROWS
        sq = _cam.CHARUCO_SQUARE_MM
        return (cols, rows, sq, cols * sq, rows * sq, 10.0, 8.0)

    @_holds_port
    def _engrave_check_position_safe(self):
        """Read current MPos and verify the engrave area would fit on
        the bed. Returns True if safe, else False (with an error
        already shown). Avoids ALARM:3 from soft-limit trips when
        the user clicks Frame with the head still at home corner.

        A check that cannot be made does not pass (2026-11-10): until
        then a connect or status failure returned True and the bed-bounds
        check was skipped silently, so a head left at the home corner
        would have engraved off the bed. Now the position must be read,
        or the engrave/frame is refused with a box that says so."""
        try:
            import falcon_sender
        except ImportError as e:
            return self._position_unknown(str(e))
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            st = sender.get_status(timeout=0.5)
        except Exception as e:
            return self._position_unknown(str(e) or type(e).__name__)
        finally:
            sender.disconnect()
        if not st or not st.get('mpos'):
            return self._position_unknown(_("no status report"))
        hx, hy = st['mpos'][0], st['mpos'][1]
        cols, rows, sq, board_w, board_h, border, label_h = \
            self._engrave_geometry()
        ox = hx - border - board_w / 2.0
        oy = hy - border - label_h - board_h / 2.0
        fx = ox + board_w + 2 * border
        fy = oy + board_h + 2 * border + label_h
        bed_x_max = float(
            self._settings.get("laser_bed_x_max", 400.0))
        bed_y_max = float(
            self._settings.get("laser_bed_y_max", 415.0))
        if (ox < 0 or oy < 0
                or fx > bed_x_max + 1.0 or fy > bed_y_max + 1.0):
            mlog.warning("card would run off the bed: head (%.1f, %.1f), "
                         "card (%.0f, %.0f) to (%.0f, %.0f), bed %.0fx%.0f",
                         hx, hy, ox, oy, fx, fy, bed_x_max, bed_y_max)
            messagebox.showerror(
                _("Card Would Run Off Bed"),
                _("With head at machine (X={x:.1f}, Y={y:.1f}), the "
                  "card would engrave at ({ox:.0f},{oy:.0f}) to "
                  "({fx:.0f},{fy:.0f}) — outside the {bx:.0f}×{by:.0f}"
                  " bed.\n\n"
                  "Jog the head further from the home corner so the "
                  "whole card area fits on the bed (and on basswood).\n\n"
                  "Need head at roughly X={midx:.0f}-{maxx:.0f}, "
                  "Y={midy:.0f}-{maxy:.0f}.").format(
                    x=hx, y=hy, ox=ox, oy=oy, fx=fx, fy=fy,
                    bx=bed_x_max, by=bed_y_max,
                    midx=border + board_w / 2.0,
                    maxx=bed_x_max - border - board_w / 2.0,
                    midy=border + label_h + board_h / 2.0,
                    maxy=bed_y_max - border - board_h / 2.0,
                ), parent=self)
            return False
        return True

    def _position_unknown(self, reason):
        """The bed-bounds check could not read the head's position: say
        so, log it, and refuse (returns False for the caller)."""
        mlog.warning("bed-bounds check: the head position could not be read "
                     "(%s); engrave/frame refused", reason)
        messagebox.showerror(
            _("Head Position Unknown"),
            _("The laser's position could not be read ({reason}), so "
              "whether the card fits on the bed cannot be checked.\n\n"
              "Check the connection (Options > Machine > Test Connection), "
              "then try again.").format(reason=reason), parent=self)
        return False

    def _engrave_check_ready(self):
        """Common pre-flight for Frame / Engrave: Falcon connected,
        laser homed. Returns True if all good, else False (with the
        appropriate error/prompt already shown). Sets self._engrave_homed
        based on user choice."""
        if not self._falcon_port:
            messagebox.showerror(
                _("Falcon Not Detected"),
                _("Can't operate — no Falcon connection."), parent=self)
            return False
        if not self._engrave_homed:
            choice = messagebox.askyesnocancel(
                _("Home Laser First?"),
                _("The laser needs to be homed so machine coordinates "
                  "have a known reference. Without homing, the "
                  "engrave's position is bogus.\n\n"
                  "Yes — home now (head moves to home corner).\n"
                  "No — I already homed this session (LightBurn / "
                  "physical button / before opening SSC).\n"
                  "Cancel — back out."),
                parent=self)
            if choice is None:
                return False
            if choice:
                self._on_engrave_home()
                if not self._engrave_homed:
                    return False
            else:
                self._engrave_homed = True
        return True

    @_holds_port
    def _on_frame_clicked(self):
        """Run a continuous framing loop that re-reads MPos each
        iteration — the trace FOLLOWS the head as the user jogs
        with the in-dialog buttons. User clicks Stop or the X to
        end framing, then can verify alignment and click Engrave."""
        if not self._engrave_check_ready():
            return
        if not self._engrave_check_position_safe():
            return
        try:
            import falcon_sender
            from gcode_engine import generate_framing_gcode
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Frame needs pyserial: {e}"
                                    ).format(e=e), parent=self)
            return
        cols, rows, sq, board_w, board_h, border, label_h = \
            self._engrave_geometry()

        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return

        def _framing_provider():
            st = sender.get_status(timeout=0.3)
            if not st or not st.get('mpos'):
                # Fall back to a default near home rather than crash
                hx, hy = 100.0, 100.0
            else:
                hx, hy = st['mpos'][0], st['mpos'][1]
            ox = hx - border - board_w / 2.0
            oy = hy - border - label_h - board_h / 2.0
            fx = ox + board_w + 2 * border
            fy = oy + board_h + 2 * border + label_h
            from config import get_framing_power_s
            # G92.1 first: the outline is computed from MPos (machine
            # coordinates) but G-code moves in work coordinates, and a
            # G92 left by an earlier Frame & Cut in the same power-on
            # would shift the frame by that offset.
            return ['G92.1'] + generate_framing_gcode(
                ox, oy, fx, fy,
                # The calibration card is engraved on basswood, so frame
                # at basswood's power rather than the global fallback.
                power_s=get_framing_power_s('basswood', self._settings),
                feed=self._settings.get("laser_framing_feed", 2000),
                return_to_origin=False,  # absolute coords; don't send
                                          # head to machine (0, 0) (=home
                                          # corner) between iterations —
                                          # that loses user's jog edits.
                # Park at the head's original MPos so the next iter
                # produces the same bbox (and so the engrave-phase
                # MPos read after Done matches what the user jogged
                # to). The card geometry is slightly asymmetric in Y
                # (label adds ~4mm offset) — parking at hx, hy keeps
                # MPos invariant across iterations.
                park_xy=(hx, hy),
            )

        # Stop MPos polling while the streamer owns the port.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None

        mlog.info("calibration: framing the card outline (power S%s)",
                  self._settings.get("laser_framing_power_by_material", {})
                  .get("basswood"))
        fdlg = None
        try:
            fdlg = _pkg.FalconRunDialog(
                self, sender, _framing_provider(),
                title=_("Framing — jog to position, click Done when "
                         "alignment looks right"),
                loop=True, gcode_provider=_framing_provider,
                show_pause_resume=False, show_cut_button=False,
                stop_needs_confirm=False)
        finally:
            try:
                sender.disconnect()
            except Exception:
                pass
            self._regrab()
            self._start_mpos_polling()
            # Recycle the camera handle. Empirically the cv2.VideoCapture
            # frame stream stalls (frozen preview) when a long-lived
            # modal FalconRunDialog closes — probably a Windows
            # DirectShow / USB-focus quirk we can't suppress at the
            # cv2 layer. Cheaper to just reopen than diagnose.
            self._auto_recycle_camera()
        # Done ends the loop gracefully ("complete") with the head parked
        # back at the framed centre. The window's X soft-resets the laser
        # mid-trace instead, and a reset during motion loses the machine
        # position: until v2.86 that still counted as framed, and Engrave
        # then centred the card wherever the head had stopped, using a
        # position Grbl no longer knew.
        if fdlg is None or fdlg._final_reason != "complete":
            mlog.warning("calibration framing ended without finishing "
                         "(%s); laser position is no longer trusted",
                         None if fdlg is None else fdlg._final_reason)
            self._framing_done = False
            self._engrave_homed = False
            try:
                self._engrave_btn.config(state="disabled")
                self._status_var.set(_(
                    "Framing was stopped mid-trace, so the laser lost its "
                    "position. Home the laser, then Frame again."))
            except tk.TclError:
                pass
            return
        # Framing finished: the user is done positioning. They may still
        # Frame again or jog more; Engrave uses the head's position then.
        self._framing_done = True
        if self._phase == 'engrave':
            try:
                self._engrave_btn.config(state="normal")
                self._status_var.set(
                    _("Framing done. Click Engrave when ready."))
            except tk.TclError:
                pass

    def _on_next_clicked(self):
        """Engrave-phase Next button: advance to capture phase. Only
        active after the engrave has completed and saved an offset."""
        if self._captured_engrave_offset is None:
            messagebox.showerror(
                _("Engrave first"),
                _("No engrave offset on file yet. Click Engrave (and "
                  "wait for it to finish) before moving to captures."),
                parent=self)
            return
        self._enter_capture_phase()

    @_holds_port
    def _on_engrave_clicked(self):
        """Commit to the engrave at the head's current position.
        Confirmation popup ("are you sure?"). If yes, ~45-60 minute
        engrave runs. If no, returns to the engrave-phase dialog
        where the user can Frame again or jog."""
        if not self._engrave_check_ready():
            return
        if not self._engrave_check_position_safe():
            return
        try:
            import falcon_sender
            from gcode_engine import generate_calibration_card_gcode
            import camera_capture as _cam
        except ImportError as e:
            messagebox.showerror(
                _("Missing Dependency"),
                _("Engrave needs pyserial:\n\n{e}").format(e=e),
                parent=self)
            return

        # Read the head's current MPos to center the engrave around it.
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
            status = sender.get_status(timeout=0.5)
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        if not status or not status.get('mpos'):
            messagebox.showerror(
                _("No machine position"),
                _("Couldn't read MPos. Home the laser first."),
                parent=self)
            sender.disconnect()
            return
        head_x, head_y = status['mpos'][0], status['mpos'][1]

        # Compute offset so the BOARD's center lands at (head_x, head_y).
        # Card layout: bottom-left at (offset_x, offset_y), board area
        # at (offset_x + border, offset_y + border + label_h) to
        # (offset_x + border + board_w, offset_y + border + label_h + board_h).
        border = 10.0
        label_h = 8.0
        cols, rows = _cam.CHARUCO_COLS, _cam.CHARUCO_ROWS
        sq = _cam.CHARUCO_SQUARE_MM
        board_w = cols * sq
        board_h = rows * sq
        offset_x = head_x - border - board_w / 2.0
        offset_y = head_y - border - label_h - board_h / 2.0

        card_far_x = offset_x + board_w + 2 * border
        card_far_y = offset_y + board_h + 2 * border + label_h
        if not messagebox.askyesno(
                _("Start Engrave?"),
                _("Are you sure? Card will engrave centered on machine "
                  "(X={x:.0f}, Y={y:.0f}), covering ({ox:.0f}, "
                  "{oy:.0f}) to ({fx:.0f}, {fy:.0f}).\n\n"
                  "Takes a while. The cover must stay closed "
                  "for the laser to fire.\n\n"
                  "If you haven't framed yet to verify placement, "
                  "click No and run Frame first.").format(
                    x=head_x, y=head_y,
                    ox=offset_x, oy=offset_y,
                    fx=card_far_x, fy=card_far_y),
                parent=self):
            sender.disconnect()
            return

        # Generate G-code to a temp file (no $H — head's already homed
        # and at the user-chosen jog position; G-code uses absolute
        # coords from there). Read as a list of lines for the streamer.
        import tempfile
        import os as _os
        with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False,
                                          mode='w') as tmp:
            tmp_path = tmp.name
        try:
            generate_calibration_card_gcode(
                tmp_path,
                cols=cols, rows=rows,
                square_mm=sq, marker_mm=_cam.CHARUCO_MARKER_MM,
                settings=self._settings,
                offset_x_mm=offset_x, offset_y_mm=offset_y,
                home_first=False)  # already homed
            with open(tmp_path, 'r') as f:
                gcode_lines = f.read().splitlines()
        finally:
            try:
                _os.remove(tmp_path)
            except OSError:
                pass

        # Stop polling MPos while the stream owns the port.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None

        bw = (self._settings.get("gcode_settings") or {}).get("basswood") or {}
        mlog.info("engraving the calibration card: head (%.2f, %.2f), card "
                  "(%.2f, %.2f) to (%.2f, %.2f), %d G-code lines, basswood "
                  "speed=%s power=%s passes=%s spacing=%s", head_x, head_y,
                  offset_x, offset_y, card_far_x, card_far_y,
                  len(gcode_lines), bw.get("filled_engraving_speed"),
                  bw.get("filled_engraving_power"),
                  bw.get("filled_engraving_passes"),
                  bw.get("filled_line_spacing"))
        try:
            run_dlg = _pkg.FalconRunDialog(
                self, sender, gcode_lines,
                title=_("Engraving Calibration Card"))
            mlog.info("calibration card engrave ended: %s",
                      run_dlg._final_reason)
            self._regrab()
            # Run dialog closed (regardless of success/stop/error).
            # Recycle the camera before the next phase needs it — see
            # _auto_recycle_camera for the rationale.
            self._auto_recycle_camera()
            if run_dlg._final_reason != "complete":
                # Engrave stopped or errored mid-flight — the head
                # is now at some weird mid-trace MPos that's no
                # longer the planned card center. Force a re-frame
                # (with fresh basswood, presumably) before allowing
                # another engrave.
                self._framing_done = False
                try:
                    self._engrave_btn.config(state="disabled")
                    self._status_var.set(
                        _("Engrave stopped. Put fresh basswood, "
                          "re-Frame, then Engrave."))
                except tk.TclError:
                    pass
                self._start_mpos_polling()
                return
        finally:
            try:
                sender.disconnect()
            except Exception:
                pass

        # Engrave done — remember the offset for the calibration solver
        # AND persist it to settings so a closed/reopened dialog or
        # crashed-then-relaunched SSC can pick up where we left off
        # (the engrave is the most expensive step; losing the offset
        # to a UI hiccup wastes ~2 hours of basswood).
        self._captured_engrave_offset = (offset_x, offset_y)
        mlog.info("engrave offset saved: (%.3f, %.3f)", offset_x, offset_y)
        try:
            self._settings['camera_calibration_engrave_offset_mm'] = [
                offset_x, offset_y]
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            mlog.exception("could not save the engrave offset")  # never block on settings save
        # Activate Next — but DON'T auto-transition. User clicks Next
        # explicitly. Prevents losing the offset to a dialog-close
        # accident, and lets the user verify the engrave looks right
        # before moving on. Disable Frame + Engrave so user can't
        # accidentally re-engrave at a different MPos (would mismatch
        # the saved offset).
        try:
            self._frame_btn.config(state="disabled")
            self._engrave_btn.config(state="disabled")
            self._next_btn.config(state="normal")
            self._status_var.set(
                _("Engrave complete. Click Next when ready for "
                  "calibration captures."))
        except tk.TclError:
            pass
        messagebox.showinfo(
            _("Card Engraved — DO NOT TOUCH"),
            _("Card engraved successfully.\n\n"
              "⚠ DO NOT MOVE THE CARD. ⚠\n\n"
              "Click Next to start calibration captures. The first "
              "capture becomes the position reference linking the "
              "camera to machine coords — if the card moves before "
              "then, the calibration is off."), parent=self)

    def _on_reengrave_clicked(self):
        """Capture-phase button: discard the saved engrave offset and
        return to engrave phase so the user can engrave a NEW card
        on fresh basswood. Can't re-engrave the same wood — each
        engrave is permanent.

        Also clears any captures already taken in this session
        (different engrave = different machine-coord reference, so
        old captures would be tied to the wrong position).
        """
        if not messagebox.askyesno(
                _("Engrave a New Card?"),
                _("Discard the existing calibration session and start "
                  "over with a fresh basswood blank?\n\n"
                  "You CAN'T re-engrave the same wood — each engrave "
                  "is permanent. Use a new piece of basswood.\n\n"
                  "Current captures ({n}) will be discarded."
                  ).format(n=len(self._captures)),
                parent=self):
            return
        mlog.info("Engrave new card: discarding %d captures and the saved "
                  "engrave offset %s", len(self._captures),
                  self._captured_engrave_offset)
        self._captures = []
        self._reference_count = 0
        self._capture_substate = 'reference'
        self._capture_btn.config(text=_("Capture reference"))
        for b in (self._done_refs_btn, self._retake_ref_btn,
                  self._reset_btn, self._calibrate_btn):
            b.config(state="disabled")
        self._captured_engrave_offset = None
        # Force fresh framing for the new card — new wood / new
        # position warrants re-positioning, so disable Engrave until
        # they re-frame.
        self._framing_done = False
        # Clear the persisted offset so the next dialog open lands
        # in engrave phase again (not auto-skipped to capture).
        try:
            self._settings.pop(
                'camera_calibration_engrave_offset_mm', None)
            from config import save_settings
            save_settings(self._settings)
        except Exception as e:
            mlog.warning("calibration: could not clear the saved engrave "
                         "offset from settings: %s", e)
        self._enter_engrave_phase()

    # ------------------------------------------------------------------
    def _open_camera_and_start(self, reason="open"):
        """Open the camera on a background thread so the dialog stays
        responsive during the 1-3 second cv2.VideoCapture warm-up.

        Does nothing while a camera is already open or opening. Engrave
        phase opens it so the user can watch the head, and until v2.86
        capture phase opened it a second time on top: two handles on one
        DirectShow device and two refresh loops, only one of which
        _stop_refresh could cancel. Paths that want a fresh handle
        (reconnect, switch, recycle) release ``self._cap`` first."""
        if self._cap is not None or self._cam_opening:
            mlog.info("camera #%s: %s -- already %s, not reopening",
                      self._camera_index, reason,
                      "open" if self._cap is not None else "opening")
            return
        self._cam_opening = True
        self._cam_open_reason = reason
        self._status_var.set(
            _("Opening camera..."))

        def _worker():
            asked = self._camera_index
            try:
                cap, used, frame = self._cam_mod.open_camera_or_fallback(
                    asked, (self._settings or {}).get("camera_index_override"))
                self._camera_index = used
                notice = self._cam_mod.fallback_notice(asked, used) if used != asked else None
                self.after(0, self._camera_ready, cap, frame, notice)
            except Exception as e:
                self.after(0, self._camera_failed, e)

        start_worker(_worker, name='camcal-open')

    def _camera_ready(self, cap, first_frame, notice=None):
        self._cam_opening = False
        if not self.winfo_exists():
            # Closed while the camera was opening. after() callbacks
            # still run on a destroyed widget, and nothing else would
            # ever release this handle.
            try:
                cap.release()
            except Exception:
                pass
            mlog.info("camera #%s opened after the dialog closed; released",
                      self._camera_index)
            return
        if self._cap is not None and self._cap is not cap:
            try:
                self._cap.release()
            except Exception:
                pass
        self._cap = cap
        h, w = first_frame.shape[:2]
        self._image_size = (w, h)
        self._last_frame_shape = (h, w)
        self._live_since = time.monotonic()
        self._live_stats = None
        self._snapshot_this_open = False
        mlog.info("camera #%s live (%s): first frame %s",
                  self._camera_index, self._cam_open_reason,
                  self._cam_mod.describe_frame(first_frame))
        if self._phase == 'capture':
            msg = _("Camera live. Captured 0 / {n} good frames.").format(
                n=self.TARGET_FRAMES)
            self._last_progress_status = msg  # ours: no hold needed
            self._status_var.set(msg)
        if notice:
            # Running on another camera than the saved one
            # (camera_capture.fallback_notice): said in the status line,
            # which the preview leaves alone for STATUS_HOLD_S.
            self._status_var.set(notice)
        # One refresh chain only: cancel any pending tick before starting.
        self._stop_refresh()
        self._refresh_loop()

    def _camera_failed(self, err):
        self._cam_opening = False
        mlog.warning("camera #%s failed to open (%s): %s",
                     self._camera_index, self._cam_open_reason, err)
        self._status_var.set(
            _("Could not open camera index {i}: {e}").format(
                i=self._camera_index, e=err))

    def _refresh_loop(self):
        if not self.winfo_exists() or self._cap is None:
            return
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok, frame = False, None
        if ok and frame is not None:
            if self._failed_reads >= 30:
                mlog.info("camera #%s: reads recovered after %d failures",
                          self._camera_index, self._failed_reads)
            self._failed_reads = 0
            self._latest_frame = frame
            try:
                self._render_frame_with_overlay(frame)
            except Exception:
                # Never kill the loop on a bad frame, and never hide it:
                # v2.80 swallowed an IndexError here on every frame the
                # board was found in (OpenCV 5 corner shape), so the
                # preview froze on the last frame it WASN'T found in and
                # nothing reached the log. The first failure logs its
                # traceback; the rest are counted in the 5 s summary.
                self._loop_errors += 1
                if self._loop_errors == 1:
                    mlog.exception(
                        "calibration preview: processing a frame failed "
                        "(further failures are counted in the summary)")
        else:
            # Camera disconnect / USB hung. Surface it after a few
            # consecutive failures so the user knows to click
            # Reconnect rather than staring at a frozen preview.
            self._failed_reads = getattr(self, '_failed_reads', 0) + 1
            if self._failed_reads == 30:  # ~3s at 10fps
                mlog.warning("camera #%s: 30 consecutive failed reads",
                             self._camera_index)
                try:
                    self._status_var.set(_(
                        "Camera lost — click Reconnect Camera to recover."))
                except tk.TclError:
                    pass
        self._refresh_after_id = self.after(self.LIVE_REFRESH_MS,
                                              self._refresh_loop)

    def _store_engrave_offset(self, ox, oy, how):
        """Remember where the card was engraved, and persist it so a
        closed dialog or a restarted app resumes at the capture step."""
        mlog.info("engrave offset %s: (%.3f, %.3f)", how, ox, oy)
        self._captured_engrave_offset = (ox, oy)
        try:
            self._settings['camera_calibration_engrave_offset_mm'] = [ox, oy]
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            mlog.exception("could not save the engrave offset")

    def _ask_manual_offset(self):
        """Prompt for the engraved card's bottom-left machine X,Y.
        Returns (x, y) or None if cancelled or unreadable (after saying
        so)."""
        from tkinter import simpledialog
        prompt = _(
            "Enter the engraved card's bottom-left machine-mm coords.\n\n"
            "Format: X,Y (e.g. 66.025,99.025)\n\n"
            "This is the offset that was persisted to settings after a "
            "successful engrave. Use this if the persisted value got "
            "lost but the card is still on the bed at known coords.")
        val = simpledialog.askstring(
            _("Manual Engrave Offset"), prompt, parent=self)
        if not val:
            return None
        try:
            parts = val.replace(';', ',').split(',')
            if len(parts) != 2:
                raise ValueError("need two numbers separated by a comma")
            return float(parts[0].strip()), float(parts[1].strip())
        except (ValueError, TypeError) as e:
            messagebox.showerror(
                _("Bad Offset"),
                _("Couldn't parse '{v}' as X,Y mm:\n{e}").format(v=val, e=e),
                parent=self)
            return None

    def _on_manual_offset_clicked(self):
        """Engrave-phase escape hatch. If the user knows the machine-mm
        coords where their engraved card sits (e.g. recovered from a
        backup, or printed at engrave time), they can punch in the
        offset and skip straight to capture phase. Avoids losing 2
        hours of basswood to a state-persistence hiccup."""
        offset = self._ask_manual_offset()
        if offset is None:
            return
        self._store_engrave_offset(*offset, how="entered by hand")
        ox, oy = offset
        # Re-enter engrave phase so Next becomes enabled (the offset
        # check happens at phase-enter time), then user can click Next
        # to advance to capture.
        self._enter_engrave_phase()
        self._status_var.set(
            _("Offset set to ({x:.3f}, {y:.3f}). Click Next to capture.").format(
                x=ox, y=oy))

    def _on_reconnect_camera_clicked(self):
        mlog.info("camera #%s: Reconnect clicked", self._camera_index)
        self._reconnect_camera()

    def _reconnect_camera(self):
        """Release the (likely-dead) cv2.VideoCapture handle and open
        a fresh one. Used when the USB camera died mid-session and the
        existing handle is stale. Doesn't touch captures already taken
        so the user keeps their progress."""
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self._failed_reads = 0
        self._status_var.set(_("Reopening camera..."))
        self.after(50, self._open_camera_and_start, "Reconnect clicked")

    def _on_switch_camera_clicked(self):
        """Cycle to the next enumerated camera; persist the choice.

        Captures from two cameras can't go into one calibration (and the
        image size is per camera), so switching with captures taken
        discards them, after asking."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        if self._captures:
            if not messagebox.askyesno(
                    _("Switch Camera?"),
                    _("Switching cameras discards the {n} captures taken "
                      "so far. Continue?").format(n=len(self._captures)),
                    parent=self):
                return
            mlog.info("switch camera: discarding %d captures",
                      len(self._captures))
            self._captures = []
            self._reference_count = 0
            self._capture_substate = 'reference'
            self._capture_btn.config(text=_("Capture reference"))
            for b in (self._done_refs_btn, self._retake_ref_btn,
                      self._reset_btn, self._calibrate_btn):
                b.config(state="disabled")
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        try:
            self._settings['camera_index_override'] = int(self._camera_index)
            from config import save_settings
            save_settings(self._settings)
        except Exception as e:
            mlog.warning("calibration: could not save the chosen camera to "
                         "settings: %s", e)
        self._status_var.set(
            _("Switched to camera index {i}. Reopening...").format(
                i=self._camera_index))
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self._failed_reads = 0
        self.after(50, self._open_camera_and_start, "Switch camera clicked")

    def _auto_recycle_camera(self):
        """Called after every FalconRunDialog close. The cv2 frame
        stream consistently stalls when a long-lived modal dialog
        closes (preview shows the last good frame forever, with no
        error from cap.read()). Reopen the camera to get fresh frames
        flowing again. No-op if the camera wasn't open."""
        if self._cap is None:
            return
        self._stop_refresh()
        try:
            self._cap.release()
        except Exception:
            pass
        self._cap = None
        self._failed_reads = 0
        self.after(50, self._open_camera_and_start,
                   "recycle after laser dialog")

    def _render_frame_with_overlay(self, frame):
        import cv2
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        det = {}
        t0 = time.perf_counter()  # monotonic() ticks every ~15 ms on Windows
        corners, ids = self._cam_mod.detect_charuco(
            gray, self._board, stats=det)
        self._note_frame(frame, gray, det,
                         (time.perf_counter() - t0) * 1000.0,
                         corners is not None)

        annotated = frame.copy()
        detected_count = 0 if corners is None else len(corners)
        if corners is not None and ids is not None and len(corners) >= 4:
            # Green dots at the detected ChArUco corners. corner_points
            # takes OpenCV 4's (N, 1, 2) and OpenCV 5's (N, 2) alike;
            # indexing pt[0][0] here is what broke v2.80 under OpenCV 5.
            for x, y in self._cam_mod.corner_points(corners):
                cv2.circle(annotated, (int(x), int(y)), 4, (0, 255, 0), -1)
            self._capture_btn.config(state="normal")
        else:
            self._capture_btn.config(state="disabled")

        # OpenCV is BGR; PIL/tk want RGB.
        rgb = cv2.cvtColor(annotated, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        # Cap preview size while preserving aspect ratio
        img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
        self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
        self._preview_label.config(image=self._latest_photo,
                                    width=img.width, height=img.height)

        n_caps = len(self._captures)
        if not self._may_write_progress():
            return
        if corners is not None:
            msg = _("Detected {d} corners — captured {n} / {target} good frames.").format(
                d=detected_count, n=n_caps, target=self.TARGET_FRAMES)
        else:
            msg = _("Move the card into view — captured {n} / {target} good frames.").format(
                n=n_caps, target=self.TARGET_FRAMES)
        self._last_progress_status = msg
        self._status_var.set(msg)

    def _may_write_progress(self):
        """Whether the preview may write its progress line now.

        It runs ~10x a second, so before v2.86 every other message
        ("Homed", "Framing done", "Engrave complete", the capture hints)
        was overwritten before anyone could read it. Now it writes only
        in capture phase, and only once anything else it finds in the
        status line has been there STATUS_HOLD_S seconds."""
        if self._phase != 'capture':
            return False
        cur = self._status_var.get()
        if cur == self._last_progress_status:
            return True
        now = time.monotonic()
        if cur != self._foreign_status:
            self._foreign_status = cur
            self._foreign_status_t = now
        return now - self._foreign_status_t >= self.STATUS_HOLD_S

    def _note_frame(self, frame, gray, det, detect_ms, good):
        """Accumulate one preview frame into the detection summary and
        write the summary to the log once per SUMMARY_INTERVAL_S.

        The summary is what answers "why are there no green corners on
        this user's machine": no markers at all (focus, exposure, wrong
        camera) reads differently from markers without corners (board
        mismatch), slow detection, or frames that error out."""
        now = time.monotonic()
        st = self._live_stats
        if st is None:
            st = self._live_stats = {
                't0': now, 'frames': 0, 'with_markers': 0, 'good': 0,
                'max_markers': 0, 'max_corners': 0,
                'ms_min': None, 'ms_max': 0.0,
                'g_min': 255, 'g_max': 0, 'g_mean_sum': 0.0,
                'errors0': self._loop_errors,
            }
        markers = det.get('markers', 0)
        st['frames'] += 1
        st['with_markers'] += 1 if markers else 0
        st['max_markers'] = max(st['max_markers'], markers)
        st['max_corners'] = max(st['max_corners'], det.get('corners', 0))
        if good:
            st['good'] += 1
            self._last_good_t = now
        st['ms_min'] = (detect_ms if st['ms_min'] is None
                        else min(st['ms_min'], detect_ms))
        st['ms_max'] = max(st['ms_max'], detect_ms)
        st['g_min'] = min(st['g_min'], int(gray.min()))
        st['g_max'] = max(st['g_max'], int(gray.max()))
        st['g_mean_sum'] += float(gray.mean())
        if now - st['t0'] >= self.SUMMARY_INTERVAL_S:
            h, w = frame.shape[:2]
            mlog.info(
                "calibration preview, last %.1fs: %d frames %dx%d; markers "
                "in %d (max %d); board found in %d (max %d corners); gray "
                "min/mean/max %d/%.0f/%d; detect %.0f-%.0f ms; frame errors "
                "%d; captures %d (%d references)",
                now - st['t0'], st['frames'], w, h, st['with_markers'],
                st['max_markers'], st['good'], st['max_corners'],
                st['g_min'], st['g_mean_sum'] / st['frames'], st['g_max'],
                st['ms_min'], st['ms_max'],
                self._loop_errors - st['errors0'],
                len(self._captures), self._reference_count)
            self._live_stats = None
        self._maybe_snapshot(frame, now)

    def _maybe_snapshot(self, frame, now):
        """Save the raw frame a few seconds after the camera goes live,
        and again (at most once a minute) while the board hasn't been
        found for SNAPSHOT_NO_DETECT_S. Overwrites one file, like
        last_frame_cut.gcode: it is the exact image the detector saw."""
        if self._live_since is None:
            return
        if not self._snapshot_this_open:
            due = now - self._live_since >= self.SNAPSHOT_FIRST_S
        else:
            since_good = now - max(self._live_since,
                                   self._last_good_t or self._live_since)
            due = (since_good >= self.SNAPSHOT_NO_DETECT_S
                   and now - (self._last_snapshot_t or 0.0)
                   >= self.SNAPSHOT_MIN_GAP_S)
        if due:
            self._save_snapshot(frame, now)

    def _save_snapshot(self, frame, now):
        self._snapshot_this_open = True
        self._last_snapshot_t = now
        try:
            import cv2
            from config import ensure_config_dir
            path = os.path.join(ensure_config_dir(), self.SNAPSHOT_NAME)
            # imencode + Python file write, not cv2.imwrite: imwrite can't
            # open a non-ASCII path on Windows (a user name with an accent
            # is enough) and only says so by returning False.
            ok, buf = cv2.imencode(".png", frame)
            if not ok:
                raise RuntimeError("PNG encoding failed")
            with open(path, "wb") as f:
                f.write(buf.tobytes())
            mlog.info("saved the raw camera frame to %s (%s)", path,
                      self._cam_mod.describe_frame(frame))
        except Exception:
            mlog.exception("could not save the raw camera frame")

    def _on_capture_clicked(self):
        if not hasattr(self, '_latest_frame'):
            return
        import cv2
        gray = cv2.cvtColor(self._latest_frame, cv2.COLOR_BGR2GRAY)
        corners, ids = self._cam_mod.detect_charuco(gray, self._board)
        if corners is None or ids is None or len(corners) < 4:
            return
        self._captures.append((corners, ids))
        is_ref = (self._capture_substate == 'reference')
        if is_ref:
            self._reference_count += 1
        mlog.info("capture %d (%s): %d corners", len(self._captures),
                  "reference" if is_ref else "card moved", len(corners))
        n = len(self._captures)
        target = self.TARGET_FRAMES
        if n >= self._cam_mod.CALIB_MIN_FRAMES:
            self._calibrate_btn.config(state="normal")
        if n >= 1:
            self._retake_ref_btn.config(state="normal")
            self._reset_btn.config(state="normal")
            self._done_refs_btn.config(
                state="normal" if is_ref else "disabled")
        if is_ref:
            hint = _("REFERENCE captured ({nref} of card untouched). "
                      "Take more references for higher accuracy, "
                      "Retake last if detection was bad, or click "
                      "'Done with references' when ready to MOVE "
                      "the card.").format(nref=self._reference_count)
        elif n >= target:
            hint = _("Click Calibrate & Save when ready.")
        else:
            hint = _("MOVE the card to a new position and capture again.")
        self._status_var.set(
            _("Captured {n} / {t} good frames. {hint}").format(
                n=n, t=target, hint=hint))

    def _on_done_refs_clicked(self):
        """Switch from reference substate to intrinsics substate.
        Subsequent captures will be 'card moved' poses for the
        distortion solve."""
        if self._reference_count < 1:
            return
        self._capture_substate = 'intrinsics'
        mlog.info("done with references: %d reference captures",
                  self._reference_count)
        self._capture_btn.config(text=_("Capture (card moved)"))
        self._done_refs_btn.config(state="disabled")
        self._status_var.set(
            _("References done ({n} pooled). Now MOVE the card to "
              "different positions / tilts and capture {m} more "
              "frames for lens-distortion calibration.").format(
                n=self._reference_count,
                m=max(0, self.TARGET_FRAMES - len(self._captures))))

    def _on_retake_last_clicked(self):
        """Drop the most recent capture so the user can redo it.

        References are always the first ``_reference_count`` captures, so
        the dropped one was a reference exactly when its index is below
        that count. If that leaves no reference, go back to taking
        references: v2.80 stayed in "card moved" mode with zero
        references, the next moved capture became capture 0, and
        Calibrate used it as the reference -- a calibration tens of mm
        off that still reported a fine RMS."""
        if not self._captures:
            return
        dropped = len(self._captures) - 1
        self._captures.pop()
        if dropped < self._reference_count:
            self._reference_count -= 1
        if self._reference_count == 0 and self._capture_substate != 'reference':
            self._capture_substate = 'reference'
            self._capture_btn.config(text=_("Capture reference"))
        n = len(self._captures)
        self._done_refs_btn.config(
            state="normal" if (self._capture_substate == 'reference'
                               and self._reference_count >= 1)
            else "disabled")
        mlog.info("retake: dropped the last capture; %d left "
                  "(%d references)", n, self._reference_count)
        self._calibrate_btn.config(
            state="normal" if n >= self._cam_mod.CALIB_MIN_FRAMES
            else "disabled")
        if n == 0:
            self._retake_ref_btn.config(state="disabled")
            self._reset_btn.config(state="disabled")
            self._done_refs_btn.config(state="disabled")
        self._status_var.set(
            _("Removed last capture. {n} / {t} remaining.").format(
                n=n, t=self.TARGET_FRAMES))

    def _on_reset_captures_clicked(self):
        """Discard all captures and restart. Card needs to be back at
        the engraved position for the next first reference capture."""
        if not messagebox.askyesno(
                _("Reset all captures?"),
                _("Discard all {n} captures and start over.\n\n"
                  "If you've moved the card, move it BACK to its "
                  "engraved position before retaking references.\n\n"
                  "Continue?").format(n=len(self._captures)),
                parent=self):
            return
        mlog.info("reset: discarded %d captures", len(self._captures))
        self._captures = []
        self._reference_count = 0
        self._capture_substate = 'reference'
        self._capture_btn.config(text=_("Capture reference"))
        self._retake_ref_btn.config(state="disabled")
        self._reset_btn.config(state="disabled")
        self._done_refs_btn.config(state="disabled")
        self._calibrate_btn.config(state="disabled")
        self._status_var.set(
            _("Captures cleared. Place the card in its engraved "
              "position and click Capture reference."))

    def _calibration_inputs_ok(self):
        """Refuse a capture set that would save a wrong calibration.

        Each of these used to calibrate and save anyway, and the RMS the
        dialog reports doesn't show it (measured on a simulated Falcon
        camera, bed error over the whole bed):
          - no reference capture left: a moved frame became the
            reference, ~50 mm off;
          - Done with references never clicked: every capture counted as
            a reference, 21-43 mm off;
          - no engrave position on file (the dialog opened in capture
            phase without one): the (50, 50) default was used silently.
        """
        refs = self._reference_count
        moved = len(self._captures) - refs
        if refs < 1:
            messagebox.showwarning(
                _("No Reference Capture"),
                _("Calibration needs at least one reference capture: the "
                  "card exactly where it was engraved, taken before "
                  "clicking 'Done with references'.\n\n"
                  "Click Reset all, put the card back where it was "
                  "engraved if you moved it, and start with Capture "
                  "reference."), parent=self)
            return False
        if (self._capture_substate != 'intrinsics'
                or moved < self.MIN_MOVED_CAPTURES):
            messagebox.showwarning(
                _("Move the Card"),
                _("Calibration needs at least {m} captures with the card "
                  "moved to different spots on the bed (you have {n}).\n\n"
                  "Click 'Done with references' if you haven't yet, then "
                  "move the card between captures.").format(
                    m=self.MIN_MOVED_CAPTURES, n=max(0, moved)
                    if self._capture_substate == 'intrinsics' else 0),
                parent=self)
            return False
        if self._captured_engrave_offset is None:
            if not messagebox.askyesno(
                    _("No Engrave Position"),
                    _("The app doesn't know where on the bed this card was "
                      "engraved, so it can't tie the camera to machine "
                      "coordinates.\n\n"
                      "Enter the card's bottom-left machine position now?"),
                    parent=self):
                return False
            offset = self._ask_manual_offset()
            if offset is None:
                return False
            self._store_engrave_offset(*offset, how="entered by hand")
        return True

    def _on_calibrate_clicked(self):
        if len(self._captures) < self._cam_mod.CALIB_MIN_FRAMES:
            return
        if not self._calibration_inputs_ok():
            return
        self._stop_refresh()
        self._status_var.set(_("Computing calibration ..."))
        self.update()
        try:
            eng_off = getattr(self, '_captured_engrave_offset', None)
            kwargs = {}
            if eng_off:
                kwargs['card_offset_x_mm'] = eng_off[0]
                kwargs['card_offset_y_mm'] = eng_off[1]
            # Pool ALL reference captures for the machine-mm homography.
            # The first N captures (where N = _reference_count) were
            # taken with the card untouched at its engraved position.
            ref_count = max(1, getattr(self, '_reference_count', 1))
            kwargs['reference_indices'] = tuple(range(ref_count))
            mlog.info("calibrating: %d captures (%d references), image "
                      "%s, engrave offset %s", len(self._captures),
                      ref_count, self._image_size, eng_off)
            calib = self._cam_mod.calibrate_from_frames(
                self._captures, self._image_size, self._board, **kwargs)
            self._cam_mod.save_calibration(calib, self._calibration_path)
            cm = calib.get('camera_matrix') or [[0] * 3] * 3
            mlog.info("calibration saved to %s: rms %.3f px, fx %.1f fy "
                      "%.1f cx %.1f cy %.1f, dist %s",
                      self._calibration_path,
                      calib.get('rms_reprojection_error_px', -1),
                      cm[0][0], cm[1][1], cm[0][2], cm[1][2],
                      calib.get('dist_coeffs'))
            # Calibration committed. Clear the in-progress recovery
            # hint (engrave offset in settings) — it was useful for
            # restoring state during the capture phase, but now the
            # full calibration is on disk. Next dialog open should
            # start in engrave phase for a FRESH recalibration, not
            # auto-skip to capture against a stale offset.
            #
            # Also persist the working camera index so future opens
            # (live preview, scrap capture) skip the enumerate /
            # find-Falcon guess and go straight to the right camera.
            try:
                self._settings.pop(
                    'camera_calibration_engrave_offset_mm', None)
                self._settings['camera_index_override'] = int(
                    self._camera_index)
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                mlog.exception(
                    "calibration saved, but the camera number and the "
                    "cleared engrave offset could not be written to settings")
            self.result = "saved"
            messagebox.showinfo(
                _("Calibration Saved"),
                _("Camera calibration complete.\n\n"
                  "RMS reprojection error: {rms:.2f} px (lower is better; "
                  "< 1 px is excellent)\n"
                  "Frames used: {n}\n\n"
                  "Saved to:\n{p}").format(
                    rms=calib['rms_reprojection_error_px'],
                    n=calib['frame_count'], p=self._calibration_path),
                parent=self)
            self._on_close()
        except Exception as e:
            mlog.exception("calibration failed")
            messagebox.showerror(_("Calibration Failed"),
                                  _("Calibration did not converge:\n\n{e}\n\n"
                                    "Try capturing more frames at different "
                                    "positions and tilts.").format(e=e),
                                  parent=self)
            # Resume the refresh loop so user can capture more
            self._refresh_loop()

    def _stop_refresh(self):
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except (tk.TclError, ValueError):
                pass
            self._refresh_after_id = None

    def _on_close(self):
        mlog.info("calibration dialog closed: result=%s, %d captures "
                  "(%d references)", self.result, len(self._captures),
                  self._reference_count)
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        try:
            import sleep_lock
            sleep_lock.allow_sleep()
        except Exception as e:
            mlog.warning("calibration: could not release the keep-awake "
                         "lock: %s", e)
        self.destroy()


class CameraCaptureDialog(tk.Toplevel):
    """Live camera-capture dialog for snapping a scrap polygon from the bed.

    Shows the live camera feed with the detected scrap contour overlaid
    in green. User can adjust the threshold sensitivity, then click Use
    to accept the polygon (returned in mm coords) or Retry to keep
    looking. Requires a saved camera calibration.

    Result:
        self.result_polygon_mm = list of (x_mm, y_mm) on success, None on cancel.
    """

    LIVE_REFRESH_MS = 100
    PREVIEW_W = 640
    PREVIEW_H = 480
    NOTICE_HOLD_S = 6.0     # the fallback notice stays this long before the detection line resumes

    def __init__(self, parent, camera_index, calibration_path, settings=None):
        super().__init__(parent)
        self.title(_("Capture Scrap Outline"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.result_polygon_mm = None

        self._camera_index = camera_index
        self._cap = None
        self._loop_errors = 0            # frames that failed to process (first logged, rest counted)
        self._calibration = None
        self._latest_undistorted = None
        self._latest_polygon_px = None
        self._latest_polygon_mm = None
        self._notice_until = 0.0
        self._bias_save_errors = 0  # failed slider saves (first one logged)
        self._latest_photo = None
        self._refresh_after_id = None
        # Optional settings dict — when present, the detection-bias
        # slider persists its value across captures via
        # 'camera_detection_threshold_bias'. When None (e.g. when the
        # dialog was spawned from inside PolygonDrawWindow without a
        # settings handle), the slider still works but its value is
        # lost on close.
        self._settings = settings
        try:
            self._initial_bias = int(
                settings.get('camera_detection_threshold_bias', 0)
                if settings else 0)
        except (TypeError, ValueError):
            self._initial_bias = 0

        # Dependency check + calibration load
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
            self._calibration = camera_capture.load_calibration(calibration_path)
            if not self._calibration:
                raise RuntimeError(_("No camera calibration found. Run "
                                       "Options > Camera Calibration first."))
        except (ImportError, RuntimeError) as e:
            tk.Label(self, text=str(e), bg=DIALOG_BG, fg="#a30000",
                     padx=20, pady=20, wraplength=400).pack()
            tk.Button(self, text=_("Close"),
                      command=self.destroy).pack(pady=(0, 15))
            return

        # ---- UI ----
        tk.Label(self, text=_("Capture Scrap Outline"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold")
                 ).pack(pady=(10, 2))
        tk.Label(self, text=_(
            "Place the scrap on the laser bed (close the cover). The green "
            "outline shows what will be captured. Click Use when it looks "
            "right, or Retry to re-detect."),
            bg=DIALOG_BG, wraplength=620, justify="left",
            font=("Helvetica", 9)).pack(padx=15, pady=(0, 6))

        # Preview: a blank image the size of the preview until the first
        # frame arrives (the status line below says the camera is
        # opening). A label holding TEXT reads width/height in
        # characters: the old 640x480 "Waking camera..." label asked for
        # ~4500x7200 px and, if the camera never opened, pushed the
        # status line and the Use/Cancel buttons off the window.
        self._preview_blank = tk.PhotoImage(
            width=self.PREVIEW_W, height=self.PREVIEW_H)
        self._preview_label = tk.Label(
            self, bg="#222222", image=self._preview_blank)
        self._preview_label.pack(padx=15, pady=(0, 8))

        self._status_var = tk.StringVar(value=_("Opening camera ..."))
        tk.Label(self, textvariable=self._status_var, bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(pady=(0, 4))

        # Detection sensitivity slider — biases the auto-threshold
        # used to separate scrap from bed. Center = use OpenCV's Otsu
        # auto-pick (works for most lighting). Slide LEFT (more
        # sensitive) when the scrap barely contrasts with the bed,
        # RIGHT (less sensitive) when honeycomb shadows are getting
        # picked up as material. Updates the live overlay on every
        # next frame.
        bias_frame = tk.Frame(self, bg=DIALOG_BG)
        bias_frame.pack(fill='x', padx=20, pady=(0, 6))
        tk.Label(bias_frame, text=_("Detection sensitivity:"),
                 bg=DIALOG_BG, font=("Helvetica", 9)
                 ).pack(side='left', padx=(0, 6))
        tk.Label(bias_frame, text=_("more"), bg=DIALOG_BG,
                 fg='#666', font=("Helvetica", 8)).pack(side='left')
        self._bias_var = tk.IntVar(value=self._initial_bias)
        # Slider value maps directly to threshold_bias:
        #   -80 (left)  = subtract 80 from Otsu → lower threshold →
        #                 more pixels classified as scrap = MORE sensitive
        #   +80 (right) = add 80 → fewer pixels = LESS sensitive
        tk.Scale(bias_frame, from_=-80, to=80, orient='horizontal',
                 variable=self._bias_var, bg=DIALOG_BG,
                 length=300, showvalue=False,
                 command=self._on_bias_changed).pack(
                     side='left', padx=4, fill='x', expand=True)
        tk.Label(bias_frame, text=_("less"), bg=DIALOG_BG,
                 fg='#666', font=("Helvetica", 8)).pack(side='left')

        # Invert toggle — flips the threshold polarity. Default off
        # (THRESH_BINARY) assumes scrap is BRIGHTER than the bed (the
        # common case: leather on a dark honeycomb). Turn on for the
        # opposite (e.g. dark felt on a light scrap board) — without
        # inversion, dark-on-light produces no contour because the
        # largest bright region is the bed itself. Persisted so users
        # who routinely shoot one polarity don't re-toggle each time.
        invert_frame = tk.Frame(self, bg=DIALOG_BG)
        invert_frame.pack(fill='x', padx=20, pady=(0, 6))
        try:
            initial_invert = bool(
                self._settings.get('camera_detection_invert', False)
                if self._settings else False)
        except (AttributeError, TypeError):
            initial_invert = False
        self._invert_var = tk.BooleanVar(value=initial_invert)
        tk.Checkbutton(
            invert_frame, bg=DIALOG_BG,
            text=_("Invert colors (use when scrap is DARKER than the bed)"),
            variable=self._invert_var,
            command=self._on_invert_changed).pack(side='left')

        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 12))
        self._use_btn = tk.Button(
            btn_frame, text=_("Use"), command=self._on_use_clicked,
            font=("Helvetica", 10, "bold"), width=12, state="disabled")
        self._use_btn.pack(side="left", padx=5)
        # Switch camera — fallback for when the auto-resolver picked the
        # wrong device (e.g. integrated laptop webcam). Cycles through
        # enumerated cameras and persists the choice.
        tk.Button(btn_frame, text=_("Switch camera"),
                  command=self._on_switch_camera, width=14
                  ).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Cancel"),
                  command=self._on_close, width=10
                  ).pack(side="left", padx=5)

        self.geometry(f"{self.PREVIEW_W + 50}x{self.PREVIEW_H + 240}")
        self.minsize(660, 660)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.after(50, self._open_camera_and_start)
        self.wait_window(self)

    def _on_switch_camera(self):
        """Cycle to the next enumerated camera; persist the choice."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        if self._settings is not None:
            self._settings['camera_index_override'] = int(self._camera_index)
            try:
                from config import save_settings
                save_settings(self._settings)
            except Exception as e:
                mlog.warning("camera capture: could not save the chosen "
                             "camera to settings: %s", e)
        self._status_var.set(
            _("Switched to camera index {i}. Reopening...").format(
                i=self._camera_index))
        # Release current and reopen.
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except (tk.TclError, ValueError):
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self.after(50, self._open_camera_and_start)

    def _open_camera_and_start(self):
        """Background-thread camera open so the dialog stays responsive
        during cv2.VideoCapture's 1-3 second initialization."""
        self._status_var.set(
            _("Opening camera..."))

        def _worker():
            asked = self._camera_index
            try:
                # calibration_present=True: this dialog maps the picture to
                # the bed through the calibration it loaded, so only the
                # Falcon-named camera may stand in for the saved one.
                cap, used, _frame = self._cam_mod.open_camera_or_fallback(
                    asked, (self._settings or {}).get("camera_index_override"),
                    calibration_present=True)
                self._camera_index = used
                notice = self._cam_mod.fallback_notice(asked, used) if used != asked else None
                self.after(0, self._camera_ready, cap, notice)
            except Exception as e:
                self.after(0, self._camera_failed, e)

        start_worker(_worker, name='camcap-open')

    def _camera_ready(self, cap, notice=None):
        if not self.winfo_exists():
            # Closed while the camera was opening; release it here or
            # nothing will.
            try:
                cap.release()
            except Exception:
                pass
            return
        self._cap = cap
        if notice:
            # Running on another camera than the saved one
            # (camera_capture.fallback_notice). The detection line is
            # rewritten on every frame, so it waits NOTICE_HOLD_S.
            self._notice_until = time.monotonic() + self.NOTICE_HOLD_S
            self._status_var.set(notice)
        self._refresh_loop()

    def _camera_failed(self, err):
        self._status_var.set(_("Could not open camera: {e}").format(e=err))

    def _set_live_status(self, text):
        """The per-frame status line, unless the fallback notice is still
        being shown."""
        if time.monotonic() < self._notice_until:
            return
        self._status_var.set(text)

    def _refresh_loop(self):
        if not self.winfo_exists() or self._cap is None:
            return
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok, frame = False, None
        if ok and frame is not None:
            try:
                self._render_frame_with_overlay(frame)
            except Exception:
                # The calibration preview's rule (2026-11-10): never kill
                # the loop on a bad frame, never hide it. Until then one
                # exception here ended the loop and froze the preview for
                # good. The first failure logs its traceback; the rest are
                # counted and said when the dialog closes.
                self._loop_errors += 1
                if self._loop_errors == 1:
                    mlog.exception("camera capture: processing a frame failed "
                                   "(further failures are counted)")
        self._refresh_after_id = self.after(self.LIVE_REFRESH_MS,
                                              self._refresh_loop)

    def _on_bias_changed(self, _value):
        """Persist the slider value to settings (if available). The
        next live-refresh tick will use the new bias automatically —
        no need to force a re-render here."""
        if self._settings is not None:
            try:
                self._settings['camera_detection_threshold_bias'] = int(
                    self._bias_var.get())
                from config import save_settings
                save_settings(self._settings)
            except Exception as e:
                # Never block the live preview on save failure. The slider
                # fires on every step of a drag: log the first, count the rest.
                self._bias_save_errors += 1
                if self._bias_save_errors == 1:
                    mlog.warning("camera capture: could not save the detection "
                                 "sensitivity to settings: %s (further "
                                 "failures are counted)", e)

    def _on_invert_changed(self):
        """Persist the invert flag — same pattern as the bias slider."""
        if self._settings is not None:
            try:
                self._settings['camera_detection_invert'] = bool(
                    self._invert_var.get())
                from config import save_settings
                save_settings(self._settings)
            except Exception as e:
                mlog.warning("camera capture: could not save the invert "
                             "setting: %s", e)

    def _render_frame_with_overlay(self, frame):
        import cv2
        mismatch = self._cam_mod.calibration_size_mismatch(
            self._calibration, frame)
        if mismatch:
            cw, ch, fw, fh = mismatch
            if not getattr(self, '_size_mismatch_logged', False):
                self._size_mismatch_logged = True
                mlog.warning("scrap capture: camera #%s sends %dx%d, "
                             "calibrated at %dx%d", self._camera_index,
                             fw, fh, cw, ch)
            self._latest_polygon_mm = None
            self._use_btn.config(state="disabled")
            self._set_live_status(_(
                "This camera is sending {fw}×{fh} images, but the "
                "calibration was made at {cw}×{ch}. Recalibrate, or use "
                "Switch camera if this is the wrong one.").format(
                    fw=fw, fh=fh, cw=cw, ch=ch))
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = self._PIL_Image.fromarray(rgb)
            img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
            self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
            self._preview_label.config(image=self._latest_photo,
                                        width=img.width, height=img.height)
            return
        # Undistort + detect contour
        undistorted = self._cam_mod.undistort_frame(frame, self._calibration)
        try:
            bias = int(self._bias_var.get())
        except (AttributeError, tk.TclError, ValueError):
            bias = 0
        try:
            invert = bool(self._invert_var.get())
        except (AttributeError, tk.TclError, ValueError):
            invert = False
        polygon_px = self._cam_mod.detect_scrap_contour(
            undistorted, threshold_bias=bias, invert=invert)
        self._latest_undistorted = undistorted
        self._latest_polygon_px = polygon_px
        if polygon_px:
            self._latest_polygon_mm = self._cam_mod.pixels_to_mm(
                polygon_px, self._calibration)
        else:
            self._latest_polygon_mm = None

        # Overlay the polygon
        annotated = undistorted.copy()
        if polygon_px and len(polygon_px) >= 3:
            import numpy as np
            pts = np.array(polygon_px, dtype=np.int32).reshape(-1, 1, 2)
            cv2.polylines(annotated, [pts], isClosed=True,
                          color=(0, 255, 0), thickness=2)
            self._use_btn.config(state="normal")
            n_pts = len(polygon_px)
            # Compute bbox extent in mm for the status line
            xs_mm = [p[0] for p in self._latest_polygon_mm]
            ys_mm = [p[1] for p in self._latest_polygon_mm]
            w_mm = max(xs_mm) - min(xs_mm)
            h_mm = max(ys_mm) - min(ys_mm)
            self._set_live_status(_(
                "Detected: {n}-vertex polygon, {w:.0f} × {h:.0f} mm bbox").format(
                n=n_pts, w=w_mm, h=h_mm))
        else:
            self._use_btn.config(state="disabled")
            self._set_live_status(_("No scrap detected — place a piece on the bed."))

        # OpenCV BGR → PIL RGB
        rgb = cv2.cvtColor(annotated, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
        self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
        self._preview_label.config(image=self._latest_photo,
                                    width=img.width, height=img.height)

    def _on_use_clicked(self):
        if self._latest_polygon_mm:
            # Return the polygon in ABSOLUTE machine-Y-up coords —
            # do NOT normalize here. Caller (main.py) handles both the
            # normalization and the Y-up→storage-convention conversion
            # so it can also derive the polygon's machine-mm offset
            # (needed for Frame & Cut auto mode). Pre-normalizing here
            # would zero the offset and lose the scrap's bed position.
            self.result_polygon_mm = list(self._latest_polygon_mm)
        self._on_close()

    def _on_close(self):
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except (tk.TclError, ValueError):
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        if self._loop_errors:
            mlog.warning("camera capture: %d frames failed to process", self._loop_errors)
            self._loop_errors = 0
        self.destroy()


# Grbl 1.1 error / alarm code translations — picked from the official
# Grbl error/alarm tables. Not exhaustive; covers the codes a user is
# most likely to hit in a Frame & Cut workflow. Returns None if we
# don't have a translation (caller shows just the raw text).
_GRBL_ERROR_TEXT = {
    '1': "G-code letter missing or malformed",
    '2': "Numeric value invalid",
    '3': "'$' command not recognized",
    '5': "Homing cycle is not enabled in Grbl ($22=0)",
    '7': "EEPROM read failed — settings restored",
    '8': "'$' command only valid when idle",
    '9': "G-code locked out during alarm or jog",
    '10': "Soft limits cannot be enabled without homing",
    '15': "Travel exceeds machine soft limits",
    '17': "Setting disabled",
    '20': "Unsupported G-code command",
    '22': "Feed rate has not been set",
    '24': "Two G-code commands tried to use the same axis",
    '33': "Motion target invalid",
}
_GRBL_ALARM_TEXT = {
    '1': "Hard limit triggered — machine position lost",
    '2': "G-code motion would exceed soft limits",
    '3': "Reset while in motion — position lost",
    '8': "Homing fail — pull-off didn't clear limit switch",
    '9': "Homing fail — limit switch not contacted in cycle",
}


def _decode_grbl_error(msg):
    """Translate 'error:N' or 'ALARM:N' to a plain-English description."""
    if not msg:
        return None
    parts = msg.strip().split(':', 1)
    if len(parts) != 2:
        return None
    tag, code = parts[0].lower(), parts[1].strip()
    if tag == 'error':
        return _GRBL_ERROR_TEXT.get(code)
    if tag == 'alarm':
        return _GRBL_ALARM_TEXT.get(code)
    return None


class LidClosedDialog(tk.Toplevel):
    """The one question asked before the laser comes on: is the lid closed?

    Two buttons with real labels, "Yes, Start Frame" and "No, Cancel",
    because tk's askyesno can only say Yes / No and Matt wanted the
    buttons to say what they do (2026-10-26). `result` is True only for
    the Yes button; No, the close box and Escape all answer False.
    """

    def __init__(self, parent, title, message, yes_label, no_label):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.resizable(False, False)
        self.result = False

        tk.Label(self, text=message, wraplength=380, bg=DIALOG_BG,
                 justify="left", font=("Helvetica", 11)).pack(padx=24, pady=(18, 10))
        buttons = tk.Frame(self, bg=DIALOG_BG)
        buttons.pack(pady=(4, 16))
        yes_btn = tk.Button(buttons, text=yes_label, command=self._on_yes)
        no_btn = tk.Button(buttons, text=no_label, command=self._on_no)
        # Both buttons one width: 16 characters, as always, or as many as
        # the longer label needs in this language. At a fixed 16 the French
        # "Oui, démarrer l'encadrement" was cut off, and for a while the
        # French and Italian labels were shortened to fit the button
        # instead of the other way round.
        needed = max(yes_btn.winfo_reqwidth(), no_btn.winfo_reqwidth())
        width = 16
        yes_btn.config(width=width)
        while yes_btn.winfo_reqwidth() < needed and width < 60:
            width += 1
            yes_btn.config(width=width)
        no_btn.config(width=width)
        yes_btn.pack(side="left", padx=8)
        no_btn.pack(side="left", padx=8)

        self.bind("<Return>", lambda e: self._on_yes())
        self.bind("<Escape>", lambda e: self._on_no())
        self.protocol("WM_DELETE_WINDOW", self._on_no)

        # Size to content and centre over the parent dialog.
        self.update_idletasks()
        self.geometry("")
        try:
            w, h = self.winfo_reqwidth(), self.winfo_reqheight()
            x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass
        yes_btn.focus_set()
        self.grab_set()
        self.wait_window(self)

    def _on_yes(self):
        self.result = True
        self.destroy()

    def _on_no(self):
        self.result = False
        self.destroy()


class FalconRunDialog(tk.Toplevel):
    """Live-progress dialog while G-code streams to the Falcon.

    Owns a FalconSender for the duration of the job. Callbacks from the
    sender's worker thread are marshalled to the Tk main thread via
    self.after(0, ...). Buttons send real-time commands via the sender.

    Window-close ⨯ behaves as Stop (with confirm if mid-job).
    """

    def __init__(self, parent, sender, gcode_lines, title=None,
                  on_finished=None, loop=False, gcode_provider=None,
                  show_pause_resume=True, show_cut_button=True,
                  stop_needs_confirm=True, done_button_label=None,
                  require_jog_first=False, remind_lid_on_advance=False,
                  auto_locate_target=None, is_homed=False):
        super().__init__(parent)
        self.title(title or _("Sending to Falcon"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        # Resizable so a long error message (Grbl code + decoded
        # description + failing line) doesn't get clipped.
        self.resizable(True, True)

        self._sender = sender
        self._gcode_lines = gcode_lines  # for error context
        self._total = len(gcode_lines)
        self._sent = 0
        self._start_time = time.monotonic()
        self._finished = False
        self._on_finished = on_finished
        self._final_reason = None
        self._last_error = None  # preserved across _on_done overwrite
        # Loop mode: restart the stream after each successful completion
        # until the user clicks "Looks Good — Cut!" or "Stop". Used
        # for the framing pass so the user has time to verify
        # alignment (or jog the head) across multiple traces.
        #
        # When ``gcode_provider`` is set, each loop iteration calls
        # it to regenerate fresh G-code instead of re-streaming the
        # same lines. Used by the calibration dialog's "Frame &
        # Engrave Here" flow so the framing trace TRACKS the head as
        # the user jogs between passes (provider re-reads MPos and
        # generates framing centered on the new position).
        self._loop = loop
        self._gcode_provider = gcode_provider
        self._cut_requested = False
        self._pass_count = 0
        # Hide pause/resume in contexts where pause-jog-resume can
        # corrupt state (e.g. calibration framing — absolute coords
        # mean the next G-code line after resume hops back to the
        # pre-jog absolute position regardless of how the user jogged).
        self._show_pause_resume = show_pause_resume
        # The green "Looks Good — Cut!" button only makes sense when
        # the run dialog is followed by a CUT in the same workflow.
        # Calibration framing has Frame and Engrave as separate
        # buttons in the parent dialog, so the in-run Cut button
        # would be misleading there.
        self._show_cut_button = show_cut_button
        # Whether Stop should pop a "are you sure? this can ruin
        # material" confirmation. False for low-power framing where
        # stopping is harmless; True for actual cuts where mid-job
        # abort can burn weird marks into the material.
        self._stop_needs_confirm = stop_needs_confirm
        # Label for the post-completion button (default "Close").
        # Callers that hand off to a follow-up step (e.g. the seed
        # dialog that drives the head to scrap position then yields
        # to a framing loop) override this to "Start Frame →" so the
        # button text describes what happens next, not just that this
        # dialog goes away.
        self._done_button_label = done_button_label or _("Close")
        # Set by callers whose advance button starts a pass that FIRES
        # THE LASER (Frame & Cut's "Start Frame →"). The dialog can't
        # know what happens after it closes, so the caller declares it.
        self._remind_lid_on_advance = bool(remind_lid_on_advance)
        # MANUAL-mode safety gate: the post-stream done button stays
        # disabled until the user clicks at least one jog button. Used
        # by the "jog to material position" step so the user can't
        # accidentally proceed with the head still at the home corner
        # (where framing/cutting could destroy material). When False
        # (the default), the done button enables immediately on
        # stream completion.
        self._require_jog_first = require_jog_first
        self._has_jogged = not require_jog_first
        # Auto-locate target: if set, show a "Try Auto Locate" button
        # next to Home Laser that drives the head to this machine coord
        # (the polygon's LB vertex). Locked until the laser has been
        # homed in this session — relying on absolute machine coords
        # before homing means MPos drift could send the head to the
        # wrong place. is_homed reflects the parent app's "homed in this
        # session" tracking; the dialog's own Home button updates it
        # when homing succeeds.
        self._auto_locate_target = auto_locate_target
        self._is_homed = is_homed

        # Hook up callbacks. Marshal each to the Tk thread.
        sender.on_status = lambda s: self.after(0, self._on_status, s)
        sender.on_progress = lambda i, n: self.after(0, self._on_progress, i, n)
        sender.on_error = lambda m: self.after(0, self._on_error, m)
        sender.on_alarm = lambda m: self.after(0, self._on_alarm, m)
        sender.on_done = lambda r: self.after(0, self._on_done, r)

        # ---- UI ----
        # Jog-only mode: caller passed an empty gcode list, meaning
        # this dialog is purely a "position the head" step (no actual
        # stream to run). Hide the progress bar / line counter /
        # elapsed timer in that case — they all read 0 forever and
        # confuse users into thinking the dialog is broken.
        self._is_jog_only_mode = (self._total == 0)

        # Wrap long titles so the dialog doesn't stretch the full
        # screen width to fit a one-line instructional sentence.
        tk.Label(self, text=title or _("Sending to Falcon"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold"),
                 wraplength=480, justify="center"
                 ).pack(pady=(12, 4), padx=12)

        # Initial state label. For jog-only mode, "Connecting…" is
        # misleading (it transitions to "Complete ✓" the moment the
        # empty stream returns, which Matt observed as confusing). Use
        # a state that describes what the user should DO instead.
        if self._is_jog_only_mode:
            initial_state = _("Position the head, then click {label}.").format(
                label=self._done_button_label)
        else:
            initial_state = _("Connecting ...")
        self._state_var = tk.StringVar(value=initial_state)
        tk.Label(self, textvariable=self._state_var, bg=DIALOG_BG,
                 font=("Helvetica", 11), wraplength=480, justify="center"
                 ).pack(pady=(0, 2), padx=12)

        self._pos_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._pos_var, bg=DIALOG_BG,
                 fg="#555555", font=("Courier", 10)).pack(pady=(0, 6))

        if not self._is_jog_only_mode:
            # Progress bar + line counter + elapsed time only when there's
            # actually something streaming.
            from tkinter import ttk
            self._progress = ttk.Progressbar(self, length=400, mode='determinate',
                                              maximum=self._total)
            self._progress.pack(padx=20, pady=(0, 4))

            self._progress_text_var = tk.StringVar(
                value=_("Line 0 / {n}").format(n=self._total))
            tk.Label(self, textvariable=self._progress_text_var, bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(pady=(0, 4))

            self._timing_var = tk.StringVar(value=_("Elapsed 0:00"))
            tk.Label(self, textvariable=self._timing_var, bg=DIALOG_BG,
                     fg="#555555", font=("Helvetica", 9)).pack(pady=(0, 10))
        else:
            # Stubs so methods that touch these don't AttributeError.
            self._progress = None
            self._progress_text_var = None
            self._timing_var = None

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 8))
        self._pause_btn = tk.Button(btn_frame, text=_("Pause"),
                                      command=self._on_pause_clicked,
                                      width=10)
        self._resume_btn = tk.Button(btn_frame, text=_("Resume"),
                                       command=self._on_resume_clicked,
                                       width=10, state="disabled")
        if self._show_pause_resume:
            self._pause_btn.pack(side="left", padx=4)
            self._resume_btn.pack(side="left", padx=4)
        # In framing context (loop=True, no cut button), the user's
        # only exit is "Done" — graceful (let the current pass finish
        # so the trailing G0 to bbox-center runs, then close the
        # dialog). In all other contexts the button is "Stop" — abort
        # via soft-reset.
        stop_text = (_("Done") if (loop and not show_cut_button)
                      else _("Stop"))
        stop_bg = "#4caf50" if (loop and not show_cut_button) else None
        stop_fg = "white" if stop_bg else None
        stop_kwargs = {'bg': stop_bg, 'fg': stop_fg,
                        'activebackground': '#388e3c'} if stop_bg else {}
        self._stop_btn = tk.Button(btn_frame, text=stop_text,
                                     command=self._on_stop_clicked,
                                     width=10,
                                     font=("Helvetica", 10, "bold"),
                                     **stop_kwargs)
        self._stop_btn.pack(side="left", padx=4)
        # In jog-only mode (position-the-head dialog), the user needs
        # an explicit Cancel to bail out without proceeding to the
        # next step. Without it, only the X-out (which we now treat
        # as cancel) is available, and the "Start Frame →" button is
        # the only visible action — Matt observed that X-ing out
        # auto-started framing because there was no clearer escape.
        if self._is_jog_only_mode:
            self._cancel_btn = tk.Button(
                btn_frame, text=_("Cancel"),
                command=self._on_cancel_jog_clicked,
                width=10, font=("Helvetica", 10))
            self._cancel_btn.pack(side="left", padx=4)
        # Loop mode shows a prominent "Looks Good — Cut!" button so
        # the user can break out of the repeating framing pass once
        # alignment looks right.
        if self._loop and self._show_cut_button:
            self._cut_btn = tk.Button(
                btn_frame, text=_("Looks Good — Cut!"),
                command=self._on_cut_clicked,
                font=("Helvetica", 10, "bold"),
                bg="#4caf50", fg="white", activebackground="#388e3c")
            self._cut_btn.pack(side="left", padx=8)

        # Jog cluster — useful in G92 mode (the jog moves the cut's
        # reference point) and harmless in auto-framing mode (the cut
        # uses absolute coords so jog only repositions the head for
        # the user's convenience). Sends Grbl $J= via sender.jog(),
        # which Grbl can process even mid-stream.
        jog_frame = tk.LabelFrame(self, text=_("Nudge head"),
                                    bg=DIALOG_BG, padx=8, pady=4)
        jog_frame.pack(pady=(0, 14))
        # Default to 5 mm: on a 400 mm bed, 1 mm is barely
        # perceptible. Users start coarse-jogging into position then
        # switch to 0.5/1 for fine adjustment.
        self._jog_step_var = tk.DoubleVar(value=5.0)
        step_row = tk.Frame(jog_frame, bg=DIALOG_BG)
        step_row.pack(anchor='w')
        tk.Label(step_row, text=_("Step (mm):"), bg=DIALOG_BG,
                  font=("Helvetica", 9)).pack(side='left')
        for step in (0.5, 1.0, 5.0, 25.0):
            tk.Radiobutton(
                step_row, text=str(step), variable=self._jog_step_var,
                value=step, bg=DIALOG_BG, font=("Helvetica", 9),
                highlightthickness=0).pack(side='left')
        btn_grid = tk.Frame(jog_frame, bg=DIALOG_BG)
        btn_grid.pack(pady=2)
        tk.Button(btn_grid, text="↑", width=3,
                   command=lambda: self._on_jog_clicked(0, 1)
                   ).grid(row=0, column=1, padx=2, pady=1)
        tk.Button(btn_grid, text="←", width=3,
                   command=lambda: self._on_jog_clicked(-1, 0)
                   ).grid(row=1, column=0, padx=2, pady=1)
        tk.Button(btn_grid, text="→", width=3,
                   command=lambda: self._on_jog_clicked(1, 0)
                   ).grid(row=1, column=2, padx=2, pady=1)
        tk.Button(btn_grid, text="↓", width=3,
                   command=lambda: self._on_jog_clicked(0, -1)
                   ).grid(row=2, column=1, padx=2, pady=1)
        # Home Laser button — handy when MPos has drifted (lid-open
        # positioning, prior frame stopped mid-trace, etc.) so the user
        # can re-zero without backing out of the dialog. Mid-stream
        # homing aborts the run, so we confirm in that case.
        action_row = tk.Frame(jog_frame, bg=DIALOG_BG)
        action_row.pack(pady=(6, 0))
        self._home_btn = tk.Button(
            action_row, text=_("Home Laser ($H)"),
            command=self._on_home_clicked,
            font=("Helvetica", 9))
        self._home_btn.pack(side='left', padx=2)
        # Try Auto Locate — only shown when caller passed an auto-locate
        # target (a polygon machine coord). Disabled until the laser
        # has been homed in this session.
        self._auto_locate_btn = None
        if self._auto_locate_target is not None:
            self._auto_locate_btn = tk.Button(
                action_row, text=_("Try Auto Locate"),
                command=self._on_auto_locate_clicked,
                font=("Helvetica", 9),
                state=("normal" if self._is_homed else "disabled"))
            self._auto_locate_btn.pack(side='left', padx=2)

        self.protocol("WM_DELETE_WINDOW", self._on_close_window)

        # Tick the timing display once a second
        self._timing_tick()

        # Kick off the stream
        try:
            sender.start_stream(gcode_lines)
            self._state_var.set(_("Running ..."))
        except Exception as e:
            self._state_var.set(_("Failed to start: {e}").format(e=e))
            self._stop_btn.config(text=_("Close"),
                                    command=self._on_close_window)
            self._finished = True

        self.wait_window(self)

    # ------------------------------------------------------------------
    # Tk-thread callbacks (marshalled from FalconSender worker)
    # ------------------------------------------------------------------

    def _on_status(self, status):
        state = status.get('state', '?')
        labels = {
            'Idle': _("Idle"),
            'Run': _("Running"),
            'Hold': _("Paused (feed hold)"),
            'Alarm': _("ALARM"),
            'Door': _("Safety door open"),
            'Jog': _("Jogging"),
            'Home': _("Homing"),
            'Sleep': _("Sleeping"),
        }
        self._state_var.set(labels.get(state, state))
        mpos = status.get('mpos')
        if mpos:
            self._pos_var.set(
                f"X {mpos[0]:7.2f}  Y {mpos[1]:7.2f}  Z {mpos[2]:7.2f}")

    def _on_progress(self, sent, total):
        self._sent = sent
        # Jog-only mode has no progress widgets — skip the updates.
        if self._progress is None or self._progress_text_var is None:
            return
        try:
            self._progress['value'] = sent
        except Exception:
            pass
        pct = (100.0 * sent / total) if total else 0
        self._progress_text_var.set(
            _("Line {s} / {t}  ({pct:.0f}%)").format(s=sent, t=total, pct=pct))

    def _on_error(self, msg):
        self._last_error = msg
        self._state_var.set(_("Error: {m}").format(m=msg))

    def _on_alarm(self, msg):
        self._last_error = msg
        self._state_var.set(_("ALARM: {m}").format(m=msg))

    def _on_done(self, reason):
        # Loop mode: if the pass completed cleanly and the user hasn't
        # clicked "Looks Good — Cut!" yet, restart the stream after
        # making sure Grbl is back in Idle. If we restart while a
        # jog is still finishing (Grbl in JOG state), the next
        # iteration's M3 hits error:9 (G-code locked out during jog).
        if (self._loop and reason == "complete"
                and not self._cut_requested):
            self._pass_count += 1
            self._sent = 0
            self._progress['value'] = 0
            self._state_var.set(
                _("Pass {n} complete — waiting for Idle...").format(
                    n=self._pass_count))
            self._wait_idle_retries = 0
            self.after(300, self._wait_idle_then_restart)
            return
        self._finished = True
        self._final_reason = reason
        # If the user explicitly requested the loop end (clicked Cut
        # or Done in framing context) AND the pass completed cleanly,
        # auto-close the dialog so the caller's next step (the actual
        # cut, or the calibration engrave) starts immediately. Without
        # this, the user has to click Close manually to release
        # wait_window — which looks like "Cut did nothing" because
        # the cut doesn't kick off until this dialog destroys.
        if self._loop and self._cut_requested and reason == "complete":
            if self._on_finished:
                try:
                    self._on_finished(reason)
                except Exception:
                    mlog.exception("laser run: the on_finished callback failed")
            self.destroy()
            return
        if reason == "complete":
            self._state_var.set(_("Complete ✓"))
        elif reason == "stopped":
            self._state_var.set(_("Stopped"))
        elif reason == "error":
            # Keep the Grbl error/alarm visible — without this, the
            # specific code (error:5, ALARM:9, etc.) flashes by and the
            # user only sees a generic "Stopped on error".
            err = self._last_error or _("unknown error")
            decoded = _decode_grbl_error(err)
            failing_line = (
                self._gcode_lines[self._sent]
                if 0 <= self._sent < len(self._gcode_lines) else "")
            txt = _("Stopped on error: {e}").format(e=err)
            if decoded:
                txt += _("  — {d}").format(d=decoded)
            if failing_line:
                txt += _("\nFailing line {n}: {l}").format(
                    n=self._sent + 1, l=failing_line)
            self._state_var.set(txt)
        # Switch buttons: only "Close" remains active. Wrap in
        # TclError guards: the worker thread's on_done can fire after
        # the dialog has been destroyed (user closed the window during
        # the small window between stream end and dialog teardown),
        # and configuring a destroyed widget raises "invalid command
        # name" exceptions that surface as unhandled-exception popups.
        try:
            self._pause_btn.config(state="disabled")
        except tk.TclError:
            pass
        try:
            self._resume_btn.config(state="disabled")
        except tk.TclError:
            pass
        try:
            # Post-stream rewire: button now means "advance to the
            # next step" (e.g. Start Frame →) — NOT the same as
            # X-out / Cancel. Use a dedicated handler that destroys
            # without clobbering _final_reason. Routing this through
            # _on_close_window broke Start Frame because that handler
            # marks the dialog as cancelled in jog-only mode (the
            # X-out / Cancel path), which makes main.py return
            # without proceeding to the framing dialog.
            self._stop_btn.config(
                text=self._done_button_label,
                command=self._on_advance_clicked,
                state="normal" if self._has_jogged else "disabled")
            if not self._has_jogged:
                self._state_var.set(
                    _("Jog the head to your material's bottom-left "
                      "corner, then click {label}.").format(
                        label=self._done_button_label))
        except tk.TclError:
            pass
        if self._on_finished:
            try:
                self._on_finished(reason)
            except Exception:
                mlog.exception("laser run: the on_finished callback failed")

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _on_pause_clicked(self):
        # The buttons follow the command only when it went out (2026-11-10):
        # the sender says False when the byte was not written; None (a
        # stand-in or an older sender) is taken as sent.
        try:
            sent = self._sender.pause()
        except Exception as e:
            mlog.warning("laser run: Pause (feed hold) failed: %s", e)
            sent = False
        if sent is not False:
            self._pause_btn.config(state="disabled")
            self._resume_btn.config(state="normal")

    def _on_resume_clicked(self):
        try:
            sent = self._sender.resume()
        except Exception as e:
            mlog.warning("laser run: Resume failed: %s", e)
            sent = False
        if sent is not False:
            self._pause_btn.config(state="normal")
            self._resume_btn.config(state="disabled")

    def _wait_idle_then_restart(self):
        """Poll Grbl's state until it's Idle (any pending jog has
        finished), then restart the stream. Sending the next iter's
        M3 while Grbl is still in JOG state would trigger error:9
        (G-code locked out during alarm or jog).

        If Cut (or Done) was clicked during this between-pass wait,
        the last pass already finished cleanly: still wait for Idle,
        then end the loop instead of starting another pass. This used
        to bail out on _cut_requested and leave nothing to end the
        loop, so the dialog sat on "finishing current framing pass"
        forever with the head stopped (Matt, 2026-10-26)."""
        if self._finished:
            return
        # Guard against the dialog being destroyed while a deferred
        # call is in flight — accessing tk widgets after destroy
        # raises TclError.
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            st = self._sender.get_status(timeout=0.2)
        except Exception:
            st = None
        state = (st or {}).get('state')
        if state in ('Idle', None):
            if self._cut_requested:
                self._on_done("complete")
                return
            # Settle delay before restart: gives any in-flight jog its
            # ok-byte time to land in the OS serial buffer so
            # start_stream's reset_input_buffer can flush it (without
            # this, a jog issued mid-settle can have its ack arrive
            # after the new worker has started reading, where it's
            # misattributed to a stream-line ack). Empirically 150ms
            # covers Falcon USB latency comfortably.
            self.after(150, self._restart_stream)
            return
        # Still busy (Jog / Run / Hold). Wait a bit and re-check.
        self._wait_idle_retries += 1
        if self._wait_idle_retries > 50:  # ~10 seconds
            try:
                self._state_var.set(
                    _("Timed out waiting for Idle (state={s})").format(s=state))
            except tk.TclError:
                pass
            self._finished = True
            self._final_reason = "error"
            return
        try:
            self.after(200, self._wait_idle_then_restart)
        except tk.TclError:
            pass

    def _restart_stream(self):
        """Re-stream gcode for the next loop iteration. If a
        ``gcode_provider`` callback is set, calls it to get fresh
        gcode (used by the calibration dialog so the framing trace
        follows the head as the user jogs between passes). Bails
        silently if the user has since stopped, and ends the loop if
        Cut was clicked during the settle delay."""
        if self._finished:
            return
        if self._cut_requested:
            self._on_done("complete")
            return
        if self._gcode_provider is not None:
            try:
                self._gcode_lines = self._gcode_provider()
                self._total = len(self._gcode_lines)
                self._progress['maximum'] = max(1, self._total)
            except Exception as e:
                self._state_var.set(_("Regenerate failed: {e}").format(e=e))
                self._finished = True
                self._final_reason = "error"
                return
        try:
            self._sender.start_stream(self._gcode_lines)
        except Exception as e:
            self._state_var.set(_("Restart failed: {e}").format(e=e))
            self._finished = True
            self._final_reason = "error"

    def _remind_lid_closed(self):
        """Ask "Is the lid closed?" once, before the laser comes on, and
        return True to proceed or False to stay put.

        Two buttons, "Yes, Start Frame" and "No, Cancel" (Matt,
        2026-10-26). From 2026-08-05 to 2026-10-26 this was a single-OK
        reminder with nothing to decline, on the reasoning that a yes/no
        gate just gets clicked through; in the shop the single button
        turned out to read as "the app thinks the lid is closed" rather
        than a question, so the answer now has two real options. No
        keeps the jog dialog open: close the lid, click Start Frame
        again, nothing has been consumed. An open lid wedges the stream
        because Grbl blocks motion in Door state (no error returned,
        just an unresponsive controller), which costs a laser power
        cycle AND an app restart to recover.

        It fires once per Frame & Cut run, on **Start Frame** — the
        boundary where the laser first comes on. Framing runs at low
        power, but low power is still power: Grbl enters Door state the
        moment the beam is asked for with the lid up, so the alarm lands
        at the framing pass, not at the cut.

        This took two corrections to get right. It originally fired at
        three boundaries (Start Frame, Cut, and Done during a framing-
        only loop), which trained the user to click it away. On
        2026-08-11 it was cut back to one — but to the *cut*, on the
        reasoning that framing "isn't about to start burning." That
        reasoning confused marking the material with energising the
        laser. By the time the cut prompt appears the user has already
        hit the alarm and needs a laser power cycle and an app restart.
        Moved to Start Frame on 2026-08-18 after Matt hit it in the
        shop. Once framing is running the lid is demonstrably shut, so
        the cut needs no second prompt.

        It also used to try to detect the door via Grbl status (Door
        state or a Pn: 'D' pin) and only ask when it couldn't tell.
        Retired after live-hardware testing on 2026-08-05 found the
        Falcon2 Pro 40W (ESP32-S3 controller) never sends a Pn: field at
        all, lid open or closed — so detection was unverifiable across
        Falcon models and dead code on the one machine tested. Same
        reminder every time is simpler and honest about what the app
        knows.
        """
        dlg = _pkg.LidClosedDialog(self, _("Lid Closed?"), _("Is the lid closed?"),
                              _("Yes, Start Frame"), _("No, Cancel"))
        return bool(dlg.result)

    def _on_cut_clicked(self):
        """User clicked 'Looks Good — Cut!' during loop mode. Flag the
        request; the current pass will be allowed to finish, then the
        loop exits cleanly with reason='complete'.

        No lid reminder here: framing has been running, so the lid is
        already shut. The reminder lives on Start Frame, where the
        laser first comes on."""
        self._cut_requested = True
        try:
            self._cut_btn.config(state='disabled',
                                  text=_("Finishing pass..."))
        except tk.TclError:
            pass
        self._state_var.set(
            _("Cut requested — finishing current framing pass..."))

    def _on_jog_clicked(self, dx_sign, dy_sign):
        """Send a relative jog via Grbl's ``$J=`` real-time command.
        Works during streaming — Grbl handles jog requests outside
        the planner queue. Useful for nudging the head while a
        framing pass is running so the user can fine-tune alignment.
        """
        try:
            step = float(self._jog_step_var.get())
        except Exception:
            step = 1.0
        try:
            sent = self._sender.jog(x=dx_sign * step, y=dy_sign * step,
                                    feed=2000, relative=True)
        except Exception as e:
            mlog.warning("laser run: jog failed: %s", e)
            sent = False
        # If this dialog is gating its done-button on a jog click
        # (MANUAL-mode position confirmation), unlock it now, but only
        # for a jog that went out (2026-11-10): the sender says False when
        # the command was not written; None (a stand-in) counts as sent.
        if sent is not False and not self._has_jogged:
            self._has_jogged = True
            if self._finished:
                try:
                    self._stop_btn.config(state="normal")
                    self._state_var.set(
                        _("Head positioned. Click {label} when ready.").format(
                            label=self._done_button_label))
                except tk.TclError:
                    pass

    def _on_home_clicked(self):
        """Home Laser ($H). Reliable way to recover from drifted MPos —
        lid-open positioning, prior frame aborted mid-trace, the head
        bumped into a stop, etc. Runs the home on a worker thread so
        the Tk loop stays responsive (Grbl's $H blocks until the cycle
        finishes — typically 15-30 seconds on a Falcon)."""
        if not self._finished:
            if not messagebox.askyesno(
                    _("Home Laser?"),
                    _("Homing will abort the run in progress. Continue?"),
                    parent=self):
                return
            self._stop_stream()
        try:
            self._home_btn.config(state="disabled", text=_("Homing..."))
        except tk.TclError:
            return
        self._state_var.set(_("Homing..."))
        result = {}

        def worker():
            try:
                ok, msg = self._sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = msg
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        t = start_worker(worker, name='run-home')

        def check():
            if t.is_alive():
                try:
                    self.after(200, check)
                except tk.TclError:
                    pass
                return
            try:
                self._home_btn.config(state="normal",
                                       text=_("Home Laser ($H)"))
                if result.get('ok'):
                    self._is_homed = True
                    # Tell the parent app so other dialogs in this
                    # session can also see "homed."
                    parent = self.master
                    if hasattr(parent, '_falcon_homed_this_session'):
                        parent._falcon_homed_this_session = True
                    if self._auto_locate_btn is not None:
                        self._auto_locate_btn.config(state="normal")
                    self._state_var.set(_("Homing complete. MPos is now (0, 0)."))
                else:
                    self._state_var.set(_("Homing failed: {m}").format(
                        m=result.get('msg', '?')))
            except tk.TclError:
                pass

        try:
            self.after(200, check)
        except tk.TclError:
            pass

    def _on_auto_locate_clicked(self):
        """Drive the head to the polygon's LB vertex via an absolute
        G0. Available only after the laser has been homed in this
        session — without homing, MPos drift could put the head far
        from the actual target. Counts as a jog for the require-jog
        gate so users can auto-locate and immediately click Start
        Frame without manually nudging."""
        if self._auto_locate_target is None:
            return
        target_x, target_y = self._auto_locate_target
        # Send via the same channel as jog so it serializes with any
        # in-flight stream activity. Use $J for absolute jog — it's
        # cancellable + tracked separately from stream lines.
        try:
            self._sender.jog(x=target_x, y=target_y,
                              feed=3000, relative=False, machine=True)
        except Exception as e:
            self._state_var.set(_("Auto-locate failed: {e}").format(e=e))
            return
        self._state_var.set(_("Driving to polygon BL at MPos "
                              "({x:.1f}, {y:.1f}) …").format(
                                x=target_x, y=target_y))
        if not self._has_jogged:
            self._has_jogged = True
            if self._finished:
                try:
                    self._stop_btn.config(state="normal")
                except tk.TclError:
                    pass

    def _on_cancel_jog_clicked(self):
        """Cancel button in the jog-only dialog. Same effect as the
        window X — sets _final_reason to "cancelled" and closes, so
        main.py knows not to proceed to framing."""
        self._final_reason = "cancelled"
        self.destroy()

    def _on_advance_clicked(self):
        """Post-stream 'advance to next step' button (e.g.
        'Start Frame →' on the jog-to-position dialog).
        _final_reason is already 'complete' from _on_done — leave it
        alone and just destroy so the caller proceeds. Routing this
        through _on_close_window would cancel via the jog-only mode
        path.

        Callers that advance straight into a laser-firing pass set
        remind_lid_on_advance, which puts the lid reminder here — the
        framing pass energises the laser, and an open lid at that
        moment costs a power cycle plus an app restart. "No, Cancel"
        returns without closing: the jog dialog stays up, the user shuts
        the lid and clicks Start Frame again."""
        if self._remind_lid_on_advance and not self._remind_lid_closed():
            return
        self.destroy()

    def _on_stop_clicked(self):
        if self._finished:
            # Post-stream "Stop" button click is the "advance to next
            # step" path (e.g. the "Start Frame →" button on the jog-
            # to-position dialog). _final_reason is already "complete"
            # from _on_done — leave it alone and just destroy so the
            # caller proceeds to the next step. Delegate rather than
            # destroy() directly so the cover check can't be skipped
            # if the _on_done button rewire ever failed.
            self._on_advance_clicked()
            return
        # Graceful "Done" path (framing context, no cut button): let
        # the current pass finish — its trailing G0 moves the head
        # to the bbox center — then exit the loop without a soft-
        # reset. Same mechanism as the Cut button uses.
        graceful = self._loop and not self._show_cut_button
        if graceful:
            # No lid reminder: this only lets the in-flight framing pass
            # finish and then exits. Nothing new is about to fire.
            self._cut_requested = True
            try:
                self._stop_btn.config(state='disabled',
                                       text=_("Finishing pass..."))
            except tk.TclError:
                pass
            self._state_var.set(
                _("Finishing current pass + centering head..."))
            return
        # Aggressive Stop path (cut context): confirm + soft-reset.
        if self._stop_needs_confirm:
            if not messagebox.askyesno(
                    _("Stop?"),
                    _("Stop the job? This will soft-reset the laser "
                      "and clear the planner buffer immediately. "
                      "Material in progress may be ruined."),
                    parent=self):
                return
        self._stop_stream()

    def _stop_stream(self):
        """Stop the run. With a pass streaming, the soft-reset ends it and
        the sender's on_done finishes the dialog. Between framing passes
        nothing is streaming, so no on_done will ever come — finish here,
        or Stop (and Home) do nothing at all."""
        try:
            streaming = self._sender.is_streaming()
        except Exception:
            streaming = True
        try:
            self._sender.stop()
        except Exception:
            mlog.exception("laser run: Stop (soft-reset) failed")
        if not streaming and not self._finished:
            self._on_done("stopped")

    def _on_close_window(self):
        if not self._finished:
            if not messagebox.askyesno(
                    _("Stop and close?"),
                    _("A job is still running. Stop it and close the window?"),
                    parent=self):
                return
            try:
                self._sender.stop()
            except Exception:
                mlog.exception("laser run: stopping the job on window close failed")
        # In jog-only mode, the stream completed instantly (empty
        # gcode list), which set _final_reason = "complete" via
        # _on_done. If the user X's out NOW (or clicks Cancel), we
        # want main.py to treat that as a cancellation, not as
        # "Start Frame was clicked." Override the reason so the
        # caller's `_final_reason != "complete"` check fires.
        if self._is_jog_only_mode:
            self._final_reason = "cancelled"
        self.destroy()

    # ------------------------------------------------------------------
    # Timing tick
    # ------------------------------------------------------------------

    def _timing_tick(self):
        if not self.winfo_exists():
            return
        # Jog-only mode skips the timer widget entirely.
        if self._timing_var is None:
            return
        elapsed = int(time.monotonic() - self._start_time)
        em, es = divmod(elapsed, 60)
        if self._finished:
            self._timing_var.set(_("Elapsed {em}:{es:02d}").format(em=em, es=es))
        elif self._sent > 0 and self._total > 0:
            # Simple linear ETA from progress so far
            est_total = elapsed * (self._total / self._sent)
            remaining = max(0, int(est_total - elapsed))
            rm, rs = divmod(remaining, 60)
            self._timing_var.set(
                _("Elapsed {em}:{es:02d}  ETA ~{rm}:{rs:02d}").format(
                    em=em, es=es, rm=rm, rs=rs))
        else:
            self._timing_var.set(_("Elapsed {em}:{es:02d}").format(em=em, es=es))
        if not self._finished:
            self.after(500, self._timing_tick)
        else:
            # One last tick after done to settle the elapsed display
            self.after(1000, self._timing_tick)


class LiveCameraWindow(tk.Toplevel):
    """Non-modal live camera preview, intended to run ALONGSIDE another
    dialog (typically FalconRunDialog while a Frame & Cut job is in flight)
    so the user can watch the laser head moving on the bed.

    Refresh rate is intentionally slow (5 fps) to keep CPU/disk overhead
    minimal while a cut is happening.
    """

    REFRESH_MS = 200      # 5 fps
    PREVIEW_W = 480
    PREVIEW_H = 360

    def __init__(self, parent, camera_index, title=None, settings=None):
        super().__init__(parent)
        self.title(title or _("Live camera"))
        self.configure(bg=DIALOG_BG)
        # Non-modal — explicitly DON'T grab_set or transient-parent, so
        # the FalconRunDialog stays interactive on top of us.
        self._camera_index = camera_index
        self._settings = settings
        self._cap = None
        self._latest_photo = None
        self._refresh_after_id = None
        self._closed = False
        self._loop_errors = 0            # frames that failed to draw (first logged, rest counted)

        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
        except ImportError:
            tk.Label(self, text=_("Camera preview unavailable"),
                     bg=DIALOG_BG, padx=15, pady=15).pack()
            return

        # Preview label — match the CameraCaptureDialog pattern (no
        # text in the initial label so width/height are interpreted
        # as the placeholder pixel size, not 480 characters wide).
        self._preview_label = tk.Label(
            self, bg="#222222",
            width=self.PREVIEW_W, height=self.PREVIEW_H)
        self._preview_label.pack(padx=8, pady=(8, 4))

        # Switch camera button below the preview — single-line bar. Its
        # left side is the window's status line: empty, unless the window
        # runs on another camera than the saved one (fallback_notice).
        ctrl_row = tk.Frame(self, bg=DIALOG_BG)
        ctrl_row.pack(fill='x', padx=8, pady=(0, 8))
        tk.Button(ctrl_row, text=_("Switch camera"),
                   command=self._on_switch_camera,
                   font=("Helvetica", 9)).pack(side='right')
        self._status_var = tk.StringVar(value="")
        tk.Label(ctrl_row, textvariable=self._status_var, bg=DIALOG_BG,
                 font=("Helvetica", 9), anchor='w').pack(side='left', fill='x', expand=True)

        self.geometry(f"{self.PREVIEW_W + 30}x{self.PREVIEW_H + 60}")
        # Position to the right of the parent so it doesn't overlap the
        # FalconRunDialog that will open after us.
        try:
            px = parent.winfo_x() + parent.winfo_width() + 20
            py = parent.winfo_y()
            self.geometry(f"+{px}+{py}")
        except tk.TclError:
            pass

        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(50, self._open_camera_async)

    def _on_switch_camera(self):
        """Cycle to the next enumerated camera; persist the choice."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            from tkinter import messagebox
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        if self._settings is not None:
            self._settings['camera_index_override'] = int(self._camera_index)
            try:
                from config import save_settings
                save_settings(self._settings)
            except Exception as e:
                mlog.warning("live camera: could not save the chosen "
                             "camera to settings: %s", e)
        # Cancel pending refresh, release current cap, reopen.
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except (tk.TclError, ValueError):
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self.after(50, self._open_camera_async)

    def _open_camera_async(self):

        def _worker():
            asked = self._camera_index
            try:
                cap, used, _frame = self._cam_mod.open_camera_or_fallback(
                    asked, (self._settings or {}).get("camera_index_override"))
                self._camera_index = used
                notice = self._cam_mod.fallback_notice(asked, used) if used != asked else None
                self.after(0, self._camera_ready, cap, notice)
            except Exception as e:
                mlog.warning("live camera preview: could not open camera "
                             "#%s: %s", asked, e)

        start_worker(_worker, name='livecam-open')

    def _camera_ready(self, cap, notice=None):
        if self._closed:
            try:
                cap.release()
            except Exception:
                pass
            return
        self._cap = cap
        try:
            self._status_var.set(notice or "")
        except tk.TclError:
            pass
        self._refresh_loop()

    def _refresh_loop(self):
        if self._closed or self._cap is None or not self.winfo_exists():
            return
        import cv2
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok = False
        if ok and frame is not None:
            try:
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = self._PIL_Image.fromarray(rgb)
                img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
                self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
                self._preview_label.config(
                    image=self._latest_photo,
                    width=img.width, height=img.height)
            except Exception:
                # The calibration preview's rule (2026-11-10): one bad frame
                # must not end the loop (it used to freeze this window for
                # good) and must not go unsaid: the first failure logs its
                # traceback, the rest are counted and said at close.
                self._loop_errors += 1
                if self._loop_errors == 1:
                    mlog.exception("live camera: drawing a frame failed "
                                   "(further failures are counted)")
        self._refresh_after_id = self.after(self.REFRESH_MS, self._refresh_loop)

    def close(self):
        """Programmatic close — called by the parent when the job is done."""
        self._closed = True
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except (tk.TclError, ValueError):
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        if self._loop_errors:
            mlog.warning("live camera: %d frames failed to draw", self._loop_errors)
            self._loop_errors = 0
        try:
            self.destroy()
        except tk.TclError:
            pass
