"""The nesting preview and the Speed & Power Test preview.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import messagebox, ttk
import math


from config import (
    save_settings,
)
from svg_engine import (
    placed_disc_area,
    feeds_speeds_label_geometry,
    adhesive_ring_geometry, adhesive_ring_number_strokes,
)

from .common import DIALOG_BG, _PREVIEW_COLORS, log


class NestingPreviewWindow(tk.Toplevel):
    """Preview window showing the nested pad layout before file generation.

    Shows the sheet boundary (or polygon) with circles at their nested
    positions. User can proceed to save or go back to adjust settings.

    Result:
        self.result = "save" if user clicks Save, None if Adjust/close.
    """

    def __init__(self, parent, placements, width_mm, height_mm, polygon=None,
                 proceed_label=None, zones=None, settings=None):
        """
        Args:
            parent: Parent tk widget
            placements: dict of {material: [(pad_size, cx, cy, r), ...]}
            width_mm: Sheet width in mm
            height_mm: Sheet height in mm
            polygon: Optional list of (x, y) polygon points in mm
            proceed_label: Text for the confirm button (default "Save Files").
                Frame & Cut passes "Continue →" since it streams to the
                laser rather than writing files. result is "save" either way.
            zones: Optional positioned zone dicts from nest_with_zones, drawn
                as bordered labeled blocks so the preview matches the cut.
            settings: App settings. Only adhesive rings need them: a ring's
                radius in `placements` is its nesting footprint, so the real
                ring and its number are drawn from adhesive_ring_geometry.
                Without settings a ring job falls back to footprint circles.
        """
        super().__init__(parent)
        self.title(_("Nesting Preview"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self._zones = list(zones or [])
        self._settings = settings

        # First-run tutorial (persisted in settings via parent app)
        app = getattr(parent, 'settings', None) if not isinstance(parent, dict) else None
        if app is None:
            # Try to get settings from the root app
            try:
                app = parent.master.settings if hasattr(parent, 'master') else None
            except Exception:
                app = None

        # Use a simple attribute check — the actual persistence happens
        # via the seen_preview_tutorial setting
        show_tutorial = False
        try:
            if hasattr(parent, 'settings') and not parent.settings.get("seen_preview_tutorial"):
                show_tutorial = True
                parent.settings["seen_preview_tutorial"] = True
                save_settings(parent.settings)
        except Exception as e:
            log.warning("could not save that the nesting preview tutorial "
                        "was seen: %s", e)

        if show_tutorial:
            messagebox.showinfo(_("Nesting Preview"),
                _("This preview shows how your pads will be arranged "
                "on the sheet before files are generated.\n\n"
                "If the layout looks good, click Save to write the files.\n\n"
                "If not, click Adjust to go back and change:\n"
                "  \u2022 Edge bias (pack toward a different edge)\n"
                "  \u2022 Sheet dimensions\n"
                "  \u2022 Custom polygon shape\n"
                "  \u2022 Pad sizes or quantities\n\n"
                "Preview works with one material at a time."),
                parent=self)

        self._placements = placements
        self._width_mm = width_mm
        self._height_mm = height_mm
        self._polygon = polygon
        self._materials = list(placements.keys())

        # Material selector (if multiple materials)
        if len(self._materials) > 1:
            sel_frame = tk.Frame(self, bg=DIALOG_BG)
            sel_frame.pack(pady=(10, 0))
            tk.Label(sel_frame, text=_("Material:"), bg=DIALOG_BG,
                     font=("Helvetica", 10)).pack(side="left", padx=(0, 5))
            self._mat_var = tk.StringVar(value=self._materials[0])
            mat_combo = ttk.Combobox(
                sel_frame, textvariable=self._mat_var,
                values=self._materials, state="readonly", width=12)
            mat_combo.pack(side="left")
            mat_combo.bind("<<ComboboxSelected>>", lambda e: self._draw())
        else:
            self._mat_var = tk.StringVar(value=self._materials[0])

        # Canvas
        self._canvas = tk.Canvas(self, bg="white", highlightthickness=1,
                                  highlightbackground="#999999")
        self._canvas.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        self._canvas.bind("<Configure>", self._on_resize)

        # Info label (updated per material)
        self._info_label = tk.Label(self, text="", bg=DIALOG_BG,
                                     font=("Helvetica", 9))
        self._info_label.pack(pady=(0, 5))

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 10))
        tk.Button(btn_frame, text=proceed_label or _("Save Files"), command=self._on_save,
                  font=("Helvetica", 10, "bold"), width=12).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Adjust"), command=self._on_adjust,
                  font=("Helvetica", 10), width=12).pack(side="left", padx=5)

        # Size the window
        self.geometry("700x550")
        self.minsize(400, 300)

        self.protocol("WM_DELETE_WINDOW", self._on_adjust)
        self.wait_window(self)

    def _on_save(self):
        self.result = "save"
        self.destroy()

    def _on_adjust(self):
        self.result = None
        self.destroy()

    def _on_resize(self, event=None):
        self._draw()

    def _draw(self):
        cv = self._canvas
        cv.delete("all")

        cw = cv.winfo_width()
        ch = cv.winfo_height()
        if cw < 20 or ch < 20:
            return

        margin = 20
        draw_w = cw - 2 * margin
        draw_h = ch - 2 * margin

        # Scale to fit sheet in canvas, maintaining aspect ratio.
        # When a polygon is loaded the nest packs in the polygon's own
        # coordinate space (its bounding box), which may not match the
        # user's typed width/height fields — typically a polygon drawn
        # at e.g. 0..200mm with a 38x76 sheet typed in. Use the polygon's
        # actual bbox for canvas scaling so the contents land in view.
        if self._polygon:
            poly_xs = [p[0] for p in self._polygon]
            poly_ys = [p[1] for p in self._polygon]
            sheet_w = max(poly_xs) - min(poly_xs)
            sheet_h = max(poly_ys) - min(poly_ys)
        else:
            sheet_w = self._width_mm
            sheet_h = self._height_mm
        scale_x = draw_w / sheet_w if sheet_w > 0 else 1
        scale_y = draw_h / sheet_h if sheet_h > 0 else 1
        scale = min(scale_x, scale_y)

        # Center the sheet in the canvas
        offset_x = margin + (draw_w - sheet_w * scale) / 2
        offset_y = margin + (draw_h - sheet_h * scale) / 2

        # Draw sheet boundary
        if self._polygon:
            pts = []
            for px, py in self._polygon:
                pts.extend([offset_x + px * scale, offset_y + py * scale])
            cv.create_polygon(pts, fill="#F8F8F0", outline="#888888", width=2)
        else:
            cv.create_rectangle(
                offset_x, offset_y,
                offset_x + sheet_w * scale,
                offset_y + sheet_h * scale,
                fill="#F8F8F0", outline="#888888", width=2)

        # Draw pads for selected material only
        material = self._mat_var.get()
        placed = self._placements.get(material, [])
        color = _PREVIEW_COLORS.get(material, '#888888')
        fill = self._lighten(color, 0.7)

        # Zone borders and labels go under the pads, matching the cut order.
        for zone in self._zones:
            zx = offset_x + zone['x'] * scale
            zy = offset_y + zone['y'] * scale
            zw = zone['w'] * scale
            zh = zone['h'] * scale
            cv.create_rectangle(zx, zy, zx + zw, zy + zh,
                                outline="#B0782A", width=1, fill="#FFFDF5")
            lx = zx + zw / 2
            ly = zy + (zone['border'] + zone['font'] / 2) * scale
            label_px = max(7, min(14, int(zone['font'] * scale)))
            cv.create_text(lx, ly, text=zone['label'], fill="#B0782A",
                           font=("Helvetica", label_px, "bold"))

        ring_job = material == 'adhesive_ring' and self._settings is not None
        ring_area = 0.0
        for pad_size, cx, cy, r in placed:
            sx = offset_x + cx * scale
            sy = offset_y + cy * scale
            sr = r * scale

            if ring_job:
                # The real part, not the footprint that nested: outer cut,
                # hole, and the number where it will be traced.
                geom = adhesive_ring_geometry(pad_size, self._settings)
                ro = geom["outer_r"] * scale
                cv.create_oval(sx - ro, sy - ro, sx + ro, sy + ro,
                               fill=fill, outline=color, width=1, tags=("ring",))
                if geom["inner_r"] > 0:
                    ri = geom["inner_r"] * scale
                    cv.create_oval(sx - ri, sy - ri, sx + ri, sy + ri,
                                   fill="#F8F8F0", outline=color, width=1,
                                   tags=("ring", "ring_inner"))
                for stroke in adhesive_ring_number_strokes(geom, cx, cy):
                    pts = [c for x, y in stroke
                           for c in (offset_x + x * scale, offset_y + y * scale)]
                    cv.create_line(*pts, fill=color, width=1, tags=("ring_number",))
                ring_area += math.pi * geom["outer_r"] ** 2
                continue

            cv.create_oval(sx - sr, sy - sr, sx + sr, sy + sr,
                           fill=fill, outline=color, width=1)

            # Label with pad size (skip if too small to read)
            if sr > 12:
                font_size = max(6, min(11, int(sr * 0.5)))
                cv.create_text(sx, sy, text=f"{pad_size:.1f}",
                               fill=color, font=("Helvetica", font_size))

        # Usage percentage for this material
        sheet_area = sheet_w * sheet_h
        if self._polygon:
            pts = self._polygon
            n = len(pts)
            area = 0
            for i in range(n):
                j = (i + 1) % n
                area += pts[i][0] * pts[j][1]
                area -= pts[j][0] * pts[i][1]
            sheet_area = abs(area) / 2

        pad_area = placed_disc_area(placed)   # same figure the optimizer ranks by
        if ring_job:
            pad_area = ring_area              # tape cut, not the nesting footprints
        usage_pct = pad_area / sheet_area * 100 if sheet_area > 0 else 0

        cv.create_text(cw - margin, ch - 5,
                       text=_("{pct:.0f}% used").format(pct=usage_pct),
                       fill="#666666", font=("Helvetica", 9),
                       anchor="se")

        # Update info label
        self._info_label.configure(
            text=_("{count} pads ({material}) on {w:.0f} x {h:.0f} mm{shape}").format(count=len(placed), material=material.replace('_', ' '), w=sheet_w, h=sheet_h, shape=_(' (custom shape)') if self._polygon else ''))

    @staticmethod
    def _lighten(hex_color, factor):
        """Lighten a hex color toward white."""
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        r = int(r + (255 - r) * factor)
        g = int(g + (255 - g) * factor)
        b = int(b + (255 - b) * factor)
        return f"#{r:02x}{g:02x}{b:02x}"


class SpeedPowerTestPreview(tk.Toplevel):
    """Preview window for the Tooling-tab Speed & Power Test.

    Mirrors NestingPreviewWindow's shape (modal canvas, Save/Adjust buttons)
    but renders the test-disc grid: each disc shows its 2-char numeric ID,
    and when "Also test with air off" is enabled the two air states are
    color-coded so the user can spot at a glance which discs are which.

    Result:
        self.result = "save" if user clicks Save, None if Adjust/close.

    proceed_label replaces "Save G-code" on the go-ahead button. Tooling's
    Frame & Cut passes "Continue →": it streams to the laser and saves no
    G-code, so "Save G-code" would mislead (the same reason
    NestingPreviewWindow takes one).
    """

    AIR_ON_COLOR = "#3a78c2"   # blue outline
    AIR_OFF_COLOR = "#cc7a3a"  # orange outline
    NEUTRAL_COLOR = "#555555"  # dark gray when air doesn't vary

    def __init__(self, parent, test_pieces, sheet_w_mm, sheet_h_mm,
                  eng_speed, eng_power, material_display, proceed_label=None):
        super().__init__(parent)
        self.title(_("Speed & Power Test Preview"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self._pieces = test_pieces
        self._sheet_w = sheet_w_mm
        self._sheet_h = sheet_h_mm
        self._eng_speed = eng_speed
        self._eng_power = eng_power
        self._material_display = material_display

        air_states = {p.get('air_assist', True) for p in test_pieces}
        self._air_varies = len(air_states) > 1

        # Summary line(s) at top
        sample = test_pieces[0] if test_pieces else None
        if sample is not None:
            speeds = sorted({p['speed'] for p in test_pieces})
            powers = sorted({p['power'] for p in test_pieces})
            passes_vals = sorted({p['passes'] for p in test_pieces})

            def _rng(vals):
                if len(vals) == 1:
                    return str(vals[0])
                return f"{vals[0]}-{vals[-1]}"

            summary_text = _("{n} discs - Speed {sp} mm/min, Power {pw}%, "
                              "Passes {ps}").format(
                n=len(test_pieces),
                sp=_rng(speeds),
                pw=_rng(powers),
                ps=_rng(passes_vals))
        else:
            summary_text = ""

        tk.Label(self, text=summary_text, bg=DIALOG_BG,
                 font=("Helvetica", 10, "bold")).pack(pady=(10, 2))

        eng_line = _("Material: {m}   Sheet: {w:.0f} x {h:.0f} mm   "
                      "Engraving: {es} mm/min @ {ep}%").format(
            m=material_display, w=sheet_w_mm, h=sheet_h_mm,
            es=int(eng_speed), ep=int(eng_power))
        tk.Label(self, text=eng_line, bg=DIALOG_BG,
                 font=("Helvetica", 9), fg="#555555").pack(pady=(0, 4))

        if self._air_varies:
            legend_frame = tk.Frame(self, bg=DIALOG_BG)
            legend_frame.pack(pady=(0, 4))
            tk.Label(legend_frame, text="●", bg=DIALOG_BG,
                     fg=self.AIR_ON_COLOR,
                     font=("Helvetica", 14)).pack(side="left")
            tk.Label(legend_frame, text=_("air ON  "), bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(side="left")
            tk.Label(legend_frame, text="●", bg=DIALOG_BG,
                     fg=self.AIR_OFF_COLOR,
                     font=("Helvetica", 14)).pack(side="left")
            tk.Label(legend_frame, text=_("air OFF"), bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(side="left")

        # Canvas
        self._canvas = tk.Canvas(self, bg="white", highlightthickness=1,
                                  highlightbackground="#999999")
        self._canvas.pack(fill="both", expand=True, padx=10, pady=(2, 5))
        self._canvas.bind("<Configure>", lambda e: self._draw())

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 10))
        self._proceed_btn = tk.Button(
            btn_frame, text=proceed_label or _("Save G-code"), command=self._on_save,
            font=("Helvetica", 10, "bold"), width=12)
        self._proceed_btn.pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Adjust"), command=self._on_adjust,
                  font=("Helvetica", 10), width=12).pack(side="left", padx=5)

        self.geometry("720x600")
        self.minsize(420, 320)
        self.protocol("WM_DELETE_WINDOW", self._on_adjust)
        self.wait_window(self)

    def _on_save(self):
        self.result = "save"
        self.destroy()

    def _on_adjust(self):
        self.result = None
        self.destroy()

    def _draw(self):
        cv = self._canvas
        cv.delete("all")

        cw = cv.winfo_width()
        ch = cv.winfo_height()
        if cw < 20 or ch < 20:
            return

        margin = 20
        draw_w = cw - 2 * margin
        draw_h = ch - 2 * margin
        scale_x = draw_w / self._sheet_w if self._sheet_w > 0 else 1
        scale_y = draw_h / self._sheet_h if self._sheet_h > 0 else 1
        scale = min(scale_x, scale_y)

        offset_x = margin + (draw_w - self._sheet_w * scale) / 2
        offset_y = margin + (draw_h - self._sheet_h * scale) / 2

        # Sheet outline
        cv.create_rectangle(
            offset_x, offset_y,
            offset_x + self._sheet_w * scale,
            offset_y + self._sheet_h * scale,
            fill="#F8F8F0", outline="#888888", width=2)

        for piece in self._pieces:
            cx = offset_x + piece['cx'] * scale
            cy = offset_y + piece['cy'] * scale
            sr = (piece['diameter'] / 2.0) * scale
            inner = piece.get('inner_diameter', 0) or 0
            if self._air_varies:
                color = (self.AIR_ON_COLOR
                          if piece.get('air_assist', True)
                          else self.AIR_OFF_COLOR)
            else:
                color = self.NEUTRAL_COLOR
            fill = _PREVIEW_COLORS.get('felt', '#dddddd')  # subtle light fill
            fill = self._lighten_hex(color, 0.85)
            cv.create_oval(cx - sr, cy - sr, cx + sr, cy + sr,
                           fill=fill, outline=color, width=1)
            # Center hole (washer / shim): punch it out of the fill in white
            # so the ring reads as solid material, matching the cut output.
            if inner > 0:
                ir = (inner / 2.0) * scale
                cv.create_oval(cx - ir, cy - ir, cx + ir, cy + ir,
                               fill="white", outline=color, width=1)
            # ID label — centered for a solid disc, dropped into the lower ring
            # for a washer (mirrors feeds_speeds_label_geometry / the G-code).
            label_dy_mm, _font_mm = feeds_speeds_label_geometry(
                piece['diameter'], inner)
            label_y = cy + label_dy_mm * scale
            if sr > 8:
                font_size = max(7, min(12, int(sr * 0.6)))
                if inner > 0:
                    # Keep the number inside the ring band on screen too.
                    ring_px = sr - (inner / 2.0) * scale
                    font_size = max(6, min(font_size, int(ring_px * 0.7)))
                cv.create_text(cx, label_y, text=piece['id'],
                               fill=color, font=("Helvetica", font_size, "bold"))

    @staticmethod
    def _lighten_hex(hex_color, factor):
        h = hex_color.lstrip('#')
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        r = int(r + (255 - r) * factor)
        g = int(g + (255 - g) * factor)
        b = int(b + (255 - b) * factor)
        return f"#{r:02x}{g:02x}{b:02x}"
