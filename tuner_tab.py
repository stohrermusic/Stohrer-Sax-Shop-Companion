"""
Strobe Tuner tab mixin for Stohrer Sax Shop Companion.

Renders a 12-wheel chromatic strobe tuner modeled after the Peterson
Stroboconn 6T-5. Each wheel shows 7 concentric rings of alternating
colored/dark segments visible through a wedge-shaped cutout. Phase
tracking from the audio engine drives the stroboscopic rotation effect.

Requires: numpy, sounddevice (graceful fallback if unavailable)
"""

import tkinter as tk
from tkinter import ttk, colorchooser
import math
import sys

IS_MACOS = sys.platform == 'darwin'

try:
    import numpy as np
    from tuner_engine import (
        TunerEngine, ReferencePlayer, AUDIO_AVAILABLE,
        PITCH_CLASSES, MIN_OCTAVE, MAX_OCTAVE,
    )
    _TUNER_IMPORTS_OK = True
except ImportError:
    _TUNER_IMPORTS_OK = False
    AUDIO_AVAILABLE = False
    PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']


# ============================================
# CONSTANTS
# ============================================

DARK_BG = "#1A1A1A"
WHEEL_BG = "#0D0D0D"
DIM_MULTIPLIER = 0.2      # Inactive wheel brightness
LABEL_COLOR = "#888888"
LABEL_ACTIVE_COLOR = "#FFFFFF"
FRAME_RATES = {"60": 16, "90": 11, "120": 8}

# Ring segment counts (7 rings, doubling — matches patent)
RING_SEGMENTS = [4, 8, 16, 32, 64, 128, 256]
NUM_RINGS = len(RING_SEGMENTS)

# Wedge cutout parameters
WEDGE_ANGLE = 80.0   # Total visible arc in degrees

# Transposition label shifts (concert pitch class 0=C → displayed label)
TRANSPOSITION_SHIFTS = {
    "C": 0,
    "Eb": 9,
    "Bb": 2,
    "F": 7,
}
TRANSPOSITION_KEYS = list(TRANSPOSITION_SHIFTS.keys())  # C, Eb, Bb, F

# Reference tone note names and frequencies (concert pitch, A=440)
def _build_ref_notes(ref_pitch=440.0):
    """Build list of (display_name, frequency) for reference tone selector."""
    notes = []
    for octave in range(3, 7):
        for pc_idx, name in enumerate(PITCH_CLASSES):
            semitones = (pc_idx - 9) + (octave - 4) * 12
            freq = ref_pitch * (2.0 ** (semitones / 12.0))
            notes.append((f"{name}{octave}", freq))
    return notes


# ============================================
# HELPER: dim a hex color
# ============================================

def _scale_color(hex_color, factor):
    """Return hex_color scaled by factor (0.0 = black, 1.0 = full brightness)."""
    hex_color = hex_color.lstrip('#')
    r = min(255, int(int(hex_color[0:2], 16) * factor))
    g = min(255, int(int(hex_color[2:4], 16) * factor))
    b = min(255, int(int(hex_color[4:6], 16) * factor))
    return f"#{r:02x}{g:02x}{b:02x}"


# ============================================
# WHEEL RENDERING
# ============================================

def _annular_sector_points(cx, cy, r_inner, r_outer, angle_start, angle_end, steps=6):
    """Compute polygon points for an annular sector (arc-shaped wedge).

    Angles in degrees, counterclockwise from east (standard math convention).
    Returns flat list [x0, y0, x1, y1, ...] for canvas.create_polygon().
    """
    points = []
    # Outer arc (start to end)
    for i in range(steps + 1):
        t = angle_start + (angle_end - angle_start) * i / steps
        rad = math.radians(t)
        points.append(cx + r_outer * math.cos(rad))
        points.append(cy - r_outer * math.sin(rad))
    # Inner arc (end to start, reversed)
    for i in range(steps + 1):
        t = angle_end + (angle_start - angle_end) * i / steps
        rad = math.radians(t)
        points.append(cx + r_inner * math.cos(rad))
        points.append(cy - r_inner * math.sin(rad))
    return points


class StrobeWheel:
    """One of the 12 strobe disc wheels."""

    def __init__(self, canvas, cx, cy, radius, stripe_color, direction="up"):
        """Create a strobe wheel.

        Args:
            direction: "up" = wedge opens upward (top row), "down" = opens downward (bottom row)
        """
        self.canvas = canvas
        self.cx = cx
        self.cy = cy
        self.radius = radius
        self.stripe_color = stripe_color
        self.direction = direction
        self._brightness = 0.0       # 0.0 = dark, 1.0 = full brightness
        self._last_ring_fills = None  # Cache to avoid redundant itemconfigure calls
        self._phase_offset = 0.0

        # Wedge center angle: 90° = up, 270° = down
        if direction == "up":
            self._wedge_center = 90.0
        else:
            self._wedge_center = 270.0

        # Ring radii — evenly spaced from center to outer radius
        gap = radius * 0.12  # Gap at center (first ring starts here)
        ring_width = (radius - gap) / NUM_RINGS
        self._ring_radii = []
        for i in range(NUM_RINGS):
            r_inner = gap + i * ring_width
            r_outer = gap + (i + 1) * ring_width - ring_width * 0.05  # Tiny gap between rings
            self._ring_radii.append((r_inner, r_outer))

        # Pre-create segment polygons
        # segments[ring_idx] = list of (polygon_id, base_start_angle, segment_span)
        self._segments = []
        self._create_segments()

        # Create masking overlay (covers everything outside the wedge)
        self._create_mask()

        # Note label (drawn on top)
        self._label_id = canvas.create_text(
            cx, cy + radius + 12,
            text="", fill=LABEL_COLOR,
            font=("Helvetica", 10, "bold"),
            anchor="center"
        )

    def _create_segments(self):
        """Create all segment polygons for this wheel (full disc, masked later)."""
        self._segments = []
        for ring_idx in range(NUM_RINGS):
            ring_segs = []
            n_total = RING_SEGMENTS[ring_idx]
            seg_span = 360.0 / n_total  # Degrees per segment
            r_inner, r_outer = self._ring_radii[ring_idx]

            # Only create the colored segments (every other one)
            for seg_i in range(n_total):
                if seg_i % 2 == 0:  # Colored segment
                    base_start = seg_i * seg_span
                    # Determine arc detail based on segment size
                    steps = max(2, min(8, int(seg_span / 3)))
                    points = _annular_sector_points(
                        self.cx, self.cy, r_inner, r_outer,
                        base_start, base_start + seg_span, steps
                    )
                    initial_fill = _scale_color(self.stripe_color, DIM_MULTIPLIER)
                    poly_id = self.canvas.create_polygon(
                        points, fill=initial_fill, outline='', width=0
                    )
                    ring_segs.append((poly_id, base_start, seg_span, steps))

            self._segments.append(ring_segs)

    def _create_mask(self):
        """Create wedge-shaped mask matching the Stroboconn cutout.

        The visible area is a sector of the disc from the first ring outward.
        Top row wheels open upward, bottom row wheels open downward —
        both pointing away from the label band in the center.
        """
        cx, cy = self.cx, self.cy
        r = self.radius
        mr = r + 8  # Mask extends slightly beyond disc edge

        # Wedge sector: centered on _wedge_center, spanning WEDGE_ANGLE
        wedge_start = self._wedge_center - WEDGE_ANGLE / 2
        wedge_end = self._wedge_center + WEDGE_ANGLE / 2

        # Outer mask: large sector covering the NON-visible portion
        mask_start = wedge_end
        mask_end = wedge_start + 360.0
        steps = 50

        points = [cx, cy]  # Center point
        for i in range(steps + 1):
            t = mask_start + (mask_end - mask_start) * i / steps
            rad = math.radians(t)
            points.append(cx + mr * math.cos(rad))
            points.append(cy - mr * math.sin(rad))

        self._mask_id = self.canvas.create_polygon(
            points, fill=DARK_BG, outline='', width=0
        )

        # Inner mask: covers the center gap (no rings visible there)
        inner_r = self._ring_radii[0][0]  # Inner radius of first ring
        inner_pts = []
        for i in range(32):
            a = i * 2 * math.pi / 32
            inner_pts.append(cx + inner_r * math.cos(a))
            inner_pts.append(cy - inner_r * math.sin(a))
        self._inner_mask = self.canvas.create_polygon(
            inner_pts, fill=DARK_BG, outline='', width=0
        )

    def set_label(self, text):
        """Set the note name label."""
        self.canvas.itemconfigure(self._label_id, text=text)

    def set_color(self, hex_color):
        """Update the stripe color."""
        self.stripe_color = hex_color
        self._last_ring_fills = None  # Force fill refresh
        fill = _scale_color(hex_color, DIM_MULTIPLIER + self._brightness * (1.0 - DIM_MULTIPLIER))
        for ring_segs in self._segments:
            for poly_id, _, _, _ in ring_segs:
                self.canvas.itemconfigure(poly_id, fill=fill)

    def update(self, phase_offset, magnitude, ring_magnitudes=None):
        """Update wheel for one animation frame.

        Args:
            phase_offset: Rotation angle in degrees (from engine phase tracking)
            magnitude: Overall signal strength 0.0-1.0 (drives label brightness)
            ring_magnitudes: Optional list of per-ring magnitudes (0.0-1.0).
                If provided, each ring gets independent brightness matching
                the real Stroboconn stroboscopic effect where the played
                octave's ring appears bright while others are dim.
        """
        # Compute overall brightness for label
        brightness = min(1.0, magnitude ** 0.6) if magnitude > 0.05 else 0.0

        # Per-ring brightness from real spectral data
        if ring_magnitudes:
            ring_fills = []
            for ring_idx in range(NUM_RINGS):
                rm = min(1.0, ring_magnitudes[ring_idx])
                rb = min(1.0, rm ** 0.6) if rm > 0.05 else 0.0
                rf = DIM_MULTIPLIER + rb * (1.0 - DIM_MULTIPLIER)
                ring_fills.append(_scale_color(self.stripe_color, rf))
        else:
            fill_factor = DIM_MULTIPLIER + brightness * (1.0 - DIM_MULTIPLIER)
            ring_fills = [_scale_color(self.stripe_color, fill_factor)] * NUM_RINGS

        # Update fill colors per ring (only if changed)
        if ring_fills != self._last_ring_fills:
            self._last_ring_fills = ring_fills
            for ring_idx, ring_segs in enumerate(self._segments):
                fill = ring_fills[ring_idx]
                for poly_id, _, _, _ in ring_segs:
                    self.canvas.itemconfigure(poly_id, fill=fill)
            # Label brightness tracks overall magnitude
            label_color = _scale_color("#FFFFFF", 0.35 + brightness * 0.65)
            self.canvas.itemconfigure(self._label_id, fill=label_color)

        self._brightness = brightness

        # Update segment positions (rotate by phase_offset)
        if abs(phase_offset - self._phase_offset) < 0.01 and brightness < 0.01:
            return  # No visible change, skip coordinate update

        self._phase_offset = phase_offset

        for ring_idx in range(NUM_RINGS):
            r_inner, r_outer = self._ring_radii[ring_idx]
            for poly_id, base_start, seg_span, steps in self._segments[ring_idx]:
                start = base_start + phase_offset
                end = start + seg_span
                points = _annular_sector_points(
                    self.cx, self.cy, r_inner, r_outer,
                    start, end, steps
                )
                self.canvas.coords(poly_id, *points)

        # Re-raise mask to stay on top
        self.canvas.tag_raise(self._mask_id)
        self.canvas.tag_raise(self._inner_mask)
        self.canvas.tag_raise(self._label_id)


# ============================================
# TUNER TAB MIXIN
# ============================================

class TunerTabMixin:
    """Mixin class that adds the Strobe Tuner tab to the main application."""

    def _init_tuner_state(self):
        """Initialize tuner state. Called from __init__."""
        self._tuner_engine = None
        self._tuner_player = None
        self._tuner_wheels = []
        self._tuner_running = False
        self._tuner_anim_id = None

    def create_tuner_tab(self, parent):
        """Build the Strobe Tuner tab UI."""
        if not _TUNER_IMPORTS_OK or not AUDIO_AVAILABLE:
            self._create_tuner_fallback(parent)
            return

        tuner_settings = self.settings.get("tuner_settings", {})
        bg = DARK_BG

        # --- Main container (skip theme walker — dark display) ---
        main_frame = tk.Frame(parent, bg=bg)
        main_frame._skip_theme = True
        main_frame.pack(fill="both", expand=True)

        # --- Canvas for strobe wheels ---
        self._tuner_canvas = tk.Canvas(
            main_frame, bg=DARK_BG, highlightthickness=0,
            borderwidth=0,
        )
        self._tuner_canvas._dark_canvas = True  # Skip theme walker
        self._tuner_canvas.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        # --- Control panel ---
        ctrl_bg = "systemWindowBackgroundColor" if IS_MACOS else "#2A2A2A"
        ctrl_fg = "white" if not IS_MACOS else "systemTextColor"

        ctrl_frame = tk.Frame(main_frame, bg=ctrl_bg, padx=10, pady=8)
        ctrl_frame.pack(fill="x", padx=5, pady=5)

        # Row 1: Transposition + Reference Pitch + Sensitivity
        row1 = tk.Frame(ctrl_frame, bg=ctrl_bg)
        row1.pack(fill="x", pady=(0, 4))

        tk.Label(row1, text="Instrument in Key of:", bg=ctrl_bg, fg=ctrl_fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 4))

        # Migrate old lowercase keys to new display format
        _trans_migrate = {"concert": "C", "bb": "Bb", "eb": "Eb", "f": "F"}
        saved_trans = tuner_settings.get("transposition", "C")
        saved_trans = _trans_migrate.get(saved_trans, saved_trans)
        self._tuner_transpose_var = tk.StringVar(value=saved_trans)
        transpose_combo = ttk.Combobox(
            row1, textvariable=self._tuner_transpose_var,
            values=TRANSPOSITION_KEYS,
            state="readonly", width=4
        )
        transpose_combo.pack(side="left", padx=(0, 15))
        transpose_combo.bind("<<ComboboxSelected>>", lambda e: self._tuner_update_labels())

        tk.Label(row1, text="A =", bg=ctrl_bg, fg=ctrl_fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 2))

        self._tuner_pitch_var = tk.StringVar(
            value=str(tuner_settings.get("reference_pitch", 440.0)))
        pitch_spin = tk.Spinbox(
            row1, textvariable=self._tuner_pitch_var,
            from_=400.0, to=480.0, increment=0.1,
            width=6, font=("Helvetica", 10),
            command=self._tuner_on_pitch_changed
        )
        pitch_spin.pack(side="left", padx=(0, 4))
        tk.Label(row1, text="Hz", bg=ctrl_bg, fg=ctrl_fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 15))

        tk.Label(row1, text="Sensitivity:", bg=ctrl_bg, fg=ctrl_fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 4))

        self._tuner_sens_var = tk.IntVar(
            value=tuner_settings.get("sensitivity", 50))
        sens_scale = tk.Scale(
            row1, variable=self._tuner_sens_var,
            from_=0, to=100, orient="horizontal",
            length=120, showvalue=False,
            bg=ctrl_bg, fg=ctrl_fg, troughcolor="#444444",
            highlightthickness=0,
            command=self._tuner_on_sensitivity_changed
        )
        sens_scale.pack(side="left", padx=(0, 15))

        # Settings initialized here, configured via Options > Settings dialog
        self._tuner_color = tuner_settings.get("stripe_color", "#00FF00")
        self._tuner_fps_var = tk.StringVar(
            value=str(tuner_settings.get("fps", "60")))

        # Row 2: Reference tone
        row2 = tk.Frame(ctrl_frame, bg=ctrl_bg)
        row2.pack(fill="x")

        tk.Label(row2, text="Reference:", bg=ctrl_bg, fg=ctrl_fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 4))

        self._tuner_ref_notes = _build_ref_notes(440.0)
        note_names = [n for n, _ in self._tuner_ref_notes]
        self._tuner_ref_note_var = tk.StringVar(value="A4")
        ref_combo = ttk.Combobox(
            row2, textvariable=self._tuner_ref_note_var,
            values=note_names, state="readonly", width=5
        )
        ref_combo.pack(side="left", padx=(0, 8))

        self._tuner_waveform_var = tk.StringVar(
            value=tuner_settings.get("waveform", "pure"))
        tk.Radiobutton(
            row2, text="Pure", variable=self._tuner_waveform_var,
            value="pure", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 10)
        ).pack(side="left", padx=(0, 2))
        tk.Radiobutton(
            row2, text="Rich", variable=self._tuner_waveform_var,
            value="rich", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 10)
        ).pack(side="left", padx=(0, 10))

        self._tuner_play_btn = tk.Button(
            row2, text="▶ Play", command=self._tuner_toggle_ref_tone,
            font=("Helvetica", 10), width=8
        )
        self._tuner_play_btn.pack(side="left")

        # --- Initialize engine and player ---
        self._tuner_engine = TunerEngine()
        self._tuner_player = ReferencePlayer()

        # Apply saved settings
        try:
            self._tuner_engine.set_reference_pitch(float(self._tuner_pitch_var.get()))
        except ValueError:
            pass
        self._tuner_engine.set_sensitivity(self._tuner_sens_var.get())

        # Bind canvas resize to rebuild wheels
        self._tuner_canvas.bind("<Configure>", self._tuner_on_canvas_resize)
        self._tuner_wheels_built = False

    def _create_tuner_fallback(self, parent):
        """Show a message when audio libraries are not available."""
        bg = DARK_BG
        frame = tk.Frame(parent, bg=bg)
        frame.pack(fill="both", expand=True)
        frame._dark_canvas = True

        msg = ("Strobe Tuner requires numpy and sounddevice.\n\n"
               "Install them with:\n"
               "  pip install numpy sounddevice\n\n"
               "Then restart the application.")
        tk.Label(frame, text=msg, bg=bg, fg="#AAAAAA",
                 font=("Helvetica", 12), justify="center").pack(expand=True)

    # ------------------------------------------------------------------
    # WHEEL CREATION & LAYOUT
    # ------------------------------------------------------------------

    def _tuner_build_wheels(self):
        """Create the 12 strobe wheels in piano keyboard layout.

        Matches the Stroboconn: naturals (C,D,E,F,G,A,B) on the bottom row,
        accidentals (C#,D#,F#,G#,A#) on the top row, staggered like piano keys.
        Decorative strip at bottom with "Stohrer" script, motor pilot, FLAT/SHARP.
        """
        canvas = self._tuner_canvas
        canvas.delete("all")

        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 100 or h < 100:
            return

        # Reserve space for decorative strip at bottom
        deco_h = 40
        wheel_h = h - deco_h

        # Piano keyboard layout
        # Naturals: bottom row, 7 evenly spaced columns
        # Accidentals: top row, positioned between naturals (like black keys)
        naturals = [(0, 0), (2, 1), (4, 2), (5, 3), (7, 4), (9, 5), (11, 6)]
        accidentals = [(1, 0.5), (3, 1.5), (6, 3.5), (8, 4.5), (10, 5.5)]

        col_w = w / 7
        margin_x = col_w * 0.08  # Small left/right margin

        # Row layout: top row, gap, bottom row
        label_gap = 28  # vertical space between rows for labels
        row_h = (wheel_h - label_gap) / 2
        top_cy = row_h * 0.50
        bottom_cy = row_h + label_gap + row_h * 0.50

        # Wheel radius — fit to cell
        radius = min(col_w * 0.40, row_h * 0.44)

        # Build position map: pitch_class -> (cx, cy)
        positions = {}
        for pc, col in naturals:
            positions[pc] = (margin_x + col_w * (col + 0.5), bottom_cy)
        for pc, col in accidentals:
            positions[pc] = (margin_x + col_w * (col + 0.5), top_cy)

        # Pitch classes in the top row (accidentals) vs bottom row (naturals)
        top_pcs = {1, 3, 6, 8, 10}  # C#, D#, F#, G#, A#

        # Create wheels in pitch class order (0-11) so engine indices match
        label_offset = 6  # pixels past the wedge apex
        self._tuner_wheels = []
        for pc in range(12):
            cx, cy = positions[pc]
            direction = "up" if pc in top_pcs else "down"
            wheel = StrobeWheel(canvas, cx, cy, radius, self._tuner_color, direction)
            self._tuner_wheels.append(wheel)
            # Position label at wedge apex:
            #   Accidentals (top row, wedge up): apex at cy+radius, label just below
            #   Naturals (bottom row, wedge down): apex at cy-radius, label just above
            if pc in top_pcs:
                lbl_y = cy + radius + label_offset
            else:
                lbl_y = cy - radius - label_offset
            canvas.coords(wheel._label_id, cx, lbl_y)

        self._tuner_update_labels()

        # --- Decorative strip at bottom ---
        deco_top = wheel_h
        deco_cy = deco_top + deco_h * 0.5

        # Subtle top border
        canvas.create_line(10, deco_top, w - 10, deco_top, fill="#333333", width=1)

        # "Stohrer" in script font
        script_font = self._tuner_get_script_font(18)
        canvas.create_text(
            w * 0.14, deco_cy,
            text="Stohrer", fill="#CCCCCC",
            font=script_font, anchor="center"
        )

        # "FLAT  ←     →  SHARP"
        canvas.create_text(
            w * 0.46, deco_cy,
            text="FLAT  \u2190       \u2192  SHARP", fill="#777777",
            font=("Helvetica", 9), anchor="center"
        )

        # Motor pilot
        pilot_cx = w * 0.68
        pilot_cy = deco_cy
        pilot_r = 7
        glow_r = pilot_r + 4

        # Glow halo (behind pilot)
        self._tuner_pilot_glow = canvas.create_oval(
            pilot_cx - glow_r, pilot_cy - glow_r,
            pilot_cx + glow_r, pilot_cy + glow_r,
            fill="#1A0A00", outline="", width=0
        )
        # Pilot circle
        self._tuner_pilot_id = canvas.create_oval(
            pilot_cx - pilot_r, pilot_cy - pilot_r,
            pilot_cx + pilot_r, pilot_cy + pilot_r,
            fill="#331100", outline="#444444", width=1
        )
        # "MOTOR" / "PILOT" labels
        canvas.create_text(
            pilot_cx - pilot_r - 6, pilot_cy,
            text="MOTOR", fill="#555555",
            font=("Helvetica", 7), anchor="e"
        )
        canvas.create_text(
            pilot_cx + pilot_r + 6, pilot_cy,
            text="PILOT", fill="#555555",
            font=("Helvetica", 7), anchor="w"
        )

        # Set pilot state
        self._tuner_set_pilot(getattr(self, '_tuner_running', False))

        self._tuner_wheels_built = True

    def _tuner_get_script_font(self, size=18):
        """Find a script/cursive font with fallback."""
        import tkinter.font as tkfont
        families = set(tkfont.families())
        for name in ("Segoe Script", "Brush Script MT", "Lucida Handwriting",
                      "Snell Roundhand", "Apple Chancery"):
            if name in families:
                return (name, size, "bold")
        return ("Georgia", size, "bold italic")

    def _tuner_set_pilot(self, active):
        """Set motor pilot glow state (orange when engine is running)."""
        if not hasattr(self, '_tuner_pilot_id'):
            return
        canvas = self._tuner_canvas
        if active:
            canvas.itemconfigure(self._tuner_pilot_id, fill="#FF8800")
            canvas.itemconfigure(self._tuner_pilot_glow, fill="#442200")
        else:
            canvas.itemconfigure(self._tuner_pilot_id, fill="#331100")
            canvas.itemconfigure(self._tuner_pilot_glow, fill="#1A0A00")

    def _tuner_on_canvas_resize(self, event):
        """Rebuild wheels when canvas size changes."""
        if event.width > 100 and event.height > 100:
            self._tuner_build_wheels()

    # ------------------------------------------------------------------
    # LABEL MANAGEMENT (transposition)
    # ------------------------------------------------------------------

    def _tuner_update_labels(self):
        """Update wheel note labels based on transposition setting."""
        shift = TRANSPOSITION_SHIFTS.get(self._tuner_transpose_var.get(), 0)
        for i, wheel in enumerate(self._tuner_wheels):
            pc = (i + shift) % 12
            wheel.set_label(PITCH_CLASSES[pc])

    # ------------------------------------------------------------------
    # CONTROLS
    # ------------------------------------------------------------------

    def _tuner_on_pitch_changed(self):
        """Reference pitch spinbox changed."""
        try:
            hz = float(self._tuner_pitch_var.get())
            if self._tuner_engine:
                self._tuner_engine.set_reference_pitch(hz)
                # Rebuild reference note list
                self._tuner_ref_notes = _build_ref_notes(hz)
        except ValueError:
            pass

    def _tuner_on_sensitivity_changed(self, value=None):
        """Sensitivity slider changed."""
        if self._tuner_engine:
            self._tuner_engine.set_sensitivity(self._tuner_sens_var.get())

    def _tuner_open_settings(self):
        """Open tuner settings dialog (backlight color, FPS)."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Strobe Tuner Settings")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        # Backlight color
        color_row = tk.Frame(frame, bg=bg)
        color_row.pack(fill="x", pady=(0, 10))
        tk.Label(color_row, text="Backlight Color:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
        color_swatch = tk.Button(
            color_row, text="  ", bg=self._tuner_color, width=4,
            relief="raised", bd=1
        )
        color_swatch.pack(side="left")

        def pick_color():
            c = colorchooser.askcolor(
                initialcolor=self._tuner_color,
                title="Choose Backlight Color", parent=dlg
            )
            if c[1]:
                self._tuner_color = c[1]
                color_swatch.configure(bg=self._tuner_color)
                for wheel in self._tuner_wheels:
                    wheel.set_color(self._tuner_color)

        color_swatch.configure(command=pick_color)

        # FPS
        fps_row = tk.Frame(frame, bg=bg)
        fps_row.pack(fill="x", pady=(0, 10))
        tk.Label(fps_row, text="Frame Rate:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
        fps_combo = ttk.Combobox(
            fps_row, textvariable=self._tuner_fps_var,
            values=["60", "90", "120"], state="readonly", width=5
        )
        fps_combo.pack(side="left")
        tk.Label(fps_row, text="fps", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(4, 0))

        # Close button
        tk.Button(frame, text="Close", command=dlg.destroy,
                  width=10).pack(pady=(5, 0))

    def _tuner_toggle_ref_tone(self):
        """Toggle reference tone playback."""
        if self._tuner_player and self._tuner_player.is_playing:
            self._tuner_player.stop()
            self._tuner_play_btn.configure(text="▶ Play")
        else:
            # Find frequency for selected note
            note_name = self._tuner_ref_note_var.get()
            freq = 440.0
            for name, f in self._tuner_ref_notes:
                if name == note_name:
                    freq = f
                    break
            waveform = self._tuner_waveform_var.get()
            if self._tuner_player:
                self._tuner_player.play(freq, waveform)
                self._tuner_play_btn.configure(text="■ Stop")

    # ------------------------------------------------------------------
    # ANIMATION LOOP
    # ------------------------------------------------------------------

    def _tuner_start(self):
        """Start the tuner (audio capture + animation)."""
        if not self._tuner_engine:
            return

        if not self._tuner_wheels_built:
            self._tuner_build_wheels()

        success, err = self._tuner_engine.start()
        if not success:
            # Show error briefly on canvas
            if hasattr(self, '_tuner_canvas'):
                self._tuner_canvas.create_text(
                    self._tuner_canvas.winfo_width() / 2,
                    self._tuner_canvas.winfo_height() / 2,
                    text=f"Audio error: {err}",
                    fill="#FF4444", font=("Helvetica", 12),
                    tags="error"
                )
            return

        self._tuner_running = True
        self._tuner_set_pilot(True)
        self._tuner_animate()

    def _tuner_stop(self):
        """Stop the tuner (audio + animation)."""
        self._tuner_running = False
        self._tuner_set_pilot(False)
        if self._tuner_anim_id is not None:
            try:
                self.root.after_cancel(self._tuner_anim_id)
            except Exception:
                pass
            self._tuner_anim_id = None

        if self._tuner_engine:
            self._tuner_engine.stop()

        if self._tuner_player and self._tuner_player.is_playing:
            self._tuner_player.stop()
            if hasattr(self, '_tuner_play_btn'):
                self._tuner_play_btn.configure(text="▶ Play")

    def _tuner_animate(self):
        """One animation frame. Schedules itself at ~60fps."""
        if not self._tuner_running:
            return

        if self._tuner_engine and self._tuner_engine.is_running:
            result = self._tuner_engine.analyze()
            # Sensitivity acts as a gain control (like the Stroboconn's CONTROL knob)
            # Low sensitivity = need loud signal to light up wheels
            # High sensitivity = quiet signals still register
            sens = self._tuner_sens_var.get() / 100.0  # 0.0 to 1.0
            gain = 0.2 + sens * 4.8  # 0.2x to 5.0x
            for i, wheel in enumerate(self._tuner_wheels):
                mag = min(1.0, result.magnitudes[i] * gain)
                # Apply gain to per-ring magnitudes too
                ring_mags = [min(1.0, rm * gain) for rm in result.ring_magnitudes[i]]
                wheel.update(result.phase_offsets[i], mag, ring_mags)

        interval = FRAME_RATES.get(self._tuner_fps_var.get(), 16)
        self._tuner_anim_id = self.root.after(interval, self._tuner_animate)

    # ------------------------------------------------------------------
    # SETTINGS SAVE/RESTORE
    # ------------------------------------------------------------------

    def _tuner_save_settings(self):
        """Save tuner settings to the settings dict."""
        self.settings["tuner_settings"] = {
            "stripe_color": self._tuner_color if hasattr(self, '_tuner_color') else "#00FF00",
            "reference_pitch": float(self._tuner_pitch_var.get()) if hasattr(self, '_tuner_pitch_var') else 440.0,
            "transposition": self._tuner_transpose_var.get() if hasattr(self, '_tuner_transpose_var') else "C",
            "sensitivity": self._tuner_sens_var.get() if hasattr(self, '_tuner_sens_var') else 50,
            "waveform": self._tuner_waveform_var.get() if hasattr(self, '_tuner_waveform_var') else "pure",
            "fps": self._tuner_fps_var.get() if hasattr(self, '_tuner_fps_var') else "60",
        }
