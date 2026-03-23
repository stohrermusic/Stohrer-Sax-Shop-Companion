"""
Tuner tab mixin for Stohrer Sax Shop Companion.

12-wheel chromatic strobe tuner modeled after the Peterson Stroboconn 6T-5.
Concentric rings, wedge cutouts, phase tracking, with an analog VU meter
showing the detected fundamental pitch.

Requires: numpy, sounddevice (graceful fallback if unavailable)
"""

import tkinter as tk
from tkinter import ttk, colorchooser
import math
import functools
import sys

IS_MACOS = sys.platform == 'darwin'

try:
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

DEFAULT_FACEPLATE = "#1A1A1A"
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

# --- StrobeWheel rendering ---
CENTER_GAP_FRACTION = 0.12       # Fraction of radius reserved as center gap
RING_GAP_FRACTION = 0.05         # Fraction of ring width used as gap between rings
MASK_EXTEND_PX = 8               # Pixels the mask extends beyond disc edge
BRIGHTNESS_GAMMA = 0.6           # Gamma curve for magnitude → brightness mapping
MAGNITUDE_THRESHOLD = 0.05       # Below this magnitude, brightness snaps to zero
PHASE_CHANGE_THRESHOLD = 0.005   # Degrees — skip redraw below this (sub-pixel)
LABEL_BRIGHTNESS_MIN = 0.35      # Minimum label brightness when signal detected
LABEL_BRIGHTNESS_RANGE = 0.65    # Additional brightness scaled by magnitude

# --- VU meter behavior ---
VU_NEEDLE_DAMPING = 0.18         # Lerp factor per frame (0=frozen, 1=instant)
VU_SIGNAL_THRESHOLD = 0.08       # Minimum magnitude to register a pitch on VU
VU_IN_TUNE_CENTS = 1.0           # Cents threshold for "IN TUNE" display
VU_CENTS_RANGE = 50.0            # Max cents shown on VU scale (symmetric ±)
VU_ARC_START_DEG = 155.0         # Needle arc start angle (degrees, left side)
VU_ARC_END_DEG = 25.0            # Needle arc end angle (degrees, right side)
VU_RADIUS = 80                   # Radius of the VU meter arc

# --- Sensitivity gain mapping ---
GAIN_MIN = 0.2                   # Gain at sensitivity=0%
GAIN_RANGE = 4.8                 # Additional gain at sensitivity=100%

# --- Wheel layout ---
LAYOUT_MARGIN_FRACTION = 0.08    # Horizontal margin as fraction of column width
LAYOUT_LABEL_GAP = 18            # Vertical pixels between top and bottom wheel rows
LAYOUT_RADIUS_COL_LIMIT = 0.48   # Max radius as fraction of column width
LAYOUT_RADIUS_ROW_LIMIT = 0.55   # Max radius as fraction of row height

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

@functools.lru_cache(maxsize=256)
def _scale_color_cached(hex_color, factor_q):
    """Inner cached implementation — factor_q is pre-quantized."""
    hex_color = hex_color.lstrip('#')
    r = min(255, int(int(hex_color[0:2], 16) * factor_q))
    g = min(255, int(int(hex_color[2:4], 16) * factor_q))
    b = min(255, int(int(hex_color[4:6], 16) * factor_q))
    return f"#{r:02x}{g:02x}{b:02x}"


def _scale_color(hex_color, factor):
    """Return hex_color scaled by factor (0.0 = black, 1.0 = full brightness).

    Factor is quantized to 2 decimal places before cache lookup — this maps
    ~84 per-frame calls down to a handful of unique cache keys without
    visible quality loss (1/255 ≈ 0.004, so 0.01 resolution is plenty).
    """
    return _scale_color_cached(hex_color, round(factor, 2))


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

    def __init__(self, canvas, cx, cy, radius, stripe_color, faceplate_color,
                 direction="up"):
        """Create a strobe wheel.

        Args:
            direction: "up" = wedge opens upward (top row), "down" = opens downward (bottom row)
            faceplate_color: Background color for mask areas
        """
        self.canvas = canvas
        self.cx = cx
        self.cy = cy
        self.radius = radius
        self.stripe_color = stripe_color
        self.faceplate_color = faceplate_color
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
        gap = radius * CENTER_GAP_FRACTION  # First ring starts here
        ring_width = (radius - gap) / NUM_RINGS
        self._ring_radii = []
        for i in range(NUM_RINGS):
            r_inner = gap + i * ring_width
            r_outer = gap + (i + 1) * ring_width - ring_width * RING_GAP_FRACTION
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
        """Create wedge-shaped mask matching the Stroboconn cutout."""
        cx, cy = self.cx, self.cy
        r = self.radius
        mr = r + MASK_EXTEND_PX  # Mask extends slightly beyond disc edge

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
            points, fill=self.faceplate_color, outline='', width=0
        )

        # Inner mask: covers the center gap (no rings visible there)
        inner_r = self._ring_radii[0][0]  # Inner radius of first ring
        inner_pts = []
        for i in range(32):
            a = i * 2 * math.pi / 32
            inner_pts.append(cx + inner_r * math.cos(a))
            inner_pts.append(cy - inner_r * math.sin(a))
        self._inner_mask = self.canvas.create_polygon(
            inner_pts, fill=self.faceplate_color, outline='', width=0
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

    def update(self, phase_offset, magnitude, ring_magnitudes=None,
               ring_brightness_pct=100, overall_brightness_pct=80):
        """Update wheel for one animation frame.

        Args:
            phase_offset: Rotation angle in degrees (from engine phase tracking)
            magnitude: Overall signal strength 0.0-1.0 (drives label brightness)
            ring_magnitudes: Optional list of per-ring magnitudes (0.0-1.0).
            ring_brightness_pct: 0-100, how much per-ring effect to apply
                (0=uniform brightness, 100=full per-ring independent brightness)
            overall_brightness_pct: 0-100, scales the overall brightness ceiling
        """
        overall_scale = overall_brightness_pct / 100.0

        # Compute overall brightness for label
        brightness = min(1.0, magnitude ** BRIGHTNESS_GAMMA) if magnitude > MAGNITUDE_THRESHOLD else 0.0

        # Per-ring brightness from real spectral data
        ring_mix = ring_brightness_pct / 100.0  # 0.0 = uniform, 1.0 = full per-ring
        if ring_magnitudes and ring_mix > 0.0:
            ring_fills = []
            uniform_factor = DIM_MULTIPLIER + brightness * (1.0 - DIM_MULTIPLIER)
            for ring_idx in range(NUM_RINGS):
                rm = min(1.0, ring_magnitudes[ring_idx])
                rb = min(1.0, rm ** BRIGHTNESS_GAMMA) if rm > MAGNITUDE_THRESHOLD else 0.0
                per_ring_factor = DIM_MULTIPLIER + rb * (1.0 - DIM_MULTIPLIER)
                # Blend between uniform and per-ring based on ring_mix
                blended = uniform_factor * (1.0 - ring_mix) + per_ring_factor * ring_mix
                ring_fills.append(_scale_color(self.stripe_color, blended * overall_scale))
        else:
            fill_factor = DIM_MULTIPLIER + brightness * (1.0 - DIM_MULTIPLIER)
            fill_factor *= overall_scale
            ring_fills = [_scale_color(self.stripe_color, fill_factor)] * NUM_RINGS

        # Update fill colors per ring (only if changed)
        if ring_fills != self._last_ring_fills:
            self._last_ring_fills = ring_fills
            for ring_idx, ring_segs in enumerate(self._segments):
                fill = ring_fills[ring_idx]
                for poly_id, _, _, _ in ring_segs:
                    self.canvas.itemconfigure(poly_id, fill=fill)
            # Label brightness tracks overall magnitude
            label_color = _scale_color("#FFFFFF", LABEL_BRIGHTNESS_MIN + brightness * LABEL_BRIGHTNESS_RANGE)
            self.canvas.itemconfigure(self._label_id, fill=label_color)

        self._brightness = brightness

        # Update segment positions (rotate by phase_offset)
        # Skip coordinate recalculation when phase hasn't moved enough to see.
        # 0.005 degrees ≈ 0.09px at 50px radius — sub-pixel, invisible.
        if abs(phase_offset - self._phase_offset) < PHASE_CHANGE_THRESHOLD:
            return  # No visible change, skip coordinate update

        self._phase_offset = phase_offset

        # Only update segments that are visible through the wedge cutout.
        # The wedge shows WEDGE_ANGLE degrees centered on 90° (top).
        # Segments outside this arc are hidden behind the mask — skip them.
        wedge_half = WEDGE_ANGLE / 2.0
        # Visible range in "math angle" space (90° = top of screen)
        vis_lo = 90.0 - wedge_half
        vis_hi = 90.0 + wedge_half

        for ring_idx in range(NUM_RINGS):
            r_inner, r_outer = self._ring_radii[ring_idx]
            for poly_id, base_start, seg_span, steps in self._segments[ring_idx]:
                start = (base_start + phase_offset) % 360.0
                end = start + seg_span
                # Check if segment overlaps the visible wedge
                # (handle wrap-around: segment at 350°-10° overlaps wedge at 50°-130°)
                if not (end > vis_lo and start < vis_hi):
                    # Also check wrap-around
                    if not (end + 360.0 > vis_lo and start < vis_hi) and \
                       not (end > vis_lo and start - 360.0 < vis_hi):
                        continue  # Hidden behind mask — skip

                points = _annular_sector_points(
                    self.cx, self.cy, r_inner, r_outer,
                    base_start + phase_offset,
                    base_start + phase_offset + seg_span, steps
                )
                self.canvas.coords(poly_id, *points)


# ============================================
# TUNER TAB MIXIN
# ============================================

class TunerTabMixin:
    """Mixin class that adds the Tuner tab to the main application."""

    def _init_tuner_state(self):
        """Initialize tuner state. Called from __init__."""
        self._tuner_engine = None
        self._tuner_player = None
        self._tuner_wheels = []
        self._tuner_running = False
        self._tuner_anim_id = None
        self._tuner_fps_times = []      # Timestamps for FPS measurement
        self._tuner_fps_display = None  # Canvas text id for FPS overlay
        self._tuner_perf_log = []       # Collected perf samples for debug dump
        self._tuner_perf_frame = 0      # Frame counter for perf logging
        self._tuner_perf_text = ""      # Latest perf text for FPS overlay

    def create_tuner_tab(self, parent):
        """Build the Tuner tab UI."""
        if not _TUNER_IMPORTS_OK or not AUDIO_AVAILABLE:
            self._create_tuner_fallback(parent)
            return

        tuner_settings = self.settings.get("tuner_settings", {})

        # Load settings
        self._tuner_color = tuner_settings.get("stripe_color", "#00FF00")
        self._tuner_faceplate_color = tuner_settings.get("faceplate_color", DEFAULT_FACEPLATE)
        self._tuner_ring_brightness = tuner_settings.get("ring_brightness", 100)
        self._tuner_overall_brightness = tuner_settings.get("overall_brightness", 80)
        self._tuner_fps_var = tk.StringVar(
            value=str(tuner_settings.get("fps", "60")))
        self._tuner_show_fps = tk.BooleanVar(
            value=tuner_settings.get("show_fps", False))

        bg = self._tuner_faceplate_color

        # --- Main container (skip theme walker — dark display) ---
        self._tuner_main_frame = tk.Frame(parent, bg=bg)
        self._tuner_main_frame._skip_theme = True
        self._tuner_main_frame.pack(fill="both", expand=True)

        # --- Canvas (strobe wheels) ---
        self._tuner_canvas = tk.Canvas(
            self._tuner_main_frame, bg=bg, highlightthickness=0,
            borderwidth=0,
        )
        self._tuner_canvas._dark_canvas = True  # Skip theme walker
        self._tuner_canvas.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        # --- Control panel (EQ sliders | flat/pilot/sharp | VU meter) ---
        ctrl_bg = "systemWindowBackgroundColor" if IS_MACOS else "#2A2A2A"
        ctrl_fg = "white" if not IS_MACOS else "systemTextColor"
        self._tuner_ctrl_bg = ctrl_bg
        self._tuner_ctrl_fg = ctrl_fg

        ctrl_frame = tk.Frame(self._tuner_main_frame, bg=ctrl_bg, padx=6, pady=6)
        ctrl_frame._skip_theme = True
        ctrl_frame.pack(fill="x", padx=5, pady=(0, 4))
        ctrl_frame.columnconfigure(0, weight=1)
        ctrl_frame.columnconfigure(1, weight=1)
        ctrl_frame.columnconfigure(2, weight=1)

        eq_lbl_font = ("Helvetica", 7)

        # ---- LEFT: Graphic EQ vertical sliders ----
        eq_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        eq_frame._skip_theme = True
        eq_frame.grid(row=0, column=0, sticky="ns", padx=(0, 8))

        def _make_vslider(parent, label, var, lo, hi, col, cmd=None,
                          resolution=1, value_fmt=None):
            ch = tk.Frame(parent, bg=ctrl_bg)
            ch._skip_theme = True
            ch.grid(row=0, column=col, padx=3, sticky="ns")
            tk.Label(ch, text=label, bg=ctrl_bg, fg="#888888",
                     font=eq_lbl_font).pack(pady=(0, 2))
            tk.Scale(ch, variable=var, from_=hi, to=lo,
                     orient="vertical", length=90, width=12,
                     showvalue=False, resolution=resolution,
                     bg="#B0B0B0", fg=ctrl_fg, activebackground="#D0D0D0",
                     troughcolor="#444444", highlightthickness=0,
                     sliderrelief="raised", sliderlength=18,
                     borderwidth=2,
                     command=cmd).pack()
            if value_fmt:
                val_lbl = tk.Label(ch, text="", bg=ctrl_bg, fg="#AAAAAA",
                                   font=eq_lbl_font)
                val_lbl.pack(pady=(2, 0))
                def _upd(*_, _l=val_lbl, _f=value_fmt, _v=var):
                    _l.configure(text=_f(_v.get()))
                var.trace_add("write", _upd)
                _upd()

        self._tuner_sens_var = tk.IntVar(
            value=tuner_settings.get("sensitivity", 50))
        _make_vslider(eq_frame, "SENS", self._tuner_sens_var, 0, 100, 0,
                      cmd=self._tuner_on_sensitivity_changed)

        self._tuner_bias_var = tk.IntVar(value=self._tuner_ring_brightness)
        _make_vslider(eq_frame, "BIAS", self._tuner_bias_var, 0, 100, 1,
                      cmd=lambda v: setattr(self, '_tuner_ring_brightness', int(v)))

        self._tuner_bright_var = tk.IntVar(value=self._tuner_overall_brightness)
        _make_vslider(eq_frame, "BRIGHT", self._tuner_bright_var, 10, 150, 2,
                      cmd=lambda v: setattr(self, '_tuner_overall_brightness', int(v)))

        # A= pitch slider (420–460 Hz)
        self._tuner_pitch_var = tk.DoubleVar(
            value=tuner_settings.get("reference_pitch", 440.0))
        _make_vslider(eq_frame, "A =", self._tuner_pitch_var, 420, 460, 3,
                      cmd=lambda v: self._tuner_on_pitch_changed(),
                      resolution=1, value_fmt=lambda v: f"{v:.0f} Hz")

        # Key slider (C / Bb / Eb / F mapped to 0–3)
        _trans_migrate = {"concert": "C", "bb": "Bb", "eb": "Eb", "f": "F"}
        saved_trans = tuner_settings.get("transposition", "C")
        saved_trans = _trans_migrate.get(saved_trans, saved_trans)
        self._tuner_transpose_var = tk.StringVar(value=saved_trans)

        saved_key_idx = TRANSPOSITION_KEYS.index(saved_trans) if saved_trans in TRANSPOSITION_KEYS else 0
        self._tuner_key_idx_var = tk.IntVar(value=saved_key_idx)

        def _on_key_slider(v):
            key = TRANSPOSITION_KEYS[int(float(v))]
            self._tuner_transpose_var.set(key)
            self._tuner_update_labels()

        _make_vslider(eq_frame, "KEY", self._tuner_key_idx_var, 0, 3, 4,
                      cmd=_on_key_slider,
                      value_fmt=lambda v: TRANSPOSITION_KEYS[int(v)])

        # FPS 3-position switch (60/90/120)
        fps_values = ["60", "90", "120"]
        fps_idx = fps_values.index(self._tuner_fps_var.get()) if self._tuner_fps_var.get() in fps_values else 0
        fps_idx_var = tk.IntVar(value=fps_idx)

        fps_ch = tk.Frame(eq_frame, bg=ctrl_bg)
        fps_ch._skip_theme = True
        fps_ch.grid(row=0, column=5, padx=3, sticky="ns")
        tk.Label(fps_ch, text="FPS", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(pady=(0, 2))
        tk.Scale(fps_ch, variable=fps_idx_var, from_=0, to=2,
                 orient="vertical", length=90, width=12,
                 showvalue=False, resolution=1,
                 bg="#B0B0B0", fg=ctrl_fg, activebackground="#D0D0D0",
                 troughcolor="#444444", highlightthickness=0,
                 sliderrelief="raised", sliderlength=18,
                 borderwidth=2).pack()
        fps_lbl = tk.Label(fps_ch, text=fps_values[fps_idx], bg=ctrl_bg,
                           fg="#AAAAAA", font=eq_lbl_font)
        fps_lbl.pack(pady=(2, 0))

        def _on_fps_slider(*args):
            idx = fps_idx_var.get()
            self._tuner_fps_var.set(fps_values[idx])
            fps_lbl.configure(text=fps_values[idx])
        fps_idx_var.trace_add("write", _on_fps_slider)

        # ---- CENTER: Experimental / flat-pilot-sharp / ref button ----
        self._tuner_ref_notes = _build_ref_notes(440.0)
        self._tuner_ref_note_var = tk.StringVar(value="A4")
        self._tuner_waveform_var = tk.StringVar(
            value=tuner_settings.get("waveform", "pure"))

        center_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        center_frame._skip_theme = True
        center_frame.grid(row=0, column=1, padx=10, sticky="ns")

        # Use grid inside center_frame so everything shares a single center axis
        center_frame.columnconfigure(0, weight=1)
        row = 0

        script_font = self._tuner_get_script_font(11)
        tk.Label(center_frame, text="Experimental", bg=ctrl_bg, fg="#666666",
                 font=script_font).grid(row=row, column=0, pady=(0, 4))
        row += 1

        indicator_frame = tk.Frame(center_frame, bg=ctrl_bg)
        indicator_frame._skip_theme = True
        indicator_frame.grid(row=row, column=0)
        row += 1

        tk.Label(indicator_frame, text=" \u2190 flat", bg=ctrl_bg, fg="#888888",
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 6))

        pilot_cv = tk.Canvas(indicator_frame, width=30, height=30,
                             bg=ctrl_bg, highlightthickness=0, bd=0)
        pilot_cv._skip_theme = True
        pilot_cv.pack(side="left")
        self._tuner_pilot_canvas = pilot_cv
        pcx, pcy, pr = 15, 15, 8
        self._tuner_pilot_glow = pilot_cv.create_oval(
            pcx - pr - 4, pcy - pr - 4, pcx + pr + 4, pcy + pr + 4,
            fill="#1A0A00", outline="", width=0)
        self._tuner_pilot_id = pilot_cv.create_oval(
            pcx - pr, pcy - pr, pcx + pr, pcy + pr,
            fill="#331100", outline="#444444", width=1)

        tk.Label(indicator_frame, text="sharp \u2192", bg=ctrl_bg, fg="#888888",
                 font=("Helvetica", 10)).pack(side="left", padx=(6, 0))

        tk.Label(center_frame, text="MOTOR PILOT", bg=ctrl_bg, fg="#555555",
                 font=("Helvetica", 7)).grid(row=row, column=0)
        row += 1

        # ---- RIGHT: VU meter (centered in its column) ----
        vu_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        vu_frame._skip_theme = True
        vu_frame.grid(row=0, column=2, sticky="ns", padx=8)

        self._vu_canvas = tk.Canvas(
            vu_frame, width=200, height=120,
            bg=ctrl_bg, highlightthickness=0, bd=0)
        self._vu_canvas._skip_theme = True
        self._vu_canvas.pack()

        # Note/cents readout below meter
        readout = tk.Frame(vu_frame, bg=ctrl_bg)
        readout._skip_theme = True
        readout.pack(pady=(2, 0))
        self._vu_note_label = tk.Label(
            readout, text="", bg=ctrl_bg, fg="#AAAAAA",
            font=("Helvetica", 11, "bold"), width=4, anchor="e")
        self._vu_note_label.pack(side="left", padx=(0, 4))
        self._vu_cents_label = tk.Label(
            readout, text="", bg=ctrl_bg, fg="#AAAAAA",
            font=("Helvetica", 10), width=8, anchor="w")
        self._vu_cents_label.pack(side="left")

        self._vu_smooth_cents = 0.0  # for needle damping
        self._tuner_build_vu()

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
        bg = DEFAULT_FACEPLATE
        frame = tk.Frame(parent, bg=bg)
        frame.pack(fill="both", expand=True)
        frame._dark_canvas = True

        import sys
        if sys.platform == 'linux':
            msg = ("The Strobe Tuner requires PortAudio.\n\n"
                   "Install it with:\n"
                   "  sudo apt install libportaudio2\n\n"
                   "Then restart the application.")
        elif sys.platform == 'darwin':
            msg = ("The Strobe Tuner is not available on this Mac.\n\n"
                   "This feature requires Apple Silicon (M1 or newer).\n\n"
                   "All other features work normally.")
        else:
            msg = ("The Strobe Tuner is not available on this system.\n\n"
                   "If you see this on Windows, try reinstalling the app.\n\n"
                   "All other features work normally.")
        tk.Label(frame, text=msg, bg=bg, fg="#AAAAAA",
                 font=("Helvetica", 12), justify="center").pack(expand=True)

    # ------------------------------------------------------------------
    # WHEEL CREATION & LAYOUT
    # ------------------------------------------------------------------

    def _tuner_build_wheels(self):
        """Create the 12 strobe wheels in piano keyboard layout."""
        canvas = self._tuner_canvas
        canvas.delete("all")
        self._tuner_fps_display = None  # Canvas items destroyed by delete("all")

        bg = self._tuner_faceplate_color
        canvas.configure(bg=bg)

        w = canvas.winfo_width()
        h = canvas.winfo_height()
        if w < 100 or h < 100:
            return

        wheel_h = h

        # Piano keyboard layout
        naturals = [(0, 0), (2, 1), (4, 2), (5, 3), (7, 4), (9, 5), (11, 6)]
        accidentals = [(1, 0.5), (3, 1.5), (6, 3.5), (8, 4.5), (10, 5.5)]

        col_w = w / 7
        margin_x = col_w * LAYOUT_MARGIN_FRACTION

        # Row layout: top row overlapping bottom row (like real Stroboconn)
        label_gap = LAYOUT_LABEL_GAP
        row_h = (wheel_h - label_gap) / 2
        top_cy = row_h * 0.50
        bottom_cy = row_h + label_gap + row_h * 0.50

        radius = min(col_w * LAYOUT_RADIUS_COL_LIMIT, row_h * LAYOUT_RADIUS_ROW_LIMIT)

        positions = {}
        for pc, col in naturals:
            positions[pc] = (margin_x + col_w * (col + 0.5), bottom_cy)
        for pc, col in accidentals:
            positions[pc] = (margin_x + col_w * (col + 0.5), top_cy)

        top_pcs = {1, 3, 6, 8, 10}

        label_offset = 6
        self._tuner_wheels = []
        for pc in range(12):
            cx, cy = positions[pc]
            direction = "up" if pc in top_pcs else "down"
            wheel = StrobeWheel(canvas, cx, cy, radius, self._tuner_color,
                                self._tuner_faceplate_color, direction)
            self._tuner_wheels.append(wheel)
            if pc in top_pcs:
                lbl_y = cy + radius + label_offset
            else:
                lbl_y = cy - radius - label_offset
            canvas.coords(wheel._label_id, cx, lbl_y)

        self._tuner_update_labels()
        self._tuner_wheels_built = True

    # ------------------------------------------------------------------
    # SHARED HELPERS
    # ------------------------------------------------------------------

    def _tuner_get_script_font(self, size=18):
        """Find a script/cursive font with fallback."""
        import tkinter.font as tkfont
        families = set(tkfont.families())
        for name in ("Segoe Script", "Brush Script MT", "Lucida Handwriting",
                      "Snell Roundhand", "Apple Chancery"):
            if name in families:
                return (name, size, "bold")
        return ("Georgia", size, "bold italic")

    def _tuner_build_vu(self):
        """Draw a vintage backlit VU meter on self._vu_canvas."""
        cv = self._vu_canvas
        cv.delete("all")

        cv_w, cv_h = 200, 120
        vu_cx = cv_w // 2
        vu_cy = cv_h - 10
        vu_r = VU_RADIUS
        arc_start = VU_ARC_START_DEG
        arc_end = VU_ARC_END_DEG

        # --- Backlit amber panel ---
        amber = "#D4920A"
        border_dark = "#3A3A3A"
        border_mid = "#555555"

        # Outer bezel
        cv.create_rectangle(2, 2, cv_w - 2, cv_h - 2,
                            fill=border_dark, outline=border_mid, width=1)
        # Inner glowing panel
        cv.create_rectangle(6, 6, cv_w - 6, cv_h - 6,
                            fill=amber, outline="#B87D08", width=1)

        # --- Tick marks (dark on amber) ---
        tick_color = "#2A1A00"
        for i in range(21):
            cents = -50 + i * 5
            frac = (cents + 50) / 100.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)

            if cents == 0:
                tick_len, tick_w = 14, 2
            elif abs(cents) % 10 == 0:
                tick_len, tick_w = 10, 1
            else:
                tick_len, tick_w = 5, 1

            x_o = vu_cx + vu_r * math.cos(angle_rad)
            y_o = vu_cy - vu_r * math.sin(angle_rad)
            x_i = vu_cx + (vu_r - tick_len) * math.cos(angle_rad)
            y_i = vu_cy - (vu_r - tick_len) * math.sin(angle_rad)
            cv.create_line(x_i, y_i, x_o, y_o, fill=tick_color, width=tick_w)

        # --- Scale arc line ---
        arc_points = []
        for i in range(51):
            frac = i / 50.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            arc_points.extend([
                vu_cx + vu_r * math.cos(angle_rad),
                vu_cy - vu_r * math.sin(angle_rad),
            ])
        cv.create_line(*arc_points, fill=tick_color, width=1, smooth=True)

        # --- Scale labels: -50 ... 0 ... +50 cents ---
        label_r = vu_r + 11
        label_font = ("Helvetica", 7)
        for cents, label in [
            (-50, "50"), (-30, "30"), (-10, "10"),
            (0, "0"), (10, "10"), (30, "30"), (50, "50"),
        ]:
            frac = (cents + 50) / 100.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = vu_cx + label_r * math.cos(angle_rad)
            ly = vu_cy - label_r * math.sin(angle_rad)
            color = "#1B5E00" if cents == 0 else tick_color
            cv.create_text(lx, ly, text=label, fill=color,
                           font=label_font, anchor="center")

        # Flat/sharp indicators at extremes
        for cents, label in [(-50, "\u2013"), (50, "+")]:
            frac = (cents + 50) / 100.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = vu_cx + (label_r + 10) * math.cos(angle_rad)
            ly = vu_cy - (label_r + 10) * math.sin(angle_rad)
            cv.create_text(lx, ly, text=label, fill=tick_color,
                           font=("Helvetica", 9, "bold"), anchor="center")

        # --- Needle (heavier, with shadow) ---
        needle_len = vu_r - 16
        center_angle = math.radians((arc_start + arc_end) / 2)
        nx = vu_cx + needle_len * math.cos(center_angle)
        ny = vu_cy - needle_len * math.sin(center_angle)
        # Shadow
        self._vu_shadow_id = cv.create_line(
            vu_cx + 1, vu_cy + 1, nx + 1, ny + 1,
            fill="#8A6500", width=3, capstyle="round")
        self._vu_needle_id = cv.create_line(
            vu_cx, vu_cy, nx, ny, fill="#1A1200", width=3, capstyle="round")

        # Pivot (dark circle)
        cv.create_oval(vu_cx - 5, vu_cy - 5, vu_cx + 5, vu_cy + 5,
                       fill="#2A1A00", outline="#1A1200", width=1)

        # Store geometry for animation
        self._vu_cx = vu_cx
        self._vu_cy = vu_cy
        self._vu_needle_len = needle_len
        self._vu_arc_start = arc_start
        self._vu_arc_end = arc_end

    def _tuner_set_pilot(self, active):
        """Set motor pilot glow state (orange when engine is running)."""
        if not hasattr(self, '_tuner_pilot_id'):
            return
        canvas = self._tuner_pilot_canvas
        if active:
            canvas.itemconfigure(self._tuner_pilot_id, fill="#FF8800")
            canvas.itemconfigure(self._tuner_pilot_glow, fill="#442200")
        else:
            canvas.itemconfigure(self._tuner_pilot_id, fill="#331100")
            canvas.itemconfigure(self._tuner_pilot_glow, fill="#1A0A00")

    def _vu_update(self, result):
        """Update the analog VU meter with detected pitch."""
        if not hasattr(self, '_vu_needle_id'):
            return

        canvas = self._vu_canvas
        shift = TRANSPOSITION_SHIFTS.get(self._tuner_transpose_var.get(), 0)

        # Find dominant pitch class
        best_pc = -1
        best_mag = 0.0
        for pc in range(12):
            if result.magnitudes[pc] > best_mag:
                best_mag = result.magnitudes[pc]
                best_pc = pc

        target_cents = 0.0
        if best_pc < 0 or best_mag < VU_SIGNAL_THRESHOLD:
            # No signal — park needle at center, clear readout
            self._vu_note_label.configure(text="")
            self._vu_cents_label.configure(text="", fg="#AAAAAA")
        else:
            target_cents = result.cents_errors[best_pc]
            display_pc = (best_pc + shift) % 12
            note_name = PITCH_CLASSES[display_pc]

            # Find dominant octave
            best_oct = 4
            best_ring_mag = 0.0
            for ring_idx in range(NUM_RINGS):
                rm = result.ring_magnitudes[best_pc][ring_idx]
                if rm > best_ring_mag:
                    best_ring_mag = rm
                    best_oct = MIN_OCTAVE + ring_idx

            # Adjust octave when transposition wraps past C
            if best_pc + shift >= 12:
                best_oct += 1

            self._vu_note_label.configure(text=f"{note_name}{best_oct}")

            if abs(target_cents) < VU_IN_TUNE_CENTS:
                self._vu_cents_label.configure(text="IN TUNE", fg="#00CC00")
            elif target_cents > 0:
                self._vu_cents_label.configure(
                    text=f"+{target_cents:.0f}\u00a2", fg="#AAAAAA")
            else:
                self._vu_cents_label.configure(
                    text=f"{target_cents:.0f}\u00a2", fg="#AAAAAA")

        # Damped needle — lerp toward target (weighted, not instant)
        self._vu_smooth_cents += (target_cents - self._vu_smooth_cents) * VU_NEEDLE_DAMPING

        clamped = max(-VU_CENTS_RANGE, min(VU_CENTS_RANGE, self._vu_smooth_cents))
        frac = (clamped + VU_CENTS_RANGE) / (2.0 * VU_CENTS_RANGE)
        angle_deg = self._vu_arc_start + (self._vu_arc_end - self._vu_arc_start) * frac
        angle_rad = math.radians(angle_deg)
        nx = self._vu_cx + self._vu_needle_len * math.cos(angle_rad)
        ny = self._vu_cy - self._vu_needle_len * math.sin(angle_rad)
        canvas.coords(self._vu_shadow_id,
                      self._vu_cx + 1, self._vu_cy + 1, nx + 1, ny + 1)
        canvas.coords(self._vu_needle_id, self._vu_cx, self._vu_cy, nx, ny)

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
                self._tuner_ref_notes = _build_ref_notes(hz)
        except ValueError:
            pass

    def _tuner_on_sensitivity_changed(self, value=None):
        """Sensitivity slider changed."""
        if self._tuner_engine:
            self._tuner_engine.set_sensitivity(self._tuner_sens_var.get())

    def _tuner_open_settings(self):
        """Open tuner settings dialog."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Tuner Settings")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        # --- Input device ---
        from config import get_input_devices
        devices = get_input_devices()
        if devices:
            mic_row = tk.Frame(frame, bg=bg)
            mic_row.pack(fill="x", pady=(0, 10))
            tk.Label(mic_row, text="Input Device:", bg=bg, fg=fg,
                     font=("Helvetica", 10)).pack(side="left", padx=(0, 8))

            current_dev = self.settings.get("audio_input_device")
            dev_names = ["System Default"] + [name for _, name in devices]
            dev_indices = [None] + [idx for idx, _ in devices]

            mic_var = tk.StringVar(value="System Default")
            if current_dev is not None:
                for idx, name in devices:
                    if idx == current_dev:
                        mic_var.set(name)
                        break

            mic_combo = ttk.Combobox(mic_row, textvariable=mic_var,
                                     values=dev_names, state="readonly", width=35)
            mic_combo.pack(side="left")

            def on_mic_changed(event=None):
                sel = mic_combo.current()
                dev_idx = dev_indices[sel] if sel >= 0 else None
                self.settings["audio_input_device"] = dev_idx
                # Restart engine with new device
                if self._tuner_engine and self._tuner_engine.is_running:
                    self._tuner_stop()
                    self._tuner_engine.start(device=dev_idx)
                    self._tuner_running = True
                    self._tuner_set_pilot(True)
                    self._tuner_animate()

            mic_combo.bind("<<ComboboxSelected>>", on_mic_changed)

        # --- Stripe/Backlight color ---
        color_row = tk.Frame(frame, bg=bg)
        color_row.pack(fill="x", pady=(0, 10))
        tk.Label(color_row, text="Stripe Color:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
        color_swatch = tk.Button(
            color_row, text="  ", bg=self._tuner_color, width=4,
            relief="raised", bd=1
        )
        color_swatch.pack(side="left")

        def pick_stripe_color():
            c = colorchooser.askcolor(
                initialcolor=self._tuner_color,
                title="Choose Stripe Color", parent=dlg
            )
            if c[1]:
                self._tuner_color = c[1]
                color_swatch.configure(bg=self._tuner_color)
                for wheel in self._tuner_wheels:
                    wheel.set_color(self._tuner_color)

        color_swatch.configure(command=pick_stripe_color)

        # --- Faceplate color ---
        fp_row = tk.Frame(frame, bg=bg)
        fp_row.pack(fill="x", pady=(0, 10))
        tk.Label(fp_row, text="Faceplate Color:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
        fp_swatch = tk.Button(
            fp_row, text="  ", bg=self._tuner_faceplate_color, width=4,
            relief="raised", bd=1
        )
        fp_swatch.pack(side="left")

        def pick_faceplate_color():
            c = colorchooser.askcolor(
                initialcolor=self._tuner_faceplate_color,
                title="Choose Faceplate Color", parent=dlg
            )
            if c[1]:
                self._tuner_faceplate_color = c[1]
                fp_swatch.configure(bg=self._tuner_faceplate_color)
                # Rebuild wheels to apply new faceplate color
                self._tuner_wheels_built = False
                self._tuner_build_wheels()

        fp_swatch.configure(command=pick_faceplate_color)

        # --- Show FPS ---
        tk.Checkbutton(
            frame, text="Show frame rate on screen",
            variable=self._tuner_show_fps,
            bg=bg, fg=fg, selectcolor=bg, activebackground=bg,
            font=("Helvetica", 10),
        ).pack(fill="x", pady=(0, 10))

        # Close button
        tk.Button(frame, text="Close", command=dlg.destroy,
                  width=10).pack(pady=(5, 0))

    # ------------------------------------------------------------------
    # ANIMATION LOOP
    # ------------------------------------------------------------------

    def _tuner_start(self):
        """Start the tuner (audio capture + animation)."""
        if not self._tuner_engine:
            return

        if hasattr(self, '_tuner_canvas'):
            self._tuner_canvas.delete("error")

        if not self._tuner_wheels_built:
            self._tuner_build_wheels()

        device = self.settings.get("audio_input_device")
        success, err = self._tuner_engine.start(device=device)
        if not success:
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
                self._tuner_play_btn.configure(text="\u25b6")

    def _tuner_animate(self):
        """One animation frame — update strobe wheels and VU meter."""
        if not self._tuner_running:
            return

        # Check if the engine died with an error
        if self._tuner_engine and self._tuner_engine.last_error:
            self._tuner_show_stream_error(self._tuner_engine.last_error)
            return

        if self._tuner_engine and self._tuner_engine.is_running:
            import time as _time
            _t0 = _time.perf_counter()

            result = self._tuner_engine.analyze()

            _t1 = _time.perf_counter()

            # Re-check after analyze() — stream may have just died
            if self._tuner_engine.last_error:
                self._tuner_show_stream_error(self._tuner_engine.last_error)
                return

            wheels_updated = 0
            coords_calls = 0
            color_calls = 0
            phase_skips = 0

            if self._tuner_wheels:
                sens = self._tuner_sens_var.get() / 100.0
                gain = GAIN_MIN + sens * GAIN_RANGE
                for i, wheel in enumerate(self._tuner_wheels):
                    mag = min(1.0, result.magnitudes[i] * gain)
                    ring_mags = [min(1.0, rm * gain)
                                 for rm in result.ring_magnitudes[i]]

                    old_phase = wheel._phase_offset
                    old_fills = wheel._last_ring_fills
                    wheel.update(
                        result.phase_offsets[i],
                        mag,
                        ring_magnitudes=ring_mags,
                        ring_brightness_pct=self._tuner_ring_brightness,
                        overall_brightness_pct=self._tuner_overall_brightness,
                    )
                    if wheel._phase_offset != old_phase:
                        wheels_updated += 1
                        # Count visible segments that were updated
                        for ring in wheel._segments:
                            coords_calls += len(ring)
                    else:
                        phase_skips += 1
                    if wheel._last_ring_fills != old_fills:
                        for ring in wheel._segments:
                            color_calls += len(ring)

            _t2 = _time.perf_counter()

            self._vu_update(result)

            _t3 = _time.perf_counter()

            # Perf logging
            if self._tuner_show_fps.get():
                analyze_ms = (_t1 - _t0) * 1000
                wheel_ms = (_t2 - _t1) * 1000
                vu_ms = (_t3 - _t2) * 1000
                total_ms = (_t3 - _t0) * 1000

                self._tuner_perf_text = (
                    f"analyze:{analyze_ms:.0f}ms "
                    f"wheels:{wheel_ms:.0f}ms "
                    f"vu:{vu_ms:.0f}ms "
                    f"total:{total_ms:.0f}ms"
                )

                # Collect detailed samples for log dump
                self._tuner_perf_frame += 1
                active_wheels = sum(1 for m in result.magnitudes if m > MAGNITUDE_THRESHOLD)
                sample = {
                    'frame': self._tuner_perf_frame,
                    'analyze_ms': round(analyze_ms, 1),
                    'wheel_ms': round(wheel_ms, 1),
                    'vu_ms': round(vu_ms, 1),
                    'total_ms': round(total_ms, 1),
                    'wheels_updated': wheels_updated,
                    'phase_skips': phase_skips,
                    'coords_calls': coords_calls,
                    'color_calls': color_calls,
                    'active_wheels': active_wheels,
                    'canvas_w': self._tuner_canvas.winfo_width(),
                    'canvas_h': self._tuner_canvas.winfo_height(),
                }
                self._tuner_perf_log.append(sample)

                # Dump log every 300 frames (~30s at 10fps)
                if len(self._tuner_perf_log) >= 300:
                    self._tuner_dump_perf_log()

        # FPS measurement
        if self._tuner_show_fps.get():
            import time as _time
            now = _time.perf_counter()
            self._tuner_fps_times.append(now)
            if len(self._tuner_fps_times) > 60:
                self._tuner_fps_times = self._tuner_fps_times[-60:]
            if len(self._tuner_fps_times) >= 2:
                elapsed = self._tuner_fps_times[-1] - self._tuner_fps_times[0]
                if elapsed > 0:
                    actual_fps = (len(self._tuner_fps_times) - 1) / elapsed
                    perf = getattr(self, '_tuner_perf_text', '')
                    self._tuner_update_fps_display(
                        f"{actual_fps:.0f} fps | {perf}")
        elif self._tuner_fps_display:
            self._tuner_canvas.itemconfigure(self._tuner_fps_display, text="")

        interval = FRAME_RATES.get(self._tuner_fps_var.get(), 16)
        self._tuner_anim_id = self.root.after(interval, self._tuner_animate)

    def _tuner_dump_perf_log(self):
        """Write collected perf samples to a debug log file."""
        import os
        log_path = os.path.join(os.path.dirname(__file__), 'tools', 'tuner_perf.log')
        samples = self._tuner_perf_log
        self._tuner_perf_log = []

        if not samples:
            return

        # Compute stats
        n = len(samples)
        def avg(key): return sum(s[key] for s in samples) / n
        def mx(key): return max(s[key] for s in samples)
        def mn(key): return min(s[key] for s in samples)

        # FPS from timestamps
        fps_times = self._tuner_fps_times
        if len(fps_times) >= 2:
            elapsed = fps_times[-1] - fps_times[0]
            actual_fps = (len(fps_times) - 1) / elapsed if elapsed > 0 else 0
        else:
            actual_fps = 0

        total_polys = sum(len(ring) for w in self._tuner_wheels
                          for ring in w._segments)

        with open(log_path, 'a') as f:
            import time as _time
            f.write(f"\n{'='*70}\n")
            f.write(f"Tuner Perf Log — {_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*70}\n")
            f.write(f"Frames sampled: {n}\n")
            f.write(f"Actual FPS: {actual_fps:.1f}\n")
            f.write(f"Target FPS: {self._tuner_fps_var.get()}\n")
            f.write(f"Canvas size: {samples[-1]['canvas_w']}x{samples[-1]['canvas_h']}\n")
            f.write(f"Total polygon objects: {total_polys}\n")
            f.write(f"Wheels: {len(self._tuner_wheels)}\n")
            f.write(f"Rings per wheel: {NUM_RINGS}\n")
            f.write(f"Ring segments: {RING_SEGMENTS}\n\n")

            f.write(f"{'Metric':25s} {'Avg':>8s} {'Min':>8s} {'Max':>8s}\n")
            f.write(f"{'-'*25} {'-'*8} {'-'*8} {'-'*8}\n")
            for key, label in [
                ('analyze_ms', 'FFT analyze'),
                ('wheel_ms', 'Wheel rendering'),
                ('vu_ms', 'VU meter'),
                ('total_ms', 'Total frame'),
                ('wheels_updated', 'Wheels w/ coord update'),
                ('phase_skips', 'Wheels skipped (phase)'),
                ('coords_calls', 'coords() calls'),
                ('color_calls', 'itemconfigure() calls'),
                ('active_wheels', 'Active wheels (signal)'),
            ]:
                f.write(f"{label:25s} {avg(key):8.1f} {mn(key):8.1f} {mx(key):8.1f}\n")

            # Histogram of total frame time
            f.write(f"\nFrame time distribution:\n")
            buckets = [0]*10  # 0-10, 10-20, ... 90-100+ ms
            for s in samples:
                b = min(9, int(s['total_ms'] / 10))
                buckets[b] += 1
            for i, count in enumerate(buckets):
                lo = i * 10
                hi = (i + 1) * 10 if i < 9 else "+"
                bar = '#' * (count * 50 // n) if n > 0 else ''
                f.write(f"  {lo:3d}-{hi:>3s}ms: {count:4d} ({count*100//n:2d}%) {bar}\n")

            # Sample of individual frames (first 20)
            f.write(f"\nFirst 20 frames:\n")
            f.write(f"{'Frame':>6s} {'Analyze':>8s} {'Wheels':>8s} {'VU':>6s} "
                    f"{'Total':>8s} {'Updated':>8s} {'Coords':>8s} {'Active':>7s}\n")
            for s in samples[:20]:
                f.write(f"{s['frame']:6d} {s['analyze_ms']:7.1f}ms "
                        f"{s['wheel_ms']:7.1f}ms {s['vu_ms']:5.1f}ms "
                        f"{s['total_ms']:7.1f}ms {s['wheels_updated']:8d} "
                        f"{s['coords_calls']:8d} {s['active_wheels']:7d}\n")

            f.write(f"\n{'='*70}\n\n")

        print(f"[Tuner] Perf log written to {log_path} ({n} frames)")

    def _tuner_update_fps_display(self, text):
        """Update the FPS counter overlay on the canvas."""
        if not hasattr(self, '_tuner_canvas'):
            return
        try:
            if self._tuner_fps_display:
                self._tuner_canvas.itemconfigure(self._tuner_fps_display, text=text)
            else:
                self._tuner_fps_display = self._tuner_canvas.create_text(
                    10, 10, text=text, anchor="nw",
                    fill="#888888", font=("Helvetica", 9),
                    tags="fps_overlay")
        except tk.TclError:
            # Canvas item was destroyed (e.g. by wheel rebuild)
            self._tuner_fps_display = None

    def _tuner_show_stream_error(self, error_msg):
        """Show audio stream error on the tuner canvas with a retry option."""
        self._tuner_running = False
        self._tuner_set_pilot(False)
        if hasattr(self, '_tuner_canvas'):
            c = self._tuner_canvas
            cx = c.winfo_width() / 2
            cy = c.winfo_height() / 2
            c.delete("error")
            c.create_text(cx, cy - 15, text=error_msg,
                          fill="#FF4444", font=("Helvetica", 12),
                          tags="error")
            c.create_text(cx, cy + 15, text="Click here to retry",
                          fill="#4488FF", font=("Helvetica", 11, "underline"),
                          tags=("error", "error_retry"))
            c.tag_bind("error_retry", "<Button-1>",
                       lambda e: self._tuner_retry())

    def _tuner_retry(self):
        """Retry starting the tuner after a stream error."""
        if self._tuner_engine:
            self._tuner_engine.last_error = None
        self._tuner_start()

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
            "ring_brightness": self._tuner_ring_brightness if hasattr(self, '_tuner_ring_brightness') else 100,
            "overall_brightness": self._tuner_overall_brightness if hasattr(self, '_tuner_overall_brightness') else 80,
            "faceplate_color": self._tuner_faceplate_color if hasattr(self, '_tuner_faceplate_color') else DEFAULT_FACEPLATE,
            "show_fps": self._tuner_show_fps.get() if hasattr(self, '_tuner_show_fps') else False,
        }
