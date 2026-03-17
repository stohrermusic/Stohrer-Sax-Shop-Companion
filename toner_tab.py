"""
Toner tab mixin for Stohrer Sax Shop Companion.

Harmonic tone analyzer for saxophone. Shows a live spectrum analyzer
(full FFT or harmonic-only bars) on the left and VU-style tone
descriptor gauges on the right. Auto-detects fundamental pitch and
analyzes harmonic content for real-time tone quality feedback.

Includes a tone profile system: guided capture sessions build up
harmonic fingerprints of individual horns over time, with comparison
overlay for before/after analysis and horn-to-horn comparison.

Requires: numpy, sounddevice (graceful fallback if unavailable)
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import math
import sys
import time

IS_MACOS = sys.platform == 'darwin'

try:
    from toner_engine import (
        TonerEngine, AUDIO_AVAILABLE, PITCH_CLASSES, MAX_HARMONICS,
        CAPTURE_DELAY_S, CAPTURE_DURATION_S, MIN_PROFILE_NOTES, SAX_TYPES,
        SAX_TRANSPOSITIONS, DEFAULT_LIBRARY, average_captures,
        compute_fingerprint, load_tone_profiles, save_tone_profiles,
        flatten_profiles,
    )
    _TONER_IMPORTS_OK = True
except ImportError:
    _TONER_IMPORTS_OK = False
    AUDIO_AVAILABLE = False
    PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

try:
    from config import TONE_PROFILES_FILE
except ImportError:
    TONE_PROFILES_FILE = "tone_profiles.json"


# ============================================
# CONSTANTS
# ============================================

BG_COLOR = "#1A1A1A"
CTRL_BG = "#2A2A2A"
AMBER = "#D4920A"
TICK_COLOR = "#2A1A00"
LABEL_DIM = "#888888"
LABEL_BRIGHT = "#FFFFFF"
SPECTRUM_BAR_COLOR = "#666666"
HARMONIC_COLOR = "#D4920A"       # Amber for detected harmonics
FUNDAMENTAL_COLOR = "#00CC66"    # Green for fundamental
GHOST_COLOR = "#4444AA"          # Blue ghost overlay for comparison
FRAME_RATES = {"30": 33, "60": 16}

# Gauge arc geometry
GAUGE_ARC_START = 155.0
GAUGE_ARC_END = 25.0


def _format_profile_info(p):
    """Format a profile dict into a readable info string."""
    info = f"{p.get('horn_make', '')} {p.get('horn_model', '')}".strip()
    if p.get('serial'):
        info += f"  (s/n {p['serial']})"
    if p.get('horn_type'):
        info += f"\nType: {p['horn_type']}"
    if p.get('player'):
        info += f"  |  Player: {p['player']}"
    parts = []
    if p.get('mouthpiece'):
        parts.append(p['mouthpiece'])
    if p.get('reed'):
        parts.append(p['reed'])
    if parts:
        info += f"\nSetup: {', '.join(parts)}"
    sessions = p.get('sessions', [])
    total_caps = sum(len(s.get('captures', [])) for s in sessions)
    notes = set()
    for s in sessions:
        for c in s.get('captures', []):
            notes.add(c.get('note', ''))
    info += f"\n{len(sessions)} sessions, {total_caps} captures, {len(notes)} unique notes"
    if 0 < len(notes) < MIN_PROFILE_NOTES:
        info += f" (need {MIN_PROFILE_NOTES - len(notes)} more)"
    elif len(notes) >= MIN_PROFILE_NOTES:
        info += " \u2713"
    if sessions:
        info += f"\nLast session: {sessions[-1].get('date', '?')}"
    if p.get('notes'):
        info += f"\nNotes: {p['notes']}"
    return info


# ============================================
# TONER TAB MIXIN
# ============================================

class TonerTabMixin:
    """Mixin class that adds the Toner tab to the main application."""

    def _init_toner_state(self):
        """Initialize toner state. Called from __init__."""
        self._toner_engine = None
        self._toner_running = False
        self._toner_anim_id = None
        self._toner_spectrum_bars = []
        self._toner_harmonic_markers = []
        self._toner_ghost_markers = []     # Ghost overlay bars for comparison
        self._toner_bars_built = False
        # Smoothed descriptor values for damped needle movement
        self._toner_smooth = {
            'resonance': 0.5, 'richness': 0.0,
            'brightness': 0.0, 'darkness': 0.0, 'fullness': 0.0,
        }
        # Profile system state — nested: {library: {profile_name: data}}
        self._toner_profiles = {}
        self._toner_active_library = None   # Library name
        self._toner_active_profile = None   # Profile name within library
        self._toner_active_session = None   # Current capture session dict
        # Capture state machine: None, 'listening', 'delay', 'recording', 'cooldown'
        self._toner_capture_state = None
        self._toner_capture_start = 0.0
        self._toner_capture_frames = []     # Accumulated frames during capture
        self._toner_capture_note = ""       # Note being captured
        self._toner_stable_note = ""        # Note currently being held steady
        self._toner_stable_count = 0        # Consecutive frames of same note
        self._toner_comparison = None       # Fingerprint dict for ghost overlay
        # How many consecutive frames of same note to trigger auto-capture (~1s at 30fps)
        self._toner_stable_threshold = 25

    def create_toner_tab(self, parent):
        """Build the Toner tab UI."""
        if not _TONER_IMPORTS_OK or not AUDIO_AVAILABLE:
            self._create_toner_fallback(parent)
            return

        toner_settings = self.settings.get("toner_settings", {})

        self._toner_engine = TonerEngine()
        self._toner_engine.set_reference_pitch(
            toner_settings.get("reference_pitch", 440.0))
        self._toner_engine.set_sensitivity(
            toner_settings.get("sensitivity", 50))

        self._toner_fps_var = tk.StringVar(
            value=str(toner_settings.get("fps", "30")))
        self._toner_view_var = tk.StringVar(
            value=toner_settings.get("view_mode", "spectrum"))
        self._toner_scale_var = tk.StringVar(
            value=toner_settings.get("scale_mode", "linear"))

        self._toner_concert_pitch = tk.BooleanVar(
            value=toner_settings.get("concert_pitch", False))

        # Bias sliders: visual offset per descriptor (-50 to +50, display only)
        saved_bias = toner_settings.get("gauge_bias", {})
        self._toner_bias_vars = {}
        for key in ("resonance", "richness", "brightness", "darkness"):
            self._toner_bias_vars[key] = tk.IntVar(
                value=saved_bias.get(key, 0))

        # Load profiles
        self._toner_profiles = load_tone_profiles(TONE_PROFILES_FILE)

        bg = BG_COLOR

        # --- Main container (dark, skip theme) ---
        self._toner_main_frame = tk.Frame(parent, bg=bg)
        self._toner_main_frame._skip_theme = True
        self._toner_main_frame.pack(fill="both", expand=True)

        # --- Top area: spectrum (left) + gauges (right) ---
        top_frame = tk.Frame(self._toner_main_frame, bg=bg)
        top_frame._skip_theme = True
        top_frame.pack(fill="both", expand=True, padx=5, pady=(5, 0))
        top_frame.columnconfigure(0, weight=3)
        top_frame.columnconfigure(1, weight=1)
        top_frame.rowconfigure(0, weight=1)

        # --- LEFT: Spectrum analyzer canvas ---
        self._toner_spectrum_canvas = tk.Canvas(
            top_frame, bg="#0D0D0D", highlightthickness=0, borderwidth=0)
        self._toner_spectrum_canvas._dark_canvas = True
        self._toner_spectrum_canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self._toner_spectrum_canvas.bind("<Configure>", self._toner_on_canvas_resize)

        # --- RIGHT: Gauge panel (grid layout for compact vertical fit) ---
        # Layout:
        #   row 0: [bias] [Intonation gauge] [Note + freq]
        #   row 1: [bias] [Resonance gauge]
        #   row 2: [bias] [Richness gauge]
        #   row 3: [bias] [Brightness gauge] [bright]
        #   row 4: [bias] [Darkness gauge]   [+ FULL]
        #                                     [dark ]
        # Outer frame fills the grid cell, inner frame centered via pack
        gauge_outer = tk.Frame(top_frame, bg=bg)
        gauge_outer._skip_theme = True
        gauge_outer.grid(row=0, column=1, sticky="nsew")

        gauge_frame = tk.Frame(gauge_outer, bg=bg)
        gauge_frame._skip_theme = True
        gauge_frame.pack(expand=True)

        gauge_row = 0

        def _make_bias_slider(parent, key, row):
            """Create a vertical bias slider in column 0."""
            sl = tk.Scale(parent, variable=self._toner_bias_vars[key],
                          from_=50, to=-50, orient="vertical",
                          length=80, width=8, showvalue=False,
                          bg=bg, fg="#888888", troughcolor="#333333",
                          highlightthickness=0)
            sl.grid(row=row, column=0, padx=(0, 2), sticky="ns")

        # --- Row 0: Intonation gauge + note display ---
        self._toner_intonation_canvas = tk.Canvas(
            gauge_frame, bg=bg, highlightthickness=0, width=260, height=110)
        self._toner_intonation_canvas._dark_canvas = True
        self._toner_intonation_canvas.grid(row=gauge_row, column=1,
                                            pady=(2, 0))
        self._toner_intonation_gauge = self._toner_build_intonation_gauge(
            self._toner_intonation_canvas)
        self._toner_smooth_cents = 0.0

        # Note display to the right of intonation gauge
        note_frame = tk.Frame(gauge_frame, bg=bg)
        note_frame._skip_theme = True
        note_frame.grid(row=gauge_row, column=2, sticky="ns", padx=(6, 2))

        self._toner_note_label = tk.Label(
            note_frame, text="\u2014", bg=bg, fg=LABEL_DIM,
            font=("Helvetica", 24, "bold"))
        self._toner_note_label.pack(expand=True)

        self._toner_freq_label = tk.Label(
            note_frame, text="", bg=bg, fg=LABEL_DIM,
            font=("Helvetica", 8))
        self._toner_freq_label.pack()

        gauge_row += 1

        # --- Rows 1-4: Descriptor gauges with bias sliders ---
        self._toner_gauges = {}
        gauge_defs = [
            ("resonance", "Dissonant", "Resonant"),
            ("richness", "Pure", "Rich"),
            ("brightness", "", "Bright"),
            ("darkness", "", "Dark"),
        ]

        for key, left_label, right_label in gauge_defs:
            _make_bias_slider(gauge_frame, key, gauge_row)

            cv = tk.Canvas(gauge_frame, bg=bg, highlightthickness=0,
                           width=260, height=110)
            cv._dark_canvas = True
            cv.grid(row=gauge_row, column=1, pady=(2, 0))
            gauge_data = self._toner_build_gauge(cv, left_label, right_label)
            self._toner_gauges[key] = gauge_data

            # Place FULL indicator to the right, spanning brightness+darkness
            if key == "brightness":
                full_frame = tk.Frame(gauge_frame, bg=bg)
                full_frame._skip_theme = True
                full_frame.grid(row=gauge_row, column=2, rowspan=2,
                                sticky="", padx=(6, 2))

                # "bright + dark = FULL" indicator
                tk.Label(full_frame, text="bright", bg=bg, fg="#555555",
                         font=("Helvetica", 7)).pack()
                tk.Label(full_frame, text="+", bg=bg, fg="#555555",
                         font=("Helvetica", 7)).pack()
                tk.Label(full_frame, text="dark", bg=bg, fg="#555555",
                         font=("Helvetica", 7)).pack()
                tk.Label(full_frame, text="=", bg=bg, fg="#555555",
                         font=("Helvetica", 7)).pack()

                self._toner_full_canvas = tk.Canvas(
                    full_frame, bg=bg, highlightthickness=0, width=30, height=30)
                self._toner_full_canvas._dark_canvas = True
                self._toner_full_canvas.pack(pady=(2, 0))

                tk.Label(full_frame, text="FULL", bg=bg, fg=LABEL_DIM,
                         font=("Helvetica", 8, "bold")).pack(pady=(2, 0))

            gauge_row += 1

        self._toner_full_glow = self._toner_full_canvas.create_oval(
            3, 3, 27, 27, fill="#1A0A00", outline="")
        self._toner_full_lamp = self._toner_full_canvas.create_oval(
            6, 6, 24, 24, fill="#331100", outline="#555555", width=1)

        # --- Capture status bar (hidden until active) ---
        self._toner_capture_frame = tk.Frame(self._toner_main_frame, bg="#333300")
        self._toner_capture_frame._skip_theme = True
        # Not packed yet — shown during capture

        self._toner_capture_label = tk.Label(
            self._toner_capture_frame, text="", bg="#333300", fg="#FFCC00",
            font=("Helvetica", 11, "bold"))
        self._toner_capture_label.pack(side="left", padx=10, pady=4)

        self._toner_capture_progress = tk.Label(
            self._toner_capture_frame, text="", bg="#333300", fg="#AAAAAA",
            font=("Helvetica", 10))
        self._toner_capture_progress.pack(side="left", padx=5, pady=4)

        self._toner_capture_cancel_btn = tk.Button(
            self._toner_capture_frame, text="Cancel", font=("Helvetica", 9),
            command=self._toner_cancel_capture)
        self._toner_capture_cancel_btn.pack(side="right", padx=10, pady=4)

        # --- Bottom: controls ---
        ctrl_bg = "systemWindowBackgroundColor" if IS_MACOS else CTRL_BG
        ctrl_fg = "white" if not IS_MACOS else "systemTextColor"
        self._toner_ctrl_bg = ctrl_bg
        self._toner_ctrl_fg = ctrl_fg

        ctrl_frame = tk.Frame(self._toner_main_frame, bg=ctrl_bg, padx=6, pady=6)
        ctrl_frame._skip_theme = True
        ctrl_frame.pack(fill="x", padx=5, pady=(0, 4))

        eq_lbl_font = ("Helvetica", 7)

        # Sax type selector (sets Benade break frequency for descriptors)
        sax_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        sax_frame._skip_theme = True
        sax_frame.pack(side="left", padx=(0, 12))

        tk.Label(sax_frame, text="SAX", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        self._toner_sax_var = tk.StringVar(
            value=toner_settings.get("sax_type", "Alto"))
        sax_combo = ttk.Combobox(
            sax_frame, textvariable=self._toner_sax_var,
            values=SAX_TYPES, state="readonly", width=9)
        sax_combo.pack(side="left")
        sax_combo.bind("<<ComboboxSelected>>", self._toner_on_sax_type_changed)
        # Apply initial sax type to engine
        if self._toner_engine:
            self._toner_engine.set_sax_type(self._toner_sax_var.get())

        # Concert pitch toggle
        self._toner_concert_btn = tk.Checkbutton(
            sax_frame, text="Concert", variable=self._toner_concert_pitch,
            bg=ctrl_bg, fg=ctrl_fg, selectcolor=ctrl_bg,
            activebackground=ctrl_bg, font=("Helvetica", 8))
        self._toner_concert_btn.pack(side="left", padx=(4, 0))

        # Sensitivity slider
        sens_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        sens_frame._skip_theme = True
        sens_frame.pack(side="left", padx=(0, 12))

        tk.Label(sens_frame, text="SENS", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        self._toner_sens_var = tk.IntVar(
            value=toner_settings.get("sensitivity", 50))
        tk.Scale(sens_frame, variable=self._toner_sens_var,
                 from_=0, to=100, orient="horizontal", length=80, width=12,
                 showvalue=False, bg=ctrl_bg, fg=ctrl_fg,
                 troughcolor="#444444", highlightthickness=0,
                 command=self._toner_on_sensitivity_changed).pack(side="left")

        # A= reference pitch
        pitch_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        pitch_frame._skip_theme = True
        pitch_frame.pack(side="left", padx=(0, 12))

        tk.Label(pitch_frame, text="A =", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        self._toner_pitch_var = tk.DoubleVar(
            value=toner_settings.get("reference_pitch", 440.0))
        tk.Spinbox(pitch_frame, textvariable=self._toner_pitch_var,
                   from_=420, to=460, increment=0.5, width=5,
                   bg="#333333", fg="white", buttonbackground="#444444",
                   command=self._toner_on_pitch_changed).pack(side="left")
        tk.Label(pitch_frame, text="Hz", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(2, 0))

        # View mode toggle
        view_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        view_frame._skip_theme = True
        view_frame.pack(side="left", padx=(0, 12))

        tk.Label(view_frame, text="VIEW", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        tk.Radiobutton(
            view_frame, text="Spectrum", variable=self._toner_view_var,
            value="spectrum", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 9),
        ).pack(side="left", padx=(0, 4))
        tk.Radiobutton(
            view_frame, text="Bars", variable=self._toner_view_var,
            value="bars", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 9),
        ).pack(side="left")

        # Scale toggle (Linear / dB)
        scale_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        scale_frame._skip_theme = True
        scale_frame.pack(side="left", padx=(0, 12))

        tk.Label(scale_frame, text="SCALE", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        tk.Radiobutton(
            scale_frame, text="Linear", variable=self._toner_scale_var,
            value="linear", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 9),
        ).pack(side="left", padx=(0, 4))
        tk.Radiobutton(
            scale_frame, text="dB", variable=self._toner_scale_var,
            value="db", bg=ctrl_bg, fg=ctrl_fg,
            selectcolor=ctrl_bg, activebackground=ctrl_bg,
            font=("Helvetica", 9),
        ).pack(side="left")

        # FPS selector
        fps_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        fps_frame._skip_theme = True
        fps_frame.pack(side="left", padx=(0, 12))

        tk.Label(fps_frame, text="FPS", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(side="left", padx=(0, 4))
        ttk.Combobox(
            fps_frame, textvariable=self._toner_fps_var,
            values=["30", "60"], state="readonly", width=3
        ).pack(side="left")

        # --- Profile controls (right side of control strip) ---
        prof_frame = tk.Frame(ctrl_frame, bg=ctrl_bg)
        prof_frame._skip_theme = True
        prof_frame.pack(side="right", padx=(12, 0))

        tk.Button(prof_frame, text="Profile...",
                  font=("Helvetica", 9),
                  command=self._toner_open_profile_dialog).pack(side="left", padx=(0, 4))

        self._toner_capture_btn = tk.Button(
            prof_frame, text="Capture",
            font=("Helvetica", 9),
            command=self._toner_toggle_capture)
        self._toner_capture_btn.pack(side="left", padx=(0, 4))

        tk.Button(prof_frame, text="Compare...",
                  font=("Helvetica", 9),
                  command=self._toner_open_compare_dialog).pack(side="left")

    def _create_toner_fallback(self, parent):
        """Fallback UI when audio libraries are unavailable."""
        bg = self.root.cget('bg')
        tk.Label(parent, text="Tone Analyzer requires numpy and sounddevice.\n\n"
                 "Install with:  pip install numpy sounddevice",
                 bg=bg, fg="gray", font=("Helvetica", 12),
                 justify="center").pack(expand=True)

    # ------------------------------------------------------------------
    # GAUGE BUILDER (VU-meter style)
    # ------------------------------------------------------------------

    def _toner_build_gauge(self, cv, left_label, right_label):
        """Draw a VU-style arc gauge and return dict with IDs for animation."""
        cv_w, cv_h = 260, 110
        cv.configure(width=cv_w, height=cv_h)
        cx = cv_w // 2
        cy = cv_h - 8
        r = 65
        arc_start = GAUGE_ARC_START
        arc_end = GAUGE_ARC_END

        bg = BG_COLOR
        tick_color = TICK_COLOR

        cv.create_rectangle(2, 2, cv_w - 2, cv_h - 2,
                            fill="#222222", outline="#3A3A3A", width=1)
        cv.create_rectangle(5, 5, cv_w - 5, cv_h - 5,
                            fill=AMBER, outline="#B87D08", width=1)

        for i in range(11):
            frac = i / 10.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)

            if i == 5:
                tick_len, tick_w = 10, 2
            elif i % 2 == 0:
                tick_len, tick_w = 7, 1
            else:
                tick_len, tick_w = 4, 1

            x_o = cx + r * math.cos(angle_rad)
            y_o = cy - r * math.sin(angle_rad)
            x_i = cx + (r - tick_len) * math.cos(angle_rad)
            y_i = cy - (r - tick_len) * math.sin(angle_rad)
            cv.create_line(x_i, y_i, x_o, y_o, fill=tick_color, width=tick_w)

        arc_points = []
        for i in range(41):
            frac = i / 40.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            arc_points.extend([
                cx + r * math.cos(angle_rad),
                cy - r * math.sin(angle_rad),
            ])
        cv.create_line(*arc_points, fill=tick_color, width=1, smooth=True)

        label_font = ("Helvetica", 7)
        if left_label:
            left_angle = math.radians(arc_start + 5)
            cv.create_text(cx + (r + 12) * math.cos(left_angle),
                           cy - (r + 12) * math.sin(left_angle),
                           text=left_label, fill=tick_color, font=label_font,
                           anchor="e")
        if right_label:
            right_angle = math.radians(arc_end - 5)
            cv.create_text(cx + (r + 12) * math.cos(right_angle),
                           cy - (r + 12) * math.sin(right_angle),
                           text=right_label, fill=tick_color, font=label_font,
                           anchor="w")

        needle_len = r - 12
        center_angle = math.radians((arc_start + arc_end) / 2)
        nx = cx + needle_len * math.cos(center_angle)
        ny = cy - needle_len * math.sin(center_angle)

        shadow_id = cv.create_line(
            cx + 1, cy + 1, nx + 1, ny + 1,
            fill="#8A6500", width=2, capstyle="round")
        needle_id = cv.create_line(
            cx, cy, nx, ny, fill="#1A1200", width=2, capstyle="round")

        cv.create_oval(cx - 4, cy - 4, cx + 4, cy + 4,
                       fill="#2A1A00", outline="#1A1200", width=1)

        return {
            'canvas': cv, 'cx': cx, 'cy': cy,
            'needle_len': needle_len, 'needle_id': needle_id,
            'shadow_id': shadow_id, 'arc_start': arc_start, 'arc_end': arc_end,
        }

    def _toner_build_intonation_gauge(self, cv):
        """Build the intonation VU gauge (±50 cents, same style as tuner VU)."""
        cv_w, cv_h = 260, 110
        cv.configure(width=cv_w, height=cv_h)
        cx = cv_w // 2
        cy = cv_h - 8
        r = 65
        arc_start = GAUGE_ARC_START
        arc_end = GAUGE_ARC_END

        tick_color = TICK_COLOR

        # Bezel + amber panel
        cv.create_rectangle(2, 2, cv_w - 2, cv_h - 2,
                            fill="#222222", outline="#3A3A3A", width=1)
        cv.create_rectangle(5, 5, cv_w - 5, cv_h - 5,
                            fill=AMBER, outline="#B87D08", width=1)

        # Tick marks (21 ticks for -50 to +50 in steps of 5)
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

            x_o = cx + r * math.cos(angle_rad)
            y_o = cy - r * math.sin(angle_rad)
            x_i = cx + (r - tick_len) * math.cos(angle_rad)
            y_i = cy - (r - tick_len) * math.sin(angle_rad)
            cv.create_line(x_i, y_i, x_o, y_o, fill=tick_color, width=tick_w)

        # Arc line
        arc_points = []
        for i in range(51):
            frac = i / 50.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            arc_points.extend([
                cx + r * math.cos(angle_rad),
                cy - r * math.sin(angle_rad),
            ])
        cv.create_line(*arc_points, fill=tick_color, width=1, smooth=True)

        # Scale labels
        label_r = r + 11
        label_font = ("Helvetica", 7)
        for cents, label in [
            (-50, "50"), (-30, "30"), (-10, "10"),
            (0, "0"), (10, "10"), (30, "30"), (50, "50"),
        ]:
            frac = (cents + 50) / 100.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = cx + label_r * math.cos(angle_rad)
            ly = cy - label_r * math.sin(angle_rad)
            color = "#1B5E00" if cents == 0 else tick_color
            cv.create_text(lx, ly, text=label, fill=color,
                           font=label_font, anchor="center")

        # Flat/sharp at extremes
        for cents, label in [(-50, "\u266d"), (50, "\u266f")]:
            frac = (cents + 50) / 100.0
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = cx + (label_r + 10) * math.cos(angle_rad)
            ly = cy - (label_r + 10) * math.sin(angle_rad)
            cv.create_text(lx, ly, text=label, fill=tick_color,
                           font=("Helvetica", 10, "bold"), anchor="center")

        # Cents readout text (updated each frame)
        self._toner_cents_text_id = cv.create_text(
            cx, cy - r - 18, text="", fill=tick_color,
            font=("Helvetica", 9, "bold"), anchor="center")

        # Needle + shadow
        needle_len = r - 16
        center_angle = math.radians((arc_start + arc_end) / 2)
        nx = cx + needle_len * math.cos(center_angle)
        ny = cy - needle_len * math.sin(center_angle)

        shadow_id = cv.create_line(
            cx + 1, cy + 1, nx + 1, ny + 1,
            fill="#8A6500", width=3, capstyle="round")
        needle_id = cv.create_line(
            cx, cy, nx, ny, fill="#1A1200", width=3, capstyle="round")

        cv.create_oval(cx - 5, cy - 5, cx + 5, cy + 5,
                       fill="#2A1A00", outline="#1A1200", width=1)

        return {
            'canvas': cv, 'cx': cx, 'cy': cy,
            'needle_len': needle_len, 'needle_id': needle_id,
            'shadow_id': shadow_id, 'arc_start': arc_start, 'arc_end': arc_end,
        }

    def _toner_update_intonation(self, cents):
        """Update the intonation gauge with cents offset, with damping."""
        gauge = self._toner_intonation_gauge
        if not gauge:
            return

        damping = 0.18
        self._toner_smooth_cents += (cents - self._toner_smooth_cents) * damping

        clamped = max(-50.0, min(50.0, self._toner_smooth_cents))
        frac = (clamped + 50.0) / 100.0
        angle_deg = gauge['arc_start'] + (gauge['arc_end'] - gauge['arc_start']) * frac
        angle_rad = math.radians(angle_deg)
        nx = gauge['cx'] + gauge['needle_len'] * math.cos(angle_rad)
        ny = gauge['cy'] - gauge['needle_len'] * math.sin(angle_rad)

        cv = gauge['canvas']
        cv.coords(gauge['shadow_id'],
                  gauge['cx'] + 1, gauge['cy'] + 1, nx + 1, ny + 1)
        cv.coords(gauge['needle_id'],
                  gauge['cx'], gauge['cy'], nx, ny)

        # Update cents readout
        if abs(self._toner_smooth_cents) < 1.0:
            cv.itemconfigure(self._toner_cents_text_id,
                             text="IN TUNE", fill="#1B5E00")
        elif self._toner_smooth_cents > 0:
            cv.itemconfigure(self._toner_cents_text_id,
                             text=f"+{self._toner_smooth_cents:.0f}\u00a2",
                             fill=TICK_COLOR)
        else:
            cv.itemconfigure(self._toner_cents_text_id,
                             text=f"{self._toner_smooth_cents:.0f}\u00a2",
                             fill=TICK_COLOR)

    def _toner_update_gauge(self, key, value):
        """Update a descriptor gauge needle to the given 0.0-1.0 value with damping.

        Applies the bias slider offset for display only. The bias shifts the
        needle position without affecting captured data.
        """
        gauge = self._toner_gauges.get(key)
        if not gauge:
            return

        damping = 0.15
        self._toner_smooth[key] += (value - self._toner_smooth[key]) * damping

        # Apply bias: slider range -50..+50 maps to -0.5..+0.5 offset
        bias = self._toner_bias_vars.get(key)
        offset = bias.get() / 100.0 if bias else 0.0
        frac = max(0.0, min(1.0, self._toner_smooth[key] + offset))

        angle_deg = gauge['arc_start'] + (gauge['arc_end'] - gauge['arc_start']) * frac
        angle_rad = math.radians(angle_deg)
        nx = gauge['cx'] + gauge['needle_len'] * math.cos(angle_rad)
        ny = gauge['cy'] - gauge['needle_len'] * math.sin(angle_rad)

        cv = gauge['canvas']
        cv.coords(gauge['shadow_id'],
                  gauge['cx'] + 1, gauge['cy'] + 1, nx + 1, ny + 1)
        cv.coords(gauge['needle_id'],
                  gauge['cx'], gauge['cy'], nx, ny)

    # ------------------------------------------------------------------
    # SPECTRUM RENDERING
    # ------------------------------------------------------------------

    def _toner_build_spectrum_bars(self):
        """Pre-create rectangle items for the spectrum display."""
        cv = self._toner_spectrum_canvas
        cv.delete("all")
        self._toner_spectrum_bars = []
        self._toner_harmonic_markers = []
        self._toner_ghost_markers = []
        self._toner_bars_built = False

        w = cv.winfo_width()
        h = cv.winfo_height()
        if w < 50 or h < 50:
            return

        self._toner_num_bars = min(300, max(60, w // 3))
        bar_w = w / self._toner_num_bars

        for i in range(self._toner_num_bars):
            x0 = i * bar_w
            x1 = x0 + bar_w - 1
            rect = cv.create_rectangle(x0, h, x1, h, fill=SPECTRUM_BAR_COLOR,
                                       outline="", width=0)
            self._toner_spectrum_bars.append(rect)

        # Harmonic markers
        for i in range(12):
            marker = cv.create_rectangle(0, 0, 0, 0, fill=HARMONIC_COLOR,
                                         outline="", width=0, state="hidden")
            self._toner_harmonic_markers.append(marker)

        # Ghost overlay markers (for comparison profile)
        for i in range(12):
            ghost = cv.create_line(0, 0, 0, 0, fill=GHOST_COLOR,
                                   width=2, state="hidden")
            self._toner_ghost_markers.append(ghost)

        # Frequency axis labels
        label_font = ("Helvetica", 7)
        for freq in [100, 500, 1000, 2000, 4000, 8000]:
            x = (freq / 8000.0) * w
            cv.create_text(x, h - 3, text=f"{freq}" if freq < 1000 else f"{freq // 1000}k",
                           fill="#555555", font=label_font, anchor="s")
            cv.create_line(x, 0, x, h - 12, fill="#333333", width=1, dash=(2, 4))

        # Note name overlay
        self._toner_spectrum_note = cv.create_text(
            10, 10, text="", fill=FUNDAMENTAL_COLOR,
            font=("Helvetica", 16, "bold"), anchor="nw")

        # Comparison label
        self._toner_compare_label = cv.create_text(
            10, 32, text="", fill=GHOST_COLOR,
            font=("Helvetica", 9), anchor="nw")

        self._toner_bars_built = True

    def _toner_db_to_height(self, db_val, max_height, db_range=60.0):
        """Convert a dB value to a bar height, respecting the current scale mode.

        In dB mode: linear mapping of dB from -db_range..0 to 0..max_height.
        In linear mode: convert dB back to linear amplitude, scale to max_height.
        """
        if self._toner_scale_var.get() == "linear":
            # dB to linear amplitude (0 dB = 1.0, -60 dB = 0.001)
            linear = 10.0 ** (max(-db_range, min(0.0, db_val)) / 20.0)
            return max(0.0, linear * max_height)
        else:
            return max(0.0, (db_val + db_range) / db_range) * max_height

    def _toner_render_spectrum(self, result):
        """Update spectrum bars from analysis result."""
        cv = self._toner_spectrum_canvas
        if not self._toner_bars_built or not self._toner_spectrum_bars:
            return

        w = cv.winfo_width()
        h = cv.winfo_height()
        num_bars = len(self._toner_spectrum_bars)

        view_mode = self._toner_view_var.get()

        if view_mode == "bars" and result.harmonic_bars:
            # --- Simple bars mode ---
            for rect in self._toner_spectrum_bars:
                cv.coords(rect, 0, h, 0, h)
            for marker in self._toner_harmonic_markers:
                cv.itemconfigure(marker, state="hidden")

            if result.harmonics:
                bar_width = max(8, w / (len(result.harmonics) * 2.5))
                total_w = bar_width * len(result.harmonics) * 2
                start_x = (w - total_w) / 2

                for idx, hi in enumerate(result.harmonics):
                    if idx >= len(self._toner_spectrum_bars):
                        break
                    bar_h = self._toner_db_to_height(hi.magnitude_db, h - 30)
                    x0 = start_x + idx * bar_width * 2
                    x1 = x0 + bar_width
                    color = FUNDAMENTAL_COLOR if hi.harmonic_number == 1 else HARMONIC_COLOR
                    cv.coords(self._toner_spectrum_bars[idx],
                              x0, h - 15, x1, h - 15 - bar_h)
                    cv.itemconfigure(self._toner_spectrum_bars[idx], fill=color)

                # Ghost overlay in bars mode
                self._toner_render_ghost_bars(result, start_x, bar_width, h)
        else:
            # --- Full spectrum mode ---
            if result.spectrum_db is not None and len(result.spectrum_db) > 0:
                spec_db = result.spectrum_db
                spec_len = len(spec_db)

                for i in range(num_bars):
                    frac = i / num_bars
                    bin_idx = min(int(frac * spec_len), spec_len - 1)
                    lo = max(0, bin_idx - 1)
                    hi_b = min(spec_len, bin_idx + 2)
                    db_val = float(max(spec_db[lo:hi_b]))
                    bar_h = self._toner_db_to_height(db_val, h - 20, db_range=80.0)
                    bar_w = w / num_bars
                    x0 = i * bar_w
                    x1 = x0 + bar_w - 1
                    cv.coords(self._toner_spectrum_bars[i],
                              x0, h - 10, x1, h - 10 - bar_h)
                    cv.itemconfigure(self._toner_spectrum_bars[i],
                                     fill=SPECTRUM_BAR_COLOR)

                # Highlight harmonics
                for m_idx, marker in enumerate(self._toner_harmonic_markers):
                    if m_idx < len(result.harmonics):
                        hi = result.harmonics[m_idx]
                        freq_frac = hi.expected_freq / 8000.0
                        if 0 <= freq_frac <= 1.0:
                            x = freq_frac * w
                            if hi.harmonic_number == 1:
                                bar_h = self._toner_db_to_height(0.0, h - 20, db_range=80.0)
                            else:
                                bar_h = self._toner_db_to_height(hi.magnitude_db, h - 20, db_range=80.0)
                            color = FUNDAMENTAL_COLOR if hi.harmonic_number == 1 else HARMONIC_COLOR
                            hw = max(2, w / num_bars)
                            cv.coords(marker, x - hw / 2, h - 10,
                                      x + hw / 2, h - 10 - bar_h)
                            cv.itemconfigure(marker, fill=color, state="normal")
                        else:
                            cv.itemconfigure(marker, state="hidden")
                    else:
                        cv.itemconfigure(marker, state="hidden")

                # Ghost overlay in spectrum mode
                self._toner_render_ghost_spectrum(result, w, h)
            else:
                for rect in self._toner_spectrum_bars:
                    cv.coords(rect, 0, h, 0, h)
                for marker in self._toner_harmonic_markers:
                    cv.itemconfigure(marker, state="hidden")

        # Note name overlay (transposed)
        if result.fundamental_note:
            display_note = self._toner_transpose_note(result.fundamental_note)
            cv.itemconfigure(self._toner_spectrum_note, text=display_note)
        else:
            cv.itemconfigure(self._toner_spectrum_note, text="")

    def _toner_render_ghost_bars(self, result, start_x, bar_width, h):
        """Render comparison ghost overlay in bars mode."""
        cv = self._toner_spectrum_canvas
        fp = self._toner_comparison
        if not fp:
            for g in self._toner_ghost_markers:
                cv.itemconfigure(g, state="hidden")
            return

        # Use per-note data if available for current note, else overall
        ghost_db = fp.get('harmonics_db', [])
        note = result.fundamental_note
        if note and 'per_note' in fp and note in fp['per_note']:
            ghost_db = fp['per_note'][note].get('harmonics_db', ghost_db)

        for idx, g in enumerate(self._toner_ghost_markers):
            if idx < len(ghost_db):
                bar_h = self._toner_db_to_height(ghost_db[idx], h - 30)
                x_center = start_x + idx * bar_width * 2 + bar_width / 2
                cv.coords(g, x_center, h - 15, x_center, h - 15 - bar_h)
                cv.itemconfigure(g, state="normal")
            else:
                cv.itemconfigure(g, state="hidden")

    def _toner_render_ghost_spectrum(self, result, w, h):
        """Render comparison ghost overlay in spectrum mode."""
        cv = self._toner_spectrum_canvas
        fp = self._toner_comparison
        if not fp or result.fundamental_freq <= 0:
            for g in self._toner_ghost_markers:
                cv.itemconfigure(g, state="hidden")
            return

        ghost_db = fp.get('harmonics_db', [])
        note = result.fundamental_note
        if note and 'per_note' in fp and note in fp['per_note']:
            ghost_db = fp['per_note'][note].get('harmonics_db', ghost_db)

        f0 = result.fundamental_freq
        for idx, g in enumerate(self._toner_ghost_markers):
            if idx < len(ghost_db):
                freq = f0 * (idx + 1)
                freq_frac = freq / 8000.0
                if 0 <= freq_frac <= 1.0:
                    x = freq_frac * w
                    bar_h = self._toner_db_to_height(ghost_db[idx], h - 20, db_range=80.0)
                    cv.coords(g, x, h - 10, x, h - 10 - bar_h)
                    cv.itemconfigure(g, state="normal")
                else:
                    cv.itemconfigure(g, state="hidden")
            else:
                cv.itemconfigure(g, state="hidden")

    # ------------------------------------------------------------------
    # CONTROLS
    # ------------------------------------------------------------------

    def _toner_on_sax_type_changed(self, event=None):
        if self._toner_engine:
            self._toner_engine.set_sax_type(self._toner_sax_var.get())

    def _toner_transpose_note(self, concert_note):
        """Transpose a concert pitch note name to written pitch for the selected sax.

        Returns the transposed note name, or the original if concert pitch is on.
        """
        if not concert_note or self._toner_concert_pitch.get():
            return concert_note

        sax_type = self._toner_sax_var.get()
        shift = SAX_TRANSPOSITIONS.get(sax_type, 0)
        if shift == 0:
            return concert_note

        # Parse note name: e.g. "C#4" -> pc_name="C#", octave=4
        if len(concert_note) >= 2 and concert_note[-1].isdigit():
            if len(concert_note) >= 3 and concert_note[-2].isdigit():
                # Shouldn't happen, but handle e.g. "C10"
                return concert_note
            if '#' in concert_note or 'b' in concert_note:
                pc_name = concert_note[:-1]
                octave = int(concert_note[-1])
            else:
                pc_name = concert_note[:-1]
                octave = int(concert_note[-1])
        else:
            return concert_note

        # Find pitch class index
        try:
            pc_idx = PITCH_CLASSES.index(pc_name)
        except ValueError:
            return concert_note

        # Apply transposition
        new_pc = (pc_idx + shift) % 12
        # Adjust octave when wrapping past B
        new_octave = octave + ((pc_idx + shift) // 12)

        return f"{PITCH_CLASSES[new_pc]}{new_octave}"

    def _toner_on_sensitivity_changed(self, value=None):
        if self._toner_engine:
            self._toner_engine.set_sensitivity(self._toner_sens_var.get())

    def _toner_on_pitch_changed(self):
        try:
            hz = float(self._toner_pitch_var.get())
            if self._toner_engine:
                self._toner_engine.set_reference_pitch(hz)
        except ValueError:
            pass

    def _toner_on_canvas_resize(self, event):
        if event.width > 50 and event.height > 50:
            self._toner_build_spectrum_bars()

    # ------------------------------------------------------------------
    # PROFILE MANAGEMENT
    # ------------------------------------------------------------------

    def _toner_open_profile_dialog(self):
        """Open the profile management dialog."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Tone Profiles")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Tone Profiles", bg=bg, fg=fg,
                 font=("Helvetica", 14, "bold")).pack(pady=(0, 10))

        # Profile list
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        self._prof_listbox = tk.Listbox(list_frame, width=45, height=10,
                                         font=("Helvetica", 10))
        self._prof_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, command=self._prof_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._prof_listbox.config(yscrollcommand=scrollbar.set)

        self._toner_refresh_profile_list()

        # Info display
        self._prof_info_label = tk.Label(frame, text="", bg=bg, fg=fg,
                                          font=("Helvetica", 9),
                                          justify="left", anchor="w")
        self._prof_info_label.pack(fill="x", pady=(0, 10))

        self._prof_listbox.bind("<<ListboxSelect>>",
                                lambda e: self._toner_on_profile_selected())

        # Buttons
        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")

        tk.Button(btn_frame, text="New Profile...",
                  command=lambda: self._toner_new_profile(dlg)).pack(
                      side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Edit Notes...",
                  command=self._toner_edit_profile_notes).pack(
                      side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Delete",
                  command=self._toner_delete_profile).pack(
                      side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Close",
                  command=dlg.destroy).pack(side="right")

    def _toner_refresh_profile_list(self):
        """Refresh the profile listbox (nested library structure)."""
        if not hasattr(self, '_prof_listbox'):
            return
        self._prof_listbox.delete(0, tk.END)
        self._prof_list_keys = []  # Track (lib, name) for selection lookup
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict) or not lib_profiles:
                continue
            self._prof_listbox.insert(tk.END, f"\u2500\u2500 {lib_name} \u2500\u2500")
            self._prof_list_keys.append(None)  # Header, not selectable
            for prof_name, profile in lib_profiles.items():
                sessions = profile.get('sessions', [])
                total_caps = sum(len(s.get('captures', []))
                               for s in sessions)
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                status = f"{len(notes)} notes" if total_caps > 0 else "empty"
                if len(notes) >= MIN_PROFILE_NOTES:
                    status += " \u2713"
                self._prof_listbox.insert(tk.END,
                    f"  {prof_name}  ({profile.get('horn_type', '?')}, {status})")
                self._prof_list_keys.append((lib_name, prof_name))

    def _toner_on_profile_selected(self):
        """Update info label when a profile is selected."""
        sel = self._prof_listbox.curselection()
        if not sel:
            return
        key = self._prof_list_keys[sel[0]]
        if key is None:
            self._prof_info_label.configure(text="")
            return  # Library header
        lib_name, prof_name = key
        profile = self._toner_profiles[lib_name][prof_name]
        self._prof_info_label.configure(text=_format_profile_info(profile))

    def _toner_edit_profile_notes(self):
        """Edit the notes field of the selected profile."""
        sel = self._prof_listbox.curselection()
        if not sel:
            return
        key = self._prof_list_keys[sel[0]]
        if key is None:
            return
        lib_name, prof_name = key
        profile = self._toner_profiles[lib_name][prof_name]

        current = profile.get('notes', '')
        new_notes = simpledialog.askstring(
            "Edit Notes",
            f"Notes for: {prof_name}",
            initialvalue=current,
            parent=self.root)
        if new_notes is not None:  # None = cancelled, "" = cleared intentionally
            profile['notes'] = new_notes
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            self._toner_on_profile_selected()  # Refresh info display

    def _toner_new_profile(self, parent_dlg):
        """Create a new horn profile via guided dialog."""
        dlg = tk.Toplevel(parent_dlg)
        dlg.title("New Tone Profile")
        dlg.resizable(False, False)
        dlg.transient(parent_dlg)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Profile", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 10))

        fields = {}

        def add_field(label, key, default="", widget_type="entry"):
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg=bg, fg=fg, width=14,
                     anchor="e", font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
            if widget_type == "combo":
                var = tk.StringVar(value=default)
                w = ttk.Combobox(row, textvariable=var, values=SAX_TYPES,
                                 state="readonly", width=20)
                w.pack(side="left", fill="x", expand=True)
                fields[key] = var
            else:
                var = tk.StringVar(value=default)
                tk.Entry(row, textvariable=var, width=25).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var

        add_field("Profile Name:", "name")
        add_field("Horn Type:", "horn_type", "Alto", widget_type="combo")
        add_field("Make:", "horn_make")
        add_field("Model:", "horn_model")
        add_field("Serial #:", "serial")
        add_field("Player:", "player")
        add_field("Mouthpiece:", "mouthpiece")
        add_field("Reed:", "reed")
        add_field("Notes:", "notes")

        def save():
            name = fields["name"].get().strip()
            if not name:
                messagebox.showwarning("Name Required",
                    "Please enter a profile name.", parent=dlg)
                return

            # Save into the default library
            lib = DEFAULT_LIBRARY
            if lib not in self._toner_profiles:
                self._toner_profiles[lib] = {}
            if name in self._toner_profiles[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            self._toner_profiles[lib][name] = {
                'horn_type': fields["horn_type"].get(),
                'horn_make': fields["horn_make"].get().strip(),
                'horn_model': fields["horn_model"].get().strip(),
                'serial': fields["serial"].get().strip(),
                'player': fields["player"].get().strip(),
                'mouthpiece': fields["mouthpiece"].get().strip(),
                'reed': fields["reed"].get().strip(),
                'notes': fields["notes"].get().strip(),
                'created': time.strftime("%Y-%m-%d"),
                'sessions': [],
            }
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            self._toner_active_library = lib
            self._toner_active_profile = name
            self._toner_refresh_profile_list()
            dlg.destroy()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Create", command=save).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="left")

    def _toner_delete_profile(self):
        """Delete the selected profile."""
        sel = self._prof_listbox.curselection()
        if not sel:
            return
        key = self._prof_list_keys[sel[0]]
        if key is None:
            return  # Library header
        lib_name, prof_name = key
        if messagebox.askyesno("Delete Profile",
                f"Delete profile '{prof_name}' from '{lib_name}'?"):
            del self._toner_profiles[lib_name][prof_name]
            # Remove empty libraries
            if not self._toner_profiles[lib_name]:
                del self._toner_profiles[lib_name]
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            if (self._toner_active_library == lib_name and
                    self._toner_active_profile == prof_name):
                self._toner_active_library = None
                self._toner_active_profile = None
                self._toner_active_session = None
            self._toner_refresh_profile_list()
            self._prof_info_label.configure(text="")

    # ------------------------------------------------------------------
    # CAPTURE SYSTEM (auto-detect stable tones)
    # ------------------------------------------------------------------

    def _toner_toggle_capture(self):
        """Toggle capture mode on/off. Guides user through setup if needed."""
        if self._toner_capture_state is not None:
            # Stop capturing
            self._toner_stop_capture()
            return

        if not self._toner_active_session:
            # No active session — walk user through setup
            self._toner_capture_setup_flow()
            return

        # Session exists — start listening
        self._toner_begin_listening()

    def _toner_begin_listening(self):
        """Enter listening mode (called after session is ready)."""
        self._toner_capture_state = 'listening'
        self._toner_stable_note = ""
        self._toner_stable_count = 0
        self._toner_capture_btn.configure(text="Stop")

        self._toner_capture_frame.pack(fill="x", padx=5,
            before=self._toner_main_frame.winfo_children()[-1])
        self._toner_capture_label.configure(
            text="Listening... play a steady note")
        self._toner_capture_progress.configure(text="")

    def _toner_capture_setup_flow(self):
        """Guide user to load or create a profile, then start capturing."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Start Capture")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        # Build flat list of (library, profile_name, profile_data)
        all_profiles = []
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict):
                continue
            for prof_name, prof_data in lib_profiles.items():
                all_profiles.append((lib_name, prof_name, prof_data))

        if all_profiles:
            tk.Label(frame, text="Load an existing profile or create a new one.",
                     bg=bg, fg=fg, font=("Helvetica", 10)).pack(pady=(0, 10))

            list_frame = tk.Frame(frame, bg=bg)
            list_frame.pack(fill="both", expand=True, pady=(0, 5))

            listbox = tk.Listbox(list_frame, width=50, height=10,
                                  font=("Helvetica", 10))
            listbox.pack(side="left", fill="both", expand=True)
            scrollbar = tk.Scrollbar(list_frame, command=listbox.yview)
            scrollbar.pack(side="right", fill="y")
            listbox.config(yscrollcommand=scrollbar.set)

            for lib_name, prof_name, prof in all_profiles:
                sessions = prof.get('sessions', [])
                total_caps = sum(len(s.get('captures', [])) for s in sessions)
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                status = f"{len(notes)} notes" if total_caps > 0 else "empty"
                if len(notes) >= MIN_PROFILE_NOTES:
                    status += " \u2713"
                listbox.insert(tk.END,
                    f"[{lib_name}] {prof_name}  ({status})")

            info_label = tk.Label(frame, text="", bg=bg, fg=fg,
                                   font=("Helvetica", 9),
                                   justify="left", anchor="w", wraplength=380)
            info_label.pack(fill="x", pady=(0, 8))

            def on_select(event=None):
                sel = listbox.curselection()
                if not sel:
                    return
                _, _, p = all_profiles[sel[0]]
                info = _format_profile_info(p)
                info_label.configure(text=info)

            listbox.bind("<<ListboxSelect>>", on_select)

            def load_profile():
                sel = listbox.curselection()
                if not sel:
                    messagebox.showinfo("Select Profile",
                        "Select a profile from the list.", parent=dlg)
                    return
                lib_name, prof_name, _ = all_profiles[sel[0]]
                self._toner_active_library = lib_name
                self._toner_active_profile = prof_name
                dlg.destroy()
                self._toner_start_new_session_and_listen()

            btn_frame = tk.Frame(frame, bg=bg)
            btn_frame.pack(fill="x", pady=(5, 0))
            tk.Button(btn_frame, text="Load && Capture",
                      command=load_profile).pack(side="left", padx=(0, 5))
            tk.Button(btn_frame, text="New Profile...",
                      command=lambda: [dlg.destroy(),
                                       self._toner_new_profile_flow()]).pack(
                          side="left", padx=(0, 5))
            tk.Button(btn_frame, text="Cancel",
                      command=dlg.destroy).pack(side="right")
        else:
            tk.Label(frame, text="No tone profiles yet.\n\n"
                     "Create a profile for the horn you want to analyze.",
                     bg=bg, fg=fg, font=("Helvetica", 10),
                     justify="left").pack(pady=(0, 10))
            btn_frame = tk.Frame(frame, bg=bg)
            btn_frame.pack(fill="x", pady=(5, 0))
            tk.Button(btn_frame, text="Create Profile...",
                      command=lambda: [dlg.destroy(),
                                       self._toner_new_profile_flow()]).pack(
                          side="left", padx=(0, 5))
            tk.Button(btn_frame, text="Cancel",
                      command=dlg.destroy).pack(side="right")

    def _toner_new_profile_flow(self):
        """Create a new profile (complete setup identity), then start capturing."""
        dlg = tk.Toplevel(self.root)
        dlg.title("New Tone Profile")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Profile", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text="A profile is a unique setup: horn + player + "
                 "mouthpiece + reed.\nChange any variable? That's a new profile.",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(pady=(0, 10))

        fields = {}

        def add_field(label, key, default="", widget_type="entry"):
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg=bg, fg=fg, width=14,
                     anchor="e", font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
            if widget_type == "combo":
                var = tk.StringVar(value=default)
                ttk.Combobox(row, textvariable=var, values=SAX_TYPES,
                             state="readonly", width=20).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var
            else:
                var = tk.StringVar(value=default)
                tk.Entry(row, textvariable=var, width=25).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var

        # Library selector
        lib_row = tk.Frame(frame, bg=bg)
        lib_row.pack(fill="x", pady=2)
        tk.Label(lib_row, text="Library:", bg=bg, fg=fg, width=14,
                 anchor="e", font=("Helvetica", 10)).pack(side="left", padx=(0, 8))
        existing_libs = [k for k in self._toner_profiles.keys()
                        if isinstance(self._toner_profiles[k], dict)]
        if not existing_libs:
            existing_libs = [DEFAULT_LIBRARY]
        lib_var = tk.StringVar(value=existing_libs[0])
        lib_combo = ttk.Combobox(lib_row, textvariable=lib_var,
                                  values=existing_libs, width=20)
        lib_combo.pack(side="left", fill="x", expand=True)

        add_field("Profile Name:", "name")
        add_field("Horn Type:", "horn_type", "Alto", widget_type="combo")
        add_field("Make:", "horn_make")
        add_field("Model:", "horn_model")
        add_field("Serial #:", "serial")
        add_field("Player:", "player")
        add_field("Mouthpiece:", "mouthpiece")
        add_field("Reed:", "reed")
        add_field("Notes:", "notes")

        def save_and_start():
            name = fields["name"].get().strip()
            lib = lib_var.get().strip()
            if not name:
                messagebox.showwarning("Name Required",
                    "Please enter a profile name.", parent=dlg)
                return
            if not lib:
                lib = DEFAULT_LIBRARY

            # Ensure library exists
            if lib not in self._toner_profiles:
                self._toner_profiles[lib] = {}

            if name in self._toner_profiles[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            self._toner_profiles[lib][name] = {
                'horn_type': fields["horn_type"].get(),
                'horn_make': fields["horn_make"].get().strip(),
                'horn_model': fields["horn_model"].get().strip(),
                'serial': fields["serial"].get().strip(),
                'player': fields["player"].get().strip(),
                'mouthpiece': fields["mouthpiece"].get().strip(),
                'reed': fields["reed"].get().strip(),
                'notes': fields["notes"].get().strip(),
                'created': time.strftime("%Y-%m-%d"),
                'sessions': [],
            }
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            self._toner_active_library = lib
            self._toner_active_profile = name
            dlg.destroy()
            self._toner_start_new_session_and_listen()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Create && Start Capturing",
                  command=save_and_start).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel",
                  command=dlg.destroy).pack(side="left")

    def _toner_start_new_session_and_listen(self):
        """Create a new session for the active profile and begin listening."""
        # Sync sax type from profile to selector and engine
        if self._toner_active_library and self._toner_active_profile:
            lib = self._toner_profiles.get(self._toner_active_library, {})
            prof = lib.get(self._toner_active_profile, {})
            sax_type = prof.get('horn_type', '')
            if sax_type:
                if hasattr(self, '_toner_sax_var'):
                    self._toner_sax_var.set(sax_type)
                if self._toner_engine:
                    self._toner_engine.set_sax_type(sax_type)

        self._toner_active_session = {
            'date': time.strftime("%Y-%m-%d %H:%M"),
            'captures': [],
        }
        self._toner_begin_listening()

    def _toner_stop_capture(self):
        """Stop capture mode entirely."""
        self._toner_capture_state = None
        self._toner_capture_frames = []
        self._toner_stable_note = ""
        self._toner_stable_count = 0
        self._toner_capture_frame.pack_forget()
        if hasattr(self, '_toner_capture_btn'):
            self._toner_capture_btn.configure(text="Capture")

    def _toner_cancel_capture(self):
        """Cancel button handler."""
        self._toner_stop_capture()

    def _toner_process_capture_frame(self, result):
        """Called each animation frame. Drives the auto-capture state machine."""
        state = self._toner_capture_state
        if state is None:
            return

        note = result.fundamental_note if result.fundamental_freq > 0 else ""

        # --- LISTENING: detect stable tone ---
        if state == 'listening':
            if note and note == self._toner_stable_note:
                self._toner_stable_count += 1
            elif note:
                self._toner_stable_note = note
                self._toner_stable_count = 1
            else:
                self._toner_stable_note = ""
                self._toner_stable_count = 0

            if self._toner_stable_count > 0:
                # Show what we're hearing
                needed = self._toner_stable_threshold - self._toner_stable_count
                if needed > 0:
                    self._toner_capture_label.configure(
                        text=f"Hearing {self._toner_stable_note}... hold steady")
                else:
                    self._toner_capture_label.configure(
                        text=f"Locking {self._toner_stable_note}...")
            else:
                self._toner_capture_label.configure(
                    text="Listening... play a steady note")

            # Trigger when stable long enough
            if self._toner_stable_count >= self._toner_stable_threshold:
                self._toner_capture_state = 'delay'
                self._toner_capture_start = time.time()
                self._toner_capture_note = self._toner_stable_note
                self._toner_capture_label.configure(
                    text=f"Settling {self._toner_capture_note}...")
            return

        # --- DELAY: settling period ---
        if state == 'delay':
            elapsed = time.time() - self._toner_capture_start
            remaining = CAPTURE_DELAY_S - elapsed

            # If the player stopped or changed notes during delay, abort back to listening
            if not note or note != self._toner_capture_note:
                self._toner_capture_state = 'listening'
                self._toner_stable_note = ""
                self._toner_stable_count = 0
                self._toner_capture_label.configure(
                    text="Listening... play a steady note")
                return

            if remaining > 0:
                self._toner_capture_label.configure(
                    text=f"Settling {self._toner_capture_note}... {remaining:.1f}s")
                return

            # Start recording
            self._toner_capture_state = 'recording'
            self._toner_capture_start = time.time()
            self._toner_capture_frames = []
            return

        # --- RECORDING: collecting frames ---
        if state == 'recording':
            elapsed = time.time() - self._toner_capture_start

            if result.fundamental_freq > 0 and result.harmonics:
                self._toner_capture_frames.append({
                    'note': result.fundamental_note,
                    'freq': result.fundamental_freq,
                    'harmonics_db': [h.magnitude_db for h in result.harmonics],
                    'descriptors': dict(result.descriptors),
                })

            remaining = CAPTURE_DURATION_S - elapsed
            n_frames = len(self._toner_capture_frames)
            self._toner_capture_label.configure(
                text=f"Recording {self._toner_capture_note}... {remaining:.1f}s")
            self._toner_capture_progress.configure(
                text=f"({n_frames} frames)")

            if remaining <= 0:
                self._toner_finish_capture()
            return

        # --- COOLDOWN: brief pause before next auto-capture ---
        if state == 'cooldown':
            elapsed = time.time() - self._toner_capture_start
            # Wait for note change or silence before listening again
            if not note or note != self._toner_capture_note or elapsed > 2.0:
                self._toner_capture_state = 'listening'
                self._toner_stable_note = ""
                self._toner_stable_count = 0
                self._toner_capture_label.configure(
                    text="Listening... play the next note")
                self._toner_capture_progress.configure(text="")
            else:
                self._toner_capture_label.configure(
                    text=f"Captured {self._toner_capture_note} \u2713  change to next note...")
            return

    def _toner_finish_capture(self):
        """Finish a capture — average frames, save, and go to cooldown."""
        frames = self._toner_capture_frames

        if not frames or len(frames) < 3:
            # Not enough data — go back to listening silently
            self._toner_capture_state = 'listening'
            self._toner_stable_note = ""
            self._toner_stable_count = 0
            self._toner_capture_label.configure(
                text="Capture skipped (unstable). Play the next note...")
            return

        # Find the most common note
        from collections import Counter
        note_counts = Counter(f['note'] for f in frames)
        dominant_note = note_counts.most_common(1)[0][0]

        # Filter to only frames matching the dominant note
        note_frames = [f for f in frames if f['note'] == dominant_note]
        if len(note_frames) < 3:
            self._toner_capture_state = 'listening'
            self._toner_stable_note = ""
            self._toner_stable_count = 0
            self._toner_capture_label.configure(
                text="Capture skipped (inconsistent). Play the next note...")
            return

        # Average
        averaged = average_captures(note_frames)
        avg_freq = sum(f['freq'] for f in note_frames) / len(note_frames)

        capture_entry = {
            'note': dominant_note,
            'fundamental_freq': round(avg_freq, 2),
            'harmonics_db': [round(db, 2) for db in averaged['harmonics_db']],
            'descriptors': {k: round(v, 3) for k, v in averaged['descriptors'].items()},
            'timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
            'n_frames': len(note_frames),
        }

        # Save to active session
        self._toner_active_session['captures'].append(capture_entry)

        # Save session to profile (nested library structure)
        lib = self._toner_active_library
        prof_name = self._toner_active_profile
        if lib and prof_name and lib in self._toner_profiles:
            lib_profiles = self._toner_profiles[lib]
            if prof_name in lib_profiles:
                profile = lib_profiles[prof_name]
                sessions = profile.setdefault('sessions', [])
                session_date = self._toner_active_session['date']
                found = False
                for s in sessions:
                    if s.get('date') == session_date:
                        s['captures'] = self._toner_active_session['captures']
                        found = True
                        break
                if not found:
                    sessions.append(self._toner_active_session)

                save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)

        # Count progress (non-blocking — just update the status bar)
        notes_so_far = set()
        for cap in self._toner_active_session['captures']:
            notes_so_far.add(cap['note'])

        n_notes = len(notes_so_far)
        desc = averaged['descriptors']
        status = f"Captured {dominant_note} \u2713  ({n_notes} notes"
        if n_notes < MIN_PROFILE_NOTES:
            status += f", need {MIN_PROFILE_NOTES - n_notes} more"
        else:
            status += ", fingerprint ready!"
        status += ")"

        self._toner_capture_label.configure(text=status)
        self._toner_capture_progress.configure(
            text=f"R:{desc.get('resonance',0):.0%} "
                 f"Rich:{desc.get('richness',0):.0%} "
                 f"Br:{desc.get('brightness',0):.0%} "
                 f"Dk:{desc.get('darkness',0):.0%}")

        # Enter cooldown — wait for note change before next capture
        self._toner_capture_state = 'cooldown'
        self._toner_capture_start = time.time()

    # ------------------------------------------------------------------
    # COMPARISON
    # ------------------------------------------------------------------

    def _toner_open_compare_dialog(self):
        """Open dialog to select a profile for comparison overlay."""
        all_profiles = []
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict):
                continue
            for prof_name, prof_data in lib_profiles.items():
                sessions = prof_data.get('sessions', [])
                total_caps = sum(len(s.get('captures', [])) for s in sessions)
                if total_caps > 0:
                    all_profiles.append((lib_name, prof_name, prof_data))

        if not all_profiles:
            messagebox.showinfo("No Profiles",
                "No profiles with captures to compare.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Compare Tone Profile")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select a profile to overlay on the spectrum.\n"
                 "The ghost shows the saved harmonic profile in blue.",
                 bg=bg, fg=fg, font=("Helvetica", 10),
                 justify="left").pack(pady=(0, 5))

        # Filter controls
        filter_frame = tk.Frame(frame, bg=bg)
        filter_frame.pack(fill="x", pady=(0, 5))

        # Collect unique values for filters
        all_types = sorted(set(p.get('horn_type', '') for _, _, p in all_profiles if p.get('horn_type')))
        all_players = sorted(set(p.get('player', '') for _, _, p in all_profiles if p.get('player')))
        all_mpcs = sorted(set(p.get('mouthpiece', '') for _, _, p in all_profiles if p.get('mouthpiece')))

        filter_type = tk.StringVar(value="All")
        filter_player = tk.StringVar(value="All")
        filter_mpc = tk.StringVar(value="All")

        for label, var, values in [
            ("Type:", filter_type, all_types),
            ("Player:", filter_player, all_players),
            ("Mpc:", filter_mpc, all_mpcs),
        ]:
            if values:
                tk.Label(filter_frame, text=label, bg=bg, fg=fg,
                         font=("Helvetica", 8)).pack(side="left", padx=(0, 2))
                cb = ttk.Combobox(filter_frame, textvariable=var,
                                   values=["All"] + values,
                                   state="readonly", width=10)
                cb.pack(side="left", padx=(0, 6))
                cb.bind("<<ComboboxSelected>>", lambda e: refresh_list())

        # Multi-select listbox
        listbox = tk.Listbox(frame, width=50, height=10,
                              font=("Helvetica", 10),
                              selectmode=tk.EXTENDED)
        listbox.pack(fill="both", expand=True, pady=(0, 10))

        filtered_profiles = []

        def refresh_list():
            nonlocal filtered_profiles
            listbox.delete(0, tk.END)
            filtered_profiles = []
            ft = filter_type.get()
            fp = filter_player.get()
            fm = filter_mpc.get()
            for lib_name, prof_name, prof in all_profiles:
                if ft != "All" and prof.get('horn_type', '') != ft:
                    continue
                if fp != "All" and prof.get('player', '') != fp:
                    continue
                if fm != "All" and prof.get('mouthpiece', '') != fm:
                    continue
                filtered_profiles.append((lib_name, prof_name, prof))
                sessions = prof.get('sessions', [])
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                status = f"{len(notes)} notes"
                if len(notes) >= MIN_PROFILE_NOTES:
                    status += " \u2713"
                listbox.insert(tk.END,
                    f"[{lib_name}] {prof_name}  ({status})")

        refresh_list()

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")

        def get_selected():
            """Return list of (name, fingerprint) for selected profiles."""
            sel = listbox.curselection()
            if not sel or not filtered_profiles:
                return []
            results = []
            for idx in sel:
                lib_name, prof_name, prof = filtered_profiles[idx]
                fp = compute_fingerprint(prof.get('sessions', []))
                fp['_name'] = prof_name
                fp['_profile'] = prof
                results.append(fp)
            return results

        def load_overlay():
            """Load first selected profile as live ghost overlay."""
            selected = get_selected()
            if not selected:
                return
            self._toner_comparison = selected[0]
            name = selected[0]['_name']
            if hasattr(self, '_toner_compare_label'):
                self._toner_spectrum_canvas.itemconfigure(
                    self._toner_compare_label,
                    text=f"Overlay: {name} ({selected[0]['note_count']} notes)")
            dlg.destroy()

        def compare_selected():
            """Open multi-profile comparison analysis window."""
            selected = get_selected()
            if len(selected) < 2:
                messagebox.showinfo("Select More",
                    "Select at least 2 profiles to compare.", parent=dlg)
                return
            dlg.destroy()
            self._toner_show_comparison_analysis(selected)

        def clear_comparison():
            self._toner_comparison = None
            if hasattr(self, '_toner_compare_label'):
                self._toner_spectrum_canvas.itemconfigure(
                    self._toner_compare_label, text="")
            for g in self._toner_ghost_markers:
                self._toner_spectrum_canvas.itemconfigure(g, state="hidden")
            dlg.destroy()

        tk.Button(btn_frame, text="Compare Selected",
                  command=compare_selected).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Overlay on Spectrum",
                  command=load_overlay).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Clear",
                  command=clear_comparison).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel",
                  command=dlg.destroy).pack(side="right")

    def _toner_show_comparison_analysis(self, fingerprints):
        """Show a multi-profile comparison window with chart and analysis.

        Supports both horn-average and per-note views.
        """
        dlg = tk.Toplevel(self.root)
        dlg.title("Tone Profile Comparison")
        dlg.geometry("720x600")
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"

        chart_colors = ["#2196F3", "#FF5722", "#4CAF50", "#FF9800",
                        "#9C27B0", "#00BCD4", "#E91E63", "#8BC34A"]

        main = tk.Frame(dlg, bg=bg)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # --- View toggle: Horn Average vs Per-Note ---
        toggle_frame = tk.Frame(main, bg=bg)
        toggle_frame.pack(fill="x", pady=(0, 5))

        view_mode = tk.StringVar(value="average")
        tk.Radiobutton(toggle_frame, text="Horn Average", variable=view_mode,
                        value="average", bg=bg, fg=fg, selectcolor=bg,
                        font=("Helvetica", 10),
                        command=lambda: refresh_all()).pack(side="left", padx=(0, 10))
        tk.Radiobutton(toggle_frame, text="Per-Note", variable=view_mode,
                        value="per_note", bg=bg, fg=fg, selectcolor=bg,
                        font=("Helvetica", 10),
                        command=lambda: refresh_all()).pack(side="left", padx=(0, 10))

        # Note selector (shown only in per-note mode)
        note_frame = tk.Frame(toggle_frame, bg=bg)
        note_frame.pack(side="left", padx=(10, 0))
        tk.Label(note_frame, text="Note:", bg=bg, fg=fg,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 4))

        # Collect all notes across all profiles
        all_notes = set()
        for fp in fingerprints:
            all_notes.update(fp.get('per_note', {}).keys())
        # Sort notes chromatically
        note_order = []
        for octave in range(1, 8):
            for pc in PITCH_CLASSES:
                n = f"{pc}{octave}"
                if n in all_notes:
                    note_order.append(n)

        note_var = tk.StringVar(value=note_order[0] if note_order else "")
        note_combo = ttk.Combobox(note_frame, textvariable=note_var,
                                   values=note_order, state="readonly", width=5)
        note_combo.pack(side="left")
        note_combo.bind("<<ComboboxSelected>>", lambda e: refresh_all())

        # --- Harmonic chart ---
        chart_frame = tk.LabelFrame(main, text="Harmonic Profiles", bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        chart_frame.pack(fill="both", expand=True, pady=(0, 6))

        chart_cv = tk.Canvas(chart_frame, bg="white", highlightthickness=0,
                              height=180)
        chart_cv.pack(fill="both", expand=True, padx=5, pady=5)

        legend_frame = tk.Frame(chart_frame, bg=bg)
        legend_frame.pack(fill="x", padx=5, pady=(0, 3))
        for i, fp in enumerate(fingerprints):
            color = chart_colors[i % len(chart_colors)]
            tk.Label(legend_frame, text="\u25a0", fg=color, bg=bg,
                     font=("Helvetica", 12)).pack(side="left")
            tk.Label(legend_frame, text=fp['_name'], bg=bg, fg=fg,
                     font=("Helvetica", 9)).pack(side="left", padx=(0, 12))

        def get_data_for_view():
            """Return list of (harmonics_db, descriptors) per fingerprint for current view."""
            mode = view_mode.get()
            result = []
            for fp in fingerprints:
                if mode == "per_note":
                    note = note_var.get()
                    per_note = fp.get('per_note', {})
                    if note in per_note:
                        result.append(per_note[note])
                    else:
                        result.append({'harmonics_db': [], 'descriptors': {}})
                else:
                    result.append({
                        'harmonics_db': fp.get('harmonics_db', []),
                        'descriptors': fp.get('descriptors', {}),
                    })
            return result

        def draw_chart(event=None):
            chart_cv.delete("all")
            w = chart_cv.winfo_width()
            h = chart_cv.winfo_height()
            if w < 50 or h < 50:
                return

            data = get_data_for_view()

            margin_l, margin_r, margin_t, margin_b = 40, 10, 10, 25
            cw = w - margin_l - margin_r
            ch = h - margin_t - margin_b
            db_min, db_max = -60.0, 0.0

            all_db = [d.get('harmonics_db', []) for d in data]
            max_h = max((len(db) for db in all_db), default=2)
            max_h = max(max_h, 2)

            # Grid
            for db in range(-60, 1, 10):
                y = margin_t + ch * (1.0 - (db - db_min) / (db_max - db_min))
                chart_cv.create_line(margin_l, y, w - margin_r, y,
                                      fill="#DDDDDD", width=1)
                chart_cv.create_text(margin_l - 4, y, text=f"{db}",
                                      fill="#888888", font=("Helvetica", 7),
                                      anchor="e")

            for hi in range(max_h):
                x = margin_l + cw * hi / (max_h - 1) if max_h > 1 else margin_l
                chart_cv.create_text(x, h - 5, text=f"{hi + 1}",
                                      fill="#888888", font=("Helvetica", 7))

            chart_cv.create_text(w // 2, h - 2, text="Harmonic #",
                                  fill="#888888", font=("Helvetica", 7))

            # "No data" message for per-note when a note is missing
            mode = view_mode.get()
            if mode == "per_note":
                note = note_var.get()
                missing = [fingerprints[i]['_name'] for i, d in enumerate(data)
                          if not d.get('harmonics_db')]
                if missing:
                    chart_cv.create_text(w // 2, margin_t + 20,
                        text=f"No data for {note}: {', '.join(missing)}",
                        fill="#CC0000", font=("Helvetica", 9))

            for i, d in enumerate(data):
                color = chart_colors[i % len(chart_colors)]
                db_list = d.get('harmonics_db', [])
                if len(db_list) < 2:
                    continue

                points = []
                for hi, db in enumerate(db_list):
                    x = margin_l + cw * hi / (max_h - 1) if max_h > 1 else margin_l
                    clamped = max(db_min, min(db_max, db))
                    y = margin_t + ch * (1.0 - (clamped - db_min) / (db_max - db_min))
                    points.extend([x, y])

                if len(points) >= 4:
                    chart_cv.create_line(*points, fill=color, width=2, smooth=True)
                    for j in range(0, len(points), 2):
                        chart_cv.create_oval(
                            points[j] - 3, points[j + 1] - 3,
                            points[j] + 3, points[j + 1] + 3,
                            fill=color, outline="")

        chart_cv.bind("<Configure>", draw_chart)

        # --- Descriptor table (rebuilt on view change) ---
        table_frame = tk.LabelFrame(main, text="Descriptor Comparison", bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        table_frame.pack(fill="x", pady=(0, 6))
        table_inner = tk.Frame(table_frame, bg=bg)
        table_inner.pack(fill="x", padx=5, pady=5)

        desc_labels = [
            ("Resonance", "resonance"),
            ("Richness", "richness"),
            ("Brightness", "brightness"),
            ("Darkness", "darkness"),
            ("Fullness", "fullness"),
        ]

        # --- Analysis text ---
        analysis_frame = tk.LabelFrame(main, text="Analysis", bg=bg, fg=fg,
                                        font=("Helvetica", 10, "bold"))
        analysis_frame.pack(fill="x")
        analysis_text = tk.Text(analysis_frame, height=4, wrap="word",
                                 font=("Helvetica", 9), bg="white",
                                 relief="flat", padx=8, pady=5)
        analysis_text.pack(fill="x", padx=5, pady=5)

        def rebuild_table():
            for w in table_inner.winfo_children():
                w.destroy()

            data = get_data_for_view()

            header = tk.Frame(table_inner, bg=bg)
            header.pack(fill="x")
            tk.Label(header, text="", width=14, bg=bg, fg=fg,
                     font=("Helvetica", 9, "bold"), anchor="w").pack(side="left")
            for fp in fingerprints:
                tk.Label(header, text=fp['_name'][:15], width=12, bg=bg, fg=fg,
                         font=("Helvetica", 9, "bold"), anchor="center").pack(side="left")

            for label, key in desc_labels:
                row = tk.Frame(table_inner, bg=bg)
                row.pack(fill="x")
                tk.Label(row, text=label, width=14, bg=bg, fg=fg,
                         font=("Helvetica", 9), anchor="w").pack(side="left")

                values = [d.get('descriptors', {}).get(key, 0) for d in data]
                max_val = max(values) if values else 0
                min_val = min(values) if values else 0

                for val in values:
                    if len(values) > 1 and max_val != min_val:
                        if val == max_val:
                            val_fg = "#006600"
                        elif val == min_val:
                            val_fg = "#880000"
                        else:
                            val_fg = fg
                    else:
                        val_fg = fg
                    text = f"{val:.0%}" if val > 0 else "\u2014"
                    tk.Label(row, text=text, width=12, bg=bg, fg=val_fg,
                             font=("Helvetica", 9), anchor="center").pack(side="left")

        def rebuild_analysis():
            data = get_data_for_view()
            analysis_text.configure(state="normal")
            analysis_text.delete("1.0", tk.END)

            mode = view_mode.get()
            prefix = ""
            if mode == "per_note":
                prefix = f"For {note_var.get()}: "

            lines = []
            if len(fingerprints) == 2:
                da = data[0].get('descriptors', {})
                db_d = data[1].get('descriptors', {})
                if not da and not db_d:
                    lines.append(f"{prefix}No data for this note in either profile.")
                else:
                    for label, key in desc_labels:
                        va = da.get(key, 0)
                        vb = db_d.get(key, 0)
                        diff = va - vb
                        if abs(diff) > 0.05:
                            winner = fingerprints[0]['_name'] if diff > 0 else fingerprints[1]['_name']
                            lines.append(f"{prefix}{winner} is more {label.lower()} "
                                        f"({abs(diff):.0%} difference)")
                    if not lines:
                        lines.append(f"{prefix}These profiles are very similar.")
            else:
                for label, key in desc_labels:
                    values = [(fingerprints[i]['_name'],
                              d.get('descriptors', {}).get(key, 0))
                             for i, d in enumerate(data)]
                    values.sort(key=lambda x: x[1], reverse=True)
                    top_name, top_val = values[0]
                    bot_name, bot_val = values[-1]
                    if top_val - bot_val > 0.1:
                        lines.append(f"{prefix}{label}: {top_name} highest "
                                    f"({top_val:.0%}), {bot_name} lowest ({bot_val:.0%})")

            analysis_text.insert("1.0", "\n".join(lines) if lines else
                                 f"{prefix}Profiles are similar \u2014 no major differences.")
            analysis_text.configure(state="disabled")

        def refresh_all():
            draw_chart()
            rebuild_table()
            rebuild_analysis()

        # Initial build
        rebuild_table()
        rebuild_analysis()

        tk.Button(main, text="Close", command=dlg.destroy).pack(pady=(5, 0))

    # ------------------------------------------------------------------
    # ANIMATION LOOP
    # ------------------------------------------------------------------

    def _toner_start(self):
        """Start the toner (audio capture + animation)."""
        if not self._toner_engine:
            return

        if not self._toner_bars_built:
            self._toner_build_spectrum_bars()

        success, err = self._toner_engine.start()
        if not success:
            if hasattr(self, '_toner_spectrum_canvas'):
                self._toner_spectrum_canvas.create_text(
                    self._toner_spectrum_canvas.winfo_width() / 2,
                    self._toner_spectrum_canvas.winfo_height() / 2,
                    text=f"Audio error: {err}",
                    fill="#FF4444", font=("Helvetica", 12),
                    tags="error"
                )
            return

        self._toner_running = True
        self._toner_animate()

    def _toner_stop(self):
        """Stop the toner (audio + animation)."""
        self._toner_running = False
        if self._toner_anim_id is not None:
            try:
                self.root.after_cancel(self._toner_anim_id)
            except Exception:
                pass
            self._toner_anim_id = None

        if self._toner_engine:
            self._toner_engine.stop()

    def _toner_animate(self):
        """One animation frame."""
        if not self._toner_running:
            return

        if self._toner_engine and self._toner_engine.is_running:
            result = self._toner_engine.analyze()

            # Update spectrum
            self._toner_render_spectrum(result)

            # Update note display + intonation gauge
            if result.fundamental_freq > 0:
                display_note = self._toner_transpose_note(result.fundamental_note)
                self._toner_note_label.configure(
                    text=display_note, fg=LABEL_BRIGHT)
                self._toner_freq_label.configure(
                    text=f"{result.fundamental_freq:.1f} Hz")
                self._toner_update_intonation(result.fundamental_cents)
            else:
                self._toner_note_label.configure(text="\u2014", fg=LABEL_DIM)
                self._toner_freq_label.configure(text="")
                self._toner_update_intonation(0.0)

            # Update descriptor gauges
            d = result.descriptors
            self._toner_update_gauge('resonance', d['resonance'])
            self._toner_update_gauge('richness', d['richness'])
            self._toner_update_gauge('brightness', d['brightness'])
            self._toner_update_gauge('darkness', d['darkness'])

            # Fullness lamp — lights when bright > 0.5 AND dark > 0.6
            # (using the smoothed gauge values which include bias)
            smooth_bright = self._toner_smooth.get('brightness', 0)
            bright_bias = self._toner_bias_vars.get('brightness')
            dark_bias = self._toner_bias_vars.get('darkness')
            display_bright = smooth_bright + (bright_bias.get() / 100.0 if bright_bias else 0)
            smooth_dark = self._toner_smooth.get('darkness', 0)
            display_dark = smooth_dark + (dark_bias.get() / 100.0 if dark_bias else 0)
            full_on = display_bright > 0.5 and display_dark > 0.6

            if full_on:
                self._toner_full_canvas.itemconfigure(
                    self._toner_full_lamp, fill="#FF8800")
                self._toner_full_canvas.itemconfigure(
                    self._toner_full_glow, fill="#442200")
            else:
                self._toner_full_canvas.itemconfigure(
                    self._toner_full_lamp, fill="#331100")
                self._toner_full_canvas.itemconfigure(
                    self._toner_full_glow, fill="#1A0A00")

            # Process capture if active
            self._toner_process_capture_frame(result)

        interval = FRAME_RATES.get(self._toner_fps_var.get(), 33)
        self._toner_anim_id = self.root.after(interval, self._toner_animate)

    # ------------------------------------------------------------------
    # SETTINGS SAVE/RESTORE
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # IMPORT / EXPORT
    # ------------------------------------------------------------------

    def _toner_export_profiles(self):
        """Export tone profiles to a JSON file."""
        from tkinter import filedialog
        if not self._toner_profiles:
            messagebox.showinfo("Nothing to Export", "No tone profiles to export.")
            return

        filepath = filedialog.asksaveasfilename(
            title="Export Tone Profiles",
            defaultextension=".json",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialfile="tone_profiles_export.json"
        )
        if not filepath:
            return

        try:
            import json
            with open(filepath, 'w') as f:
                json.dump(self._toner_profiles, f, indent=2)
            messagebox.showinfo("Export Successful",
                f"Exported tone profiles to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Could not export:\n{e}")

    def _toner_import_profiles(self):
        """Import tone profiles from a JSON file, merging into existing libraries."""
        from tkinter import filedialog
        import json

        filepath = filedialog.askopenfilename(
            title="Import Tone Profiles",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
        )
        if not filepath:
            return

        try:
            with open(filepath, 'r') as f:
                imported = json.load(f)
        except Exception as e:
            messagebox.showerror("Import Error", f"Could not read file:\n{e}")
            return

        if not isinstance(imported, dict):
            messagebox.showerror("Invalid Format", "File is not a valid tone profiles export.")
            return

        # Check if flat (old format) or nested (library format)
        is_flat = any(isinstance(v, dict) and 'sessions' in v
                      for v in imported.values())

        if is_flat:
            # Wrap in a library
            imported = {"Imported": imported}

        # Merge
        count = 0
        for lib_name, lib_profiles in imported.items():
            if not isinstance(lib_profiles, dict):
                continue
            if lib_name not in self._toner_profiles:
                self._toner_profiles[lib_name] = {}
            for prof_name, prof_data in lib_profiles.items():
                if prof_name not in self._toner_profiles[lib_name]:
                    self._toner_profiles[lib_name][prof_name] = prof_data
                    count += 1
                else:
                    # Merge sessions into existing profile
                    existing = self._toner_profiles[lib_name][prof_name]
                    existing_dates = {s.get('date') for s in existing.get('sessions', [])}
                    for session in prof_data.get('sessions', []):
                        if session.get('date') not in existing_dates:
                            existing.setdefault('sessions', []).append(session)
                            count += 1

        save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
        messagebox.showinfo("Import Complete",
            f"Imported {count} new profiles/sessions.")

    def _toner_save_settings(self):
        """Save toner settings to the settings dict."""
        self.settings["toner_settings"] = {
            "reference_pitch": float(self._toner_pitch_var.get()) if hasattr(self, '_toner_pitch_var') else 440.0,
            "sensitivity": self._toner_sens_var.get() if hasattr(self, '_toner_sens_var') else 50,
            "fps": self._toner_fps_var.get() if hasattr(self, '_toner_fps_var') else "30",
            "view_mode": self._toner_view_var.get() if hasattr(self, '_toner_view_var') else "spectrum",
            "scale_mode": self._toner_scale_var.get() if hasattr(self, '_toner_scale_var') else "linear",
            "gauge_bias": {k: v.get() for k, v in self._toner_bias_vars.items()} if hasattr(self, '_toner_bias_vars') else {},
            "sax_type": self._toner_sax_var.get() if hasattr(self, '_toner_sax_var') else "Alto",
            "concert_pitch": self._toner_concert_pitch.get() if hasattr(self, '_toner_concert_pitch') else False,
        }
        # Also save any pending session
        if self._toner_active_session and self._toner_active_profile:
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
