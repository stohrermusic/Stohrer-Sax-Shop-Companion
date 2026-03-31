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
import os
import sys
import time

IS_MACOS = sys.platform == 'darwin'

try:
    from toner_engine import (
        TonerEngine, AUDIO_AVAILABLE, PITCH_CLASSES, MAX_HARMONICS,
        MIN_PROFILE_NOTES, SAX_TYPES, MIN_FUNDAMENTAL_HZ,
        SAX_TRANSPOSITIONS, FREE_STABLE_FRAMES, FREE_MIN_FRAMES,
        ATTACK_SKIP_FRAMES, CALIBRATION_NOTES, CALIBRATION_DURATION_S,
        DEFAULT_LIBRARY, average_captures, compute_fingerprint,
        compute_session_fingerprint, compute_session_variation,
        compute_group_fingerprint, compute_rolloff_rate,
        compute_delta_descriptors,
        ROLLOFF_WARN_THRESHOLD, ROLLOFF_MIN_CAPTURES,
        load_tone_profiles, save_tone_profiles,
        analyze_audio_file, check_mic_quality,
        transpose_note, reverse_transpose_note, note_to_freq,
    )
    _TONER_IMPORTS_OK = True
except ImportError:
    _TONER_IMPORTS_OK = False
    AUDIO_AVAILABLE = False
    PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

try:
    from config import TONE_PROFILES_FILE, save_settings
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

# Gauge arc geometry and dimensions
GAUGE_ARC_START = 155.0
GAUGE_ARC_END = 25.0
GAUGE_WIDTH = 240
GAUGE_HEIGHT = 100
GAUGE_ARC_RADIUS = 65
GAUGE_DESCRIPTOR_NEEDLE_OFFSET = 12   # needle_len = radius - offset
GAUGE_INTONATION_NEEDLE_OFFSET = 16
GAUGE_DESCRIPTOR_DAMPING = 0.15       # lerp factor per frame for descriptor needles
GAUGE_INTONATION_DAMPING = 0.18       # lerp factor per frame for intonation needle

# Animation / display thresholds
IN_TUNE_CENTS_THRESHOLD = 4.0         # cents — below this, "in tune" lamp lights
INTONATION_RANGE_CENTS = 50.0         # cents — full deflection of intonation gauge
SPECTRUM_MAX_FREQ = 8000.0            # Hz — right edge of spectrum display
SPECTRAL_CHECK_FRAME_COUNT = 25       # frames before mic quality check fires
LOW_DATA_FRAME_THRESHOLD = 15         # frames (~0.5s) before "low data" overlay shows
FRAME_DURATION_S = 0.033              # approximate duration of one frame at 30 fps

# Capture state machine
CAL_COUNTDOWN_S = 10.0                # seconds of countdown before calibration starts
CAL_ATTACK_SETTLE_S = 0.1             # seconds of attack transient to skip in calibration
DEFAULT_STABLE_THRESHOLD = 25         # frames — initial stable-note threshold
PROFILE_NAME_MAX_DISPLAY = 20         # characters before truncating profile name


def _note_sort_key(note_name):
    """Return a numeric sort key for a note name like 'C#4'.

    Returns MIDI-like number: C4=60, A4=69, etc.
    """
    if not note_name or len(note_name) < 2:
        return 0
    pc_order = {'C': 0, 'C#': 1, 'D': 2, 'D#': 3, 'E': 4, 'F': 5,
                'F#': 6, 'G': 7, 'G#': 8, 'A': 9, 'A#': 10, 'B': 11}
    try:
        if '#' in note_name:
            pc = note_name[:-1]
            octave = int(note_name[-1])
        else:
            pc = note_name[:-1]
            octave = int(note_name[-1])
        return (octave + 1) * 12 + pc_order.get(pc, 0)
    except (ValueError, KeyError):
        return 0


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
    total_caps = 0
    notes = set()
    method_counts = {}
    for s in sessions:
        for c in s.get('captures', []):
            total_caps += 1
            notes.add(c.get('note', ''))
            m = c.get('method', 'structured')
            method_counts[m] = method_counts.get(m, 0) + 1
    info += f"\n{len(sessions)} sessions, {total_caps} captures, {len(notes)} unique notes"
    if 0 < len(notes) < MIN_PROFILE_NOTES:
        info += f" (need {MIN_PROFILE_NOTES - len(notes)} more)"
    elif len(notes) >= MIN_PROFILE_NOTES:
        info += " \u2713"
    if method_counts:
        method_parts = []
        for m in ['structured', 'free', 'file']:
            if m in method_counts:
                method_parts.append(f"{method_counts[m]} {m}")
        info += f"\nCaptures by method: {', '.join(method_parts)}"
    if total_caps > 0:
        total_frames = sum(c.get('n_frames', 0)
                          for s in sessions for c in s.get('captures', []))
        if total_frames > 0:
            # Each frame is ~33ms at 30fps
            avg_frames = total_frames / total_caps
            avg_duration = avg_frames * FRAME_DURATION_S
            info += f"\nAvg capture: {avg_frames:.0f} frames ({avg_duration:.1f}s)"
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
            'richness': 0.0, 'warmth': 0.0,
        }
        # Low harmonic data tracking — grey out gauges when sustained
        self._toner_low_data_frames = 0
        self._toner_low_data_shown = False
        # Profile system state — nested: {library: {profile_name: data}}
        self._toner_profiles = {}
        self._toner_active_library = None   # Library name
        self._toner_active_profile = None   # Profile name within library
        self._toner_active_session = None   # Current capture session dict
        # Capture state machine: None, 'listening', 'delay', 'recording'
        self._toner_capture_state = None
        self._toner_capture_start = 0.0
        self._toner_capture_frames = []     # Accumulated frames during capture
        self._toner_stable_note = ""        # Note currently being held steady
        self._toner_stable_count = 0        # Consecutive frames of same note
        self._toner_comparison = None       # Fingerprint dict for ghost overlay
        self._toner_capture_mode = "free"  # "free" or "calibration"
        # Free: ~0.5s stability, continuous micro-captures, attack skip
        # Calibration: guided chromatic scale, 5s per note, 1s settle
        self._toner_paused = False         # Pause state
        self._toner_paused_state = None    # State to resume to
        self._toner_stable_threshold = DEFAULT_STABLE_THRESHOLD  # Updated per mode
        self._toner_free_accumulator = []  # For free mode: frames of current stable note
        self._toner_cal_index = 0          # Current note index in calibration sequence
        self._toner_cal_notes = list(CALIBRATION_NOTES)  # Filtered at calibration start
        self._toner_cal_recording = False  # Whether we're actively recording in calibration
        self._toner_mic_checked = False    # Whether mic availability has been checked

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
        gauge_outer = tk.Frame(top_frame, bg=bg)
        gauge_outer._skip_theme = True
        gauge_outer.grid(row=0, column=1, sticky="nsew")

        gauge_frame = tk.Frame(gauge_outer, bg=bg)
        gauge_frame._skip_theme = True
        gauge_frame.pack(expand=True)

        gauge_row = 0

        # --- Row 0: Intonation gauge + note display ---
        self._toner_intonation_canvas = tk.Canvas(
            gauge_frame, bg=bg, highlightthickness=0, width=240, height=100)
        self._toner_intonation_canvas._dark_canvas = True
        self._toner_intonation_canvas.grid(row=gauge_row, column=1,
                                            pady=(2, 0))
        self._toner_intonation_gauge = self._toner_build_intonation_gauge(
            self._toner_intonation_canvas)
        self._toner_smooth_cents = 0.0

        # Note display + in-tune lamp to the right of intonation gauge
        note_frame = tk.Frame(gauge_frame, bg=bg)
        note_frame._skip_theme = True
        note_frame.grid(row=gauge_row, column=2, sticky="ns", padx=(4, 2))

        self._toner_note_label = tk.Label(
            note_frame, text="\u2014", bg=bg, fg=LABEL_DIM,
            font=("Helvetica", 20, "bold"))
        self._toner_note_label.pack(pady=(2, 0))

        self._toner_freq_label = tk.Label(
            note_frame, text="", bg=bg, fg=LABEL_DIM,
            font=("Helvetica", 7))
        self._toner_freq_label.pack()

        # In-tune lamp (green when within 1 cent)
        intune_frame = tk.Frame(note_frame, bg=bg)
        intune_frame._skip_theme = True
        intune_frame.pack(pady=(4, 0))

        self._toner_intune_canvas = tk.Canvas(
            intune_frame, bg=bg, highlightthickness=0, width=20, height=20)
        self._toner_intune_canvas._dark_canvas = True
        self._toner_intune_canvas.pack(side="left")
        self._toner_intune_glow = self._toner_intune_canvas.create_oval(
            2, 2, 18, 18, fill="#0A1A0A", outline="")
        self._toner_intune_lamp = self._toner_intune_canvas.create_oval(
            4, 4, 16, 16, fill="#113311", outline="#444444", width=1)

        tk.Label(intune_frame, text="IN TUNE", bg=bg, fg=LABEL_DIM,
                 font=("Helvetica", 7, "bold")).pack(side="left", padx=(3, 0))

        gauge_row += 1

        # --- Descriptor gauges ---
        self._toner_gauges = {}
        gauge_defs = [
            ("richness", "Pure", "Complex"),    # Harmonic Spread
            ("warmth", "Thin", "Warm"),          # H2 Strength
        ]

        for key, left_label, right_label in gauge_defs:
            cv = tk.Canvas(gauge_frame, bg=bg, highlightthickness=0,
                           width=260, height=110)
            cv._dark_canvas = True
            cv.grid(row=gauge_row, column=0, columnspan=3, pady=(2, 0))
            gauge_data = self._toner_build_gauge(cv, left_label, right_label)
            self._toner_gauges[key] = gauge_data

            gauge_row += 1

        # --- Delta mode toggle + comparison-only gauges ---
        self._toner_delta_mode = tk.BooleanVar(value=False)
        self._toner_delta_frame = tk.Frame(gauge_frame, bg=bg)
        self._toner_delta_frame._skip_theme = True
        # Hidden until an overlay is loaded
        self._toner_delta_row = gauge_row

        delta_toggle = tk.Checkbutton(
            self._toner_delta_frame, text="\u0394 Delta",
            variable=self._toner_delta_mode, bg=bg, fg=LABEL_DIM,
            selectcolor="#333333", activebackground=bg,
            activeforeground=LABEL_DIM,
            font=("Helvetica", 8, "bold"))
        delta_toggle.pack(side="left", padx=4)
        gauge_row += 1

        # Comparison-only gauges (spectral tilt, mid-harmonic)
        self._toner_delta_gauges = {}
        self._toner_delta_gauge_frames = {}
        delta_vis = self.settings.get("visible_delta_gauges", {})
        delta_gauge_defs = [
            ("spectral_tilt", "Darker", "Brighter", "Spectral Tilt"),
            ("mid_harmonic", "Weaker", "Stronger", "Mid-Harmonic (H3\u2013H6)"),
        ]

        for key, left_label, right_label, title in delta_gauge_defs:
            frame = tk.Frame(gauge_frame, bg=bg)
            frame._skip_theme = True
            self._toner_delta_gauge_frames[key] = (frame, gauge_row)

            # Title label above gauge
            tk.Label(frame, text=title, bg=bg, fg=LABEL_DIM,
                     font=("Helvetica", 7)).pack()

            cv = tk.Canvas(frame, bg=bg, highlightthickness=0,
                           width=260, height=100)
            cv._dark_canvas = True
            cv.pack()
            gauge_data = self._toner_build_gauge(cv, left_label, right_label,
                                                  centered=True)
            self._toner_delta_gauges[key] = gauge_data
            gauge_row += 1

        # Initialize delta smooth values
        self._toner_delta_smooth = {
            'richness_delta': 0.0, 'warmth_delta': 0.0,
            'spectral_tilt': 0.0, 'mid_harmonic': 0.0,
        }

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

        self._toner_pause_btn = tk.Button(
            self._toner_capture_frame, text="Pause", font=("Helvetica", 9),
            command=self._toner_toggle_pause)
        self._toner_pause_btn.pack(side="right", padx=(0, 5), pady=4)

        # --- Bottom: controls (three-column layout) ---
        ctrl_bg = "systemWindowBackgroundColor" if IS_MACOS else CTRL_BG
        ctrl_fg = "white" if not IS_MACOS else "systemTextColor"
        self._toner_ctrl_bg = ctrl_bg
        self._toner_ctrl_fg = ctrl_fg

        ctrl_frame = tk.Frame(self._toner_main_frame, bg=ctrl_bg, padx=8, pady=10)
        ctrl_frame._skip_theme = True
        ctrl_frame.pack(fill="x", padx=5, pady=(0, 4))

        eq_lbl_font = ("Helvetica", 8)

        def _make_vswitch(parent, label, var, val_top, val_bottom,
                          lbl_top, lbl_bottom, col):
            """Create a vertical two-position switch."""
            ch = tk.Frame(parent, bg=ctrl_bg)
            ch._skip_theme = True
            ch.grid(row=0, column=col, padx=5, sticky="ns")
            tk.Label(ch, text=label, bg=ctrl_bg, fg="#888888",
                     font=eq_lbl_font).pack(pady=(0, 1))
            tk.Label(ch, text=lbl_top, bg=ctrl_bg, fg="#AAAAAA",
                     font=("Helvetica", 8), width=4).pack()
            int_var = tk.IntVar(value=0 if var.get() == val_top else 1)
            tk.Scale(ch, variable=int_var, from_=0, to=1,
                     orient="vertical", length=50, width=14,
                     showvalue=False, resolution=1,
                     bg="#B0B0B0", fg="#888888",
                     activebackground="#D0D0D0",
                     troughcolor="#444444", highlightthickness=0,
                     sliderrelief="raised", sliderlength=20,
                     borderwidth=2).pack()
            tk.Label(ch, text=lbl_bottom, bg=ctrl_bg, fg="#AAAAAA",
                     font=("Helvetica", 8), width=4).pack()

            def _on_change(*args):
                var.set(val_top if int_var.get() == 0 else val_bottom)
            int_var.trace_add("write", _on_change)

        # ========== LEFT COLUMN: display switches ==========
        left_col = tk.Frame(ctrl_frame, bg=ctrl_bg)
        left_col._skip_theme = True
        left_col.pack(side="left", padx=(0, 12))

        _make_vswitch(left_col, "VIEW", self._toner_view_var,
                      "spectrum", "bars", "spct", "bars", 0)
        _make_vswitch(left_col, "SCALE", self._toner_scale_var,
                      "linear", "db", "lin", "dB", 1)
        _make_vswitch(left_col, "FPS", self._toner_fps_var,
                      "30", "60", "30", "60", 2)

        # SENS as a two-position switch (high/low)
        self._toner_sens_var = tk.IntVar(
            value=toner_settings.get("sensitivity", 50))
        sens_str = tk.StringVar(
            value="high" if self._toner_sens_var.get() >= 50 else "low")

        def _sens_from_switch(*args):
            self._toner_sens_var.set(75 if sens_str.get() == "high" else 25)
            self._toner_on_sensitivity_changed()

        _make_vswitch(left_col, "SENS", sens_str,
                      "high", "low", "high", "low", 3)
        sens_str.trace_add("write", _sens_from_switch)

        # ========== CENTER COLUMN: sax selector ==========
        # Center column expands to push left/right apart and center the selector
        center_col = tk.Frame(ctrl_frame, bg=ctrl_bg)
        center_col._skip_theme = True
        center_col.pack(side="left", fill="x", expand=True)

        # Inner frame keeps the selector compact and centered
        sax_inner = tk.Frame(center_col, bg=ctrl_bg)
        sax_inner._skip_theme = True
        sax_inner.pack()  # centered within the expanding column

        tk.Label(sax_inner, text="SAX SELECTOR", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack()

        # Determine visible sax types from settings
        visible_sax = self.settings.get("visible_sax_types", None)
        if not visible_sax:
            visible_sax = SAX_TYPES[:]
        self._toner_visible_sax = visible_sax

        self._toner_sax_var = tk.StringVar(
            value=toner_settings.get("sax_type", "Alto"))

        # Horizontal scale mapped to sax type index
        sax_idx_var = tk.IntVar(value=0)
        if self._toner_sax_var.get() in visible_sax:
            sax_idx_var.set(visible_sax.index(self._toner_sax_var.get()))

        self._toner_sax_scale = tk.Scale(
            sax_inner, variable=sax_idx_var,
            from_=0, to=max(0, len(visible_sax) - 1),
            orient="horizontal", length=max(120, len(visible_sax) * 32),
            width=14, showvalue=False, resolution=1,
            bg="#B0B0B0", fg="#888888", activebackground="#D0D0D0",
            troughcolor="#444444", highlightthickness=0,
            sliderrelief="raised", sliderlength=20, borderwidth=2)
        self._toner_sax_scale.pack()
        self._toner_sax_idx_var = sax_idx_var

        # Labels under the slider
        lbl_frame = tk.Frame(sax_inner, bg=ctrl_bg)
        lbl_frame._skip_theme = True
        lbl_frame.pack(fill="x")
        _sax_abbrev = {"Sopranino": "Nino", "Soprano": "Sop", "F Mezzo": "Mez",
                       "Baritone": "Bari", "C Melody": "C Mel"}
        for i, stype in enumerate(visible_sax):
            abbrev = _sax_abbrev.get(stype, stype[:3] if len(stype) > 5 else stype)
            tk.Label(lbl_frame, text=abbrev, bg=ctrl_bg, fg="#AAAAAA",
                     font=("Helvetica", 6)).pack(side="left", expand=True)

        def _on_sax_changed(*args):
            idx = sax_idx_var.get()
            if 0 <= idx < len(visible_sax):
                self._toner_sax_var.set(visible_sax[idx])
                self._toner_on_sax_type_changed()
        sax_idx_var.trace_add("write", _on_sax_changed)

        # Lock sax selector when profile is loaded
        self._toner_sax_scale_widget = self._toner_sax_scale

        if self._toner_engine:
            self._toner_engine.set_sax_type(self._toner_sax_var.get())

        # A= and concert pitch moved to Options menu — just store vars
        self._toner_pitch_var = tk.DoubleVar(
            value=toner_settings.get("reference_pitch", 440.0))

        # ========== RIGHT COLUMN: all controls in a horizontal row ==========
        right_col = tk.Frame(ctrl_frame, bg=ctrl_bg)
        right_col._skip_theme = True
        right_col.pack(side="right", padx=(12, 0))

        # Session lamp
        self._toner_session_canvas = tk.Canvas(
            right_col, bg=ctrl_bg, highlightthickness=0, width=24, height=24)
        self._toner_session_canvas._dark_canvas = True
        self._toner_session_canvas.pack(side="left")
        self._toner_session_glow = self._toner_session_canvas.create_oval(
            2, 2, 22, 22, fill="#1A150A", outline="")
        self._toner_session_lamp = self._toner_session_canvas.create_oval(
            4, 4, 20, 20, fill="#332200", outline="#444444", width=1)

        tk.Label(right_col, text="SESSION", bg=ctrl_bg, fg=LABEL_DIM,
                 font=("Helvetica", 7, "bold")).pack(side="left", padx=(3, 10))

        # Profile indicator + Load/Unload toggle
        self._toner_profile_label = tk.Label(
            right_col, text="no profile", bg=ctrl_bg, fg="#666666",
            font=("Helvetica", 8))
        self._toner_profile_label.pack(side="left", padx=(0, 4))

        self._toner_load_unload_btn = tk.Button(
            right_col, text="Load...",
            font=("Helvetica", 9),
            command=self._toner_load_unload_toggle)
        self._toner_load_unload_btn.pack(side="left", padx=(0, 10))

        # MODE switch (free/cal)
        self._toner_mode_var = tk.StringVar(value="free")
        mode_ch = tk.Frame(right_col, bg=ctrl_bg)
        mode_ch._skip_theme = True
        mode_ch.pack(side="left", padx=(0, 8))
        tk.Label(mode_ch, text="MODE", bg=ctrl_bg, fg="#888888",
                 font=eq_lbl_font).pack(pady=(0, 1))
        tk.Label(mode_ch, text="free", bg=ctrl_bg, fg="#AAAAAA",
                 font=("Helvetica", 8), width=4).pack()
        mode_int = tk.IntVar(value=0)
        tk.Scale(mode_ch, variable=mode_int, from_=0, to=1,
                 orient="vertical", length=50, width=14,
                 showvalue=False, resolution=1,
                 bg="#B0B0B0", fg="#888888",
                 activebackground="#D0D0D0",
                 troughcolor="#444444", highlightthickness=0,
                 sliderrelief="raised", sliderlength=20,
                 borderwidth=2).pack()
        tk.Label(mode_ch, text="cal", bg=ctrl_bg, fg="#AAAAAA",
                 font=("Helvetica", 8), width=4).pack()

        def _on_mode_change(*args):
            self._toner_mode_var.set("free" if mode_int.get() == 0
                                     else "calibration")
        mode_int.trace_add("write", _on_mode_change)

        # Capture button
        self._toner_capture_btn = tk.Button(
            right_col, text="Capture",
            font=("Helvetica", 10, "bold"),
            command=self._toner_toggle_capture)
        self._toner_capture_btn.pack(side="left", ipady=4)

    def _create_toner_fallback(self, parent):
        """Fallback UI when audio libraries are unavailable."""
        bg = self.root.cget('bg')
        if sys.platform == 'linux':
            msg = ("The Tone Analyzer requires PortAudio.\n\n"
                   "Install it with:\n"
                   "  sudo apt install libportaudio2\n\n"
                   "Then restart the application.")
        elif sys.platform == 'darwin':
            msg = ("The Tone Analyzer is not available on this Mac.\n\n"
                   "This feature requires Apple Silicon (M1 or newer).\n\n"
                   "All other features work normally.")
        else:
            msg = ("The Tone Analyzer is not available on this system.\n\n"
                   "If you see this on Windows, try reinstalling the app.\n\n"
                   "All other features work normally.")
        tk.Label(parent, text=msg,
                 bg=bg, fg="gray", font=("Helvetica", 12),
                 justify="center").pack(expand=True)

    # ------------------------------------------------------------------
    # GAUGE BUILDER (VU-meter style)
    # ------------------------------------------------------------------

    def _toner_build_gauge(self, cv, left_label, right_label, centered=False):
        """Draw a VU-style arc gauge and return dict with IDs for animation.

        If centered=True, the gauge is a delta gauge: 0.5 = center/zero,
        with a prominent center tick mark and "0" label.
        """
        cv_w, cv_h = GAUGE_WIDTH, GAUGE_HEIGHT
        cv.configure(width=cv_w, height=cv_h)
        cx = cv_w // 2
        cy = cv_h - 8
        r = GAUGE_ARC_RADIUS
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

        # Center zero marker for delta gauges
        if centered:
            center_deg = (arc_start + arc_end) / 2
            center_rad = math.radians(center_deg)
            x_o = cx + (r + 3) * math.cos(center_rad)
            y_o = cy - (r + 3) * math.sin(center_rad)
            cv.create_text(x_o, y_o - 6, text="0",
                           fill=tick_color, font=("Helvetica", 7, "bold"))

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

        needle_len = r - GAUGE_DESCRIPTOR_NEEDLE_OFFSET
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

        # "Low data" overlay — hidden by default, shown when harmonics insufficient
        low_data_id = cv.create_text(
            cx, cy - 30, text="Low harmonic data",
            fill="#AA4400", font=("Helvetica", 8, "bold"),
            state="hidden")

        return {
            'canvas': cv, 'cx': cx, 'cy': cy,
            'needle_len': needle_len, 'needle_id': needle_id,
            'shadow_id': shadow_id, 'arc_start': arc_start, 'arc_end': arc_end,
            'low_data_id': low_data_id, 'centered': centered,
        }

    def _toner_build_intonation_gauge(self, cv):
        """Build the intonation VU gauge (±50 cents, same style as tuner VU)."""
        cv_w, cv_h = GAUGE_WIDTH, GAUGE_HEIGHT
        cv.configure(width=cv_w, height=cv_h)
        cx = cv_w // 2
        cy = cv_h - 8
        r = GAUGE_ARC_RADIUS
        arc_start = GAUGE_ARC_START
        arc_end = GAUGE_ARC_END

        tick_color = TICK_COLOR

        # Bezel + amber panel
        cv.create_rectangle(2, 2, cv_w - 2, cv_h - 2,
                            fill="#222222", outline="#3A3A3A", width=1)
        cv.create_rectangle(5, 5, cv_w - 5, cv_h - 5,
                            fill=AMBER, outline="#B87D08", width=1)

        # Tick marks for -INTONATION_RANGE_CENTS to +INTONATION_RANGE_CENTS in steps of 5
        half = INTONATION_RANGE_CENTS
        full = 2 * half
        for i in range(21):
            cents = -half + i * 5
            frac = (cents + half) / full
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
            frac = (cents + half) / full
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = cx + label_r * math.cos(angle_rad)
            ly = cy - label_r * math.sin(angle_rad)
            color = "#1B5E00" if cents == 0 else tick_color
            cv.create_text(lx, ly, text=label, fill=color,
                           font=label_font, anchor="center")

        # Flat/sharp at extremes
        for cents, label in [(-50, "\u266d"), (50, "\u266f")]:
            frac = (cents + half) / full
            angle_deg = arc_start + (arc_end - arc_start) * frac
            angle_rad = math.radians(angle_deg)
            lx = cx + (label_r + 10) * math.cos(angle_rad)
            ly = cy - (label_r + 10) * math.sin(angle_rad)
            cv.create_text(lx, ly, text=label, fill=tick_color,
                           font=("Helvetica", 10, "bold"), anchor="center")

        # Cents readout text (top-right corner, out of the way)
        self._toner_cents_text_id = cv.create_text(
            cv_w - 6, 6, text="", fill=tick_color,
            font=("Helvetica", 8), anchor="ne")

        # Needle + shadow
        needle_len = r - GAUGE_INTONATION_NEEDLE_OFFSET
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

        damping = GAUGE_INTONATION_DAMPING
        self._toner_smooth_cents += (cents - self._toner_smooth_cents) * damping

        clamped = max(-INTONATION_RANGE_CENTS, min(INTONATION_RANGE_CENTS, self._toner_smooth_cents))
        frac = (clamped + INTONATION_RANGE_CENTS) / (2 * INTONATION_RANGE_CENTS)
        angle_deg = gauge['arc_start'] + (gauge['arc_end'] - gauge['arc_start']) * frac
        angle_rad = math.radians(angle_deg)
        nx = gauge['cx'] + gauge['needle_len'] * math.cos(angle_rad)
        ny = gauge['cy'] - gauge['needle_len'] * math.sin(angle_rad)

        cv = gauge['canvas']
        cv.coords(gauge['shadow_id'],
                  gauge['cx'] + 1, gauge['cy'] + 1, nx + 1, ny + 1)
        cv.coords(gauge['needle_id'],
                  gauge['cx'], gauge['cy'], nx, ny)

        # Update cents readout + in-tune lamp
        in_tune = abs(self._toner_smooth_cents) < IN_TUNE_CENTS_THRESHOLD
        if hasattr(self, '_toner_intune_lamp'):
            if in_tune:
                self._toner_intune_canvas.itemconfigure(
                    self._toner_intune_lamp, fill="#00CC00")
                self._toner_intune_canvas.itemconfigure(
                    self._toner_intune_glow, fill="#004400")
            else:
                self._toner_intune_canvas.itemconfigure(
                    self._toner_intune_lamp, fill="#113311")
                self._toner_intune_canvas.itemconfigure(
                    self._toner_intune_glow, fill="#0A1A0A")

        if in_tune:
            cv.itemconfigure(self._toner_cents_text_id,
                             text="0\u00a2", fill="#1B5E00")
        elif self._toner_smooth_cents > 0:
            cv.itemconfigure(self._toner_cents_text_id,
                             text=f"+{self._toner_smooth_cents:.0f}\u00a2",
                             fill=TICK_COLOR)
        else:
            cv.itemconfigure(self._toner_cents_text_id,
                             text=f"{self._toner_smooth_cents:.0f}\u00a2",
                             fill=TICK_COLOR)

    def _toner_update_gauge(self, key, value):
        """Update a descriptor gauge needle to the given 0.0-1.0 value with damping."""
        gauge = self._toner_gauges.get(key)
        if not gauge:
            return

        damping = GAUGE_DESCRIPTOR_DAMPING
        self._toner_smooth[key] += (value - self._toner_smooth[key]) * damping
        frac = max(0.0, min(1.0, self._toner_smooth[key]))

        angle_deg = gauge['arc_start'] + (gauge['arc_end'] - gauge['arc_start']) * frac
        angle_rad = math.radians(angle_deg)
        nx = gauge['cx'] + gauge['needle_len'] * math.cos(angle_rad)
        ny = gauge['cy'] - gauge['needle_len'] * math.sin(angle_rad)

        cv = gauge['canvas']
        cv.coords(gauge['shadow_id'],
                  gauge['cx'] + 1, gauge['cy'] + 1, nx + 1, ny + 1)
        cv.coords(gauge['needle_id'],
                  gauge['cx'], gauge['cy'], nx, ny)

    def _toner_set_low_data_overlay(self, show):
        """Show or hide 'Low harmonic data' overlay on descriptor gauges."""
        self._toner_low_data_shown = show
        state = "normal" if show else "hidden"
        # Dim needles when showing, restore when hiding
        needle_color = "#554400" if show else "#1A1200"
        shadow_color = "#665522" if show else "#8A6500"
        for key, gauge in self._toner_gauges.items():
            cv = gauge['canvas']
            if 'low_data_id' in gauge:
                cv.itemconfigure(gauge['low_data_id'], state=state)
            cv.itemconfigure(gauge['needle_id'], fill=needle_color)
            cv.itemconfigure(gauge['shadow_id'], fill=shadow_color)

    def _toner_show_delta_gauges(self, show):
        """Show or hide the delta toggle and comparison-only gauges.

        When showing, auto-enables delta mode. The user can uncheck
        the toggle to return to absolute while keeping the overlay.
        """
        bg = BG_COLOR
        if show:
            self._toner_delta_mode.set(True)
            self._toner_delta_frame.grid(
                row=self._toner_delta_row, column=0, columnspan=3,
                sticky="w", padx=5)
            delta_vis = self.settings.get("visible_delta_gauges", {})
            for key, (frame, row) in self._toner_delta_gauge_frames.items():
                if delta_vis.get(key, True):
                    frame.grid(row=row, column=0, columnspan=3, pady=(2, 0))
        else:
            self._toner_delta_frame.grid_forget()
            for key, (frame, _) in self._toner_delta_gauge_frames.items():
                frame.grid_forget()
            self._toner_delta_mode.set(False)
            # Reset delta gauge needles to center
            for key, gauge in self._toner_delta_gauges.items():
                self._toner_delta_smooth[key] = 0.0
                center_angle = math.radians(
                    (gauge['arc_start'] + gauge['arc_end']) / 2)
                nx = gauge['cx'] + gauge['needle_len'] * math.cos(center_angle)
                ny = gauge['cy'] - gauge['needle_len'] * math.sin(center_angle)
                gauge['canvas'].coords(gauge['needle_id'],
                                       gauge['cx'], gauge['cy'], nx, ny)
                gauge['canvas'].coords(gauge['shadow_id'],
                                       gauge['cx'] + 1, gauge['cy'] + 1,
                                       nx + 1, ny + 1)

    def _toner_update_delta_gauge(self, key, value):
        """Update a delta gauge. Value is -1.0 to +1.0, mapped to 0.0-1.0."""
        gauge = self._toner_delta_gauges.get(key)
        if not gauge:
            return
        if value is None:
            # No data — center the needle
            frac = 0.5
        else:
            # Map -1..+1 to 0..1
            self._toner_delta_smooth[key] += (
                value - self._toner_delta_smooth[key]) * GAUGE_DESCRIPTOR_DAMPING
            frac = max(0.0, min(1.0, 0.5 + self._toner_delta_smooth[key] * 0.5))

        angle_deg = gauge['arc_start'] + (
            gauge['arc_end'] - gauge['arc_start']) * frac
        angle_rad = math.radians(angle_deg)
        nx = gauge['cx'] + gauge['needle_len'] * math.cos(angle_rad)
        ny = gauge['cy'] - gauge['needle_len'] * math.sin(angle_rad)

        cv = gauge['canvas']
        cv.coords(gauge['shadow_id'],
                  gauge['cx'] + 1, gauge['cy'] + 1, nx + 1, ny + 1)
        cv.coords(gauge['needle_id'], gauge['cx'], gauge['cy'], nx, ny)

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
        for i in range(MAX_HARMONICS):
            marker = cv.create_rectangle(0, 0, 0, 0, fill=HARMONIC_COLOR,
                                         outline="", width=0, state="hidden")
            self._toner_harmonic_markers.append(marker)

        # Ghost overlay markers (for comparison profile)
        for i in range(MAX_HARMONICS):
            ghost = cv.create_line(0, 0, 0, 0, fill=GHOST_COLOR,
                                   width=2, state="hidden")
            self._toner_ghost_markers.append(ghost)

        # Frequency axis labels
        label_font = ("Helvetica", 7)
        for freq in [100, 500, 1000, 2000, 4000, int(SPECTRUM_MAX_FREQ)]:
            x = (freq / SPECTRUM_MAX_FREQ) * w
            cv.create_text(x, h - 3, text=f"{freq}" if freq < 1000 else f"{freq // 1000}k",
                           fill="#555555", font=label_font, anchor="s")
            cv.create_line(x, 0, x, h - 12, fill="#333333", width=1, dash=(2, 4))

        # Note name overlay
        self._toner_spectrum_note = cv.create_text(
            10, 10, text="", fill=FUNDAMENTAL_COLOR,
            font=("Helvetica", 16, "bold"), anchor="nw")

        # Pitch mode indicator (e.g. "Written (Tenor)" or "Concert")
        self._toner_pitch_mode_label = cv.create_text(
            10, 30, text="", fill="#888888",
            font=("Helvetica", 8), anchor="nw")
        self._toner_update_pitch_mode_label()

        # Large calibration prompt (visible from across the room)
        self._toner_cal_prompt = cv.create_text(
            0, 0, text="", fill="#FFCC00",
            font=("Helvetica", 48, "bold"), anchor="ne", state="hidden")
        # Status line below the big prompt
        self._toner_cal_status = cv.create_text(
            0, 0, text="", fill="#AAAAAA",
            font=("Helvetica", 14), anchor="ne", state="hidden")

        # Comparison label
        self._toner_compare_label = cv.create_text(
            10, 44, text="", fill=GHOST_COLOR,
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
                        freq_frac = hi.expected_freq / SPECTRUM_MAX_FREQ
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
        # per_note keys are concert pitch, result.fundamental_note is concert pitch
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
        note = result.fundamental_note  # concert pitch matches per_note keys
        if note and 'per_note' in fp and note in fp['per_note']:
            ghost_db = fp['per_note'][note].get('harmonics_db', ghost_db)

        f0 = result.fundamental_freq
        for idx, g in enumerate(self._toner_ghost_markers):
            if idx < len(ghost_db):
                freq = f0 * (idx + 1)
                freq_frac = freq / SPECTRUM_MAX_FREQ
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
        self._toner_update_pitch_mode_label()

    def _toner_lock_sax_selector(self, locked):
        """Lock or unlock the sax selector (locked when profile is loaded)."""
        if hasattr(self, '_toner_sax_scale_widget'):
            state = "disabled" if locked else "normal"
            self._toner_sax_scale_widget.configure(state=state)

    def _toner_open_pitch_dialog(self):
        """Open reference pitch (A=) dialog."""
        result = simpledialog.askfloat("Reference Pitch",
            "A = ? Hz", initialvalue=self._toner_pitch_var.get(),
            minvalue=420, maxvalue=460, parent=self.root)
        if result is not None:
            self._toner_pitch_var.set(result)
            self._toner_on_pitch_changed()

    def _toner_open_pitch_display_dialog(self):
        """Toggle between written and concert pitch display."""
        current = "concert" if self._toner_concert_pitch.get() else "written"
        dlg = tk.Toplevel(self.root)
        dlg.title("Display Pitch")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Note Display", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 10))

        var = tk.StringVar(value=current)
        tk.Radiobutton(frame, text="Written pitch (what the player fingers)",
                       variable=var, value="written", bg=bg, fg=fg,
                       font=("Helvetica", 10)).pack(anchor="w")
        tk.Radiobutton(frame, text="Concert pitch (actual sounding frequency)",
                       variable=var, value="concert", bg=bg, fg=fg,
                       font=("Helvetica", 10)).pack(anchor="w")

        def apply():
            self._toner_concert_pitch.set(var.get() == "concert")
            self._toner_update_pitch_mode_label()
            dlg.destroy()

        tk.Button(frame, text="OK", command=apply, width=10).pack(pady=(10, 0))

    def _toner_transpose_note(self, concert_note):
        """Transpose a concert pitch note name to written pitch for display.

        Returns the original note if concert pitch display is on.
        Uses the engine's pure transpose_note() function.
        """
        if not concert_note or self._toner_concert_pitch.get():
            return concert_note
        return transpose_note(concert_note, self._toner_sax_var.get())

    def _toner_display_note_for_profile(self, concert_note, profile=None):
        """Transpose a concert note for display using a profile's horn type.

        Used in reports and comparisons where the active sax selector might
        not match the profile being viewed.
        """
        if not concert_note or self._toner_concert_pitch.get():
            return concert_note
        if profile:
            sax_type = profile.get('horn_type', self._toner_sax_var.get())
        else:
            sax_type = self._toner_sax_var.get()
        return transpose_note(concert_note, sax_type)

    def _toner_update_pitch_mode_label(self):
        """Update the pitch mode indicator on the spectrum canvas."""
        if not hasattr(self, '_toner_pitch_mode_label'):
            return
        cv = self._toner_spectrum_canvas
        if self._toner_concert_pitch.get():
            cv.itemconfigure(self._toner_pitch_mode_label, text="Concert")
        else:
            sax = self._toner_sax_var.get()
            cv.itemconfigure(self._toner_pitch_mode_label,
                             text=f"Written ({sax})")

    def _toner_update_session_lamp(self):
        """Update the session indicator lamp (amber when session is active)."""
        if not hasattr(self, '_toner_session_lamp'):
            return
        active = self._toner_active_session is not None
        if active:
            self._toner_session_canvas.itemconfigure(
                self._toner_session_lamp, fill="#CC8800")
            self._toner_session_canvas.itemconfigure(
                self._toner_session_glow, fill="#332200")
        else:
            self._toner_session_canvas.itemconfigure(
                self._toner_session_lamp, fill="#332200")
            self._toner_session_canvas.itemconfigure(
                self._toner_session_glow, fill="#1A150A")

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
            if hasattr(self, '_toner_resize_id'):
                self.root.after_cancel(self._toner_resize_id)
            self._toner_resize_id = self.root.after(50, self._toner_build_spectrum_bars)

    # ------------------------------------------------------------------
    # PROFILE MANAGEMENT
    # ------------------------------------------------------------------

    def _toner_open_profile_dialog(self):
        """Open the profile management dialog — central hub for all profile operations."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Tone Profiles")
        dlg.geometry("550x500")
        dlg.minsize(400, 350)
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Tone Profiles", bg=bg, fg=fg,
                 font=("Helvetica", 14, "bold")).pack(pady=(0, 10))

        # Profile list
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(fill="both", expand=True, pady=(0, 5))

        self._prof_listbox = tk.Listbox(list_frame, width=55, height=12,
                                         font=("Helvetica", 10))
        self._prof_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, command=self._prof_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._prof_listbox.config(yscrollcommand=scrollbar.set)

        self._toner_refresh_profile_list()

        # Info display
        self._prof_info_label = tk.Label(frame, text="Select a profile to see details.",
                                          bg=bg, fg=fg,
                                          font=("Helvetica", 9),
                                          justify="left", anchor="w",
                                          wraplength=500)
        self._prof_info_label.pack(fill="x", pady=(0, 10))

        self._prof_listbox.bind("<<ListboxSelect>>",
                                lambda e: self._toner_on_profile_selected())

        # --- Action buttons: two rows ---
        # Row 1: profile operations
        row1 = tk.Frame(frame, bg=bg)
        row1.pack(fill="x", pady=(0, 5))

        tk.Button(row1, text="Load for Capture",
                  command=lambda: self._toner_load_from_dialog(dlg)).pack(
                      side="left", padx=(0, 5))
        tk.Button(row1, text="Analyze...",
                  command=lambda: [dlg.destroy(),
                                   self._toner_open_analyze_dialog()]).pack(
                      side="left", padx=(0, 5))
        tk.Button(row1, text="Edit Notes...",
                  command=self._toner_edit_profile_notes).pack(
                      side="left", padx=(0, 5))

        # Row 2: create, import, delete, close
        row2 = tk.Frame(frame, bg=bg)
        row2.pack(fill="x")

        tk.Button(row2, text="New Profile...",
                  command=lambda: self._toner_new_profile(dlg)).pack(
                      side="left", padx=(0, 5))
        tk.Button(row2, text="Import Audio File...",
                  command=lambda: [dlg.destroy(),
                                   self._toner_import_audio_file()]).pack(
                      side="left", padx=(0, 5))
        tk.Button(row2, text="Delete",
                  command=self._toner_delete_profile).pack(
                      side="left", padx=(0, 5))
        tk.Button(row2, text="Close",
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
        self._toner_notes_dialog(prof_name, profile)

    def _toner_notes_dialog(self, prof_name, profile, prompt_text=None):
        """Open a multi-line notes editor for a profile."""
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Notes \u2014 {prof_name}")
        dlg.resizable(True, True)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        frame = tk.Frame(dlg, bg=bg, padx=15, pady=10)
        frame.pack(fill="both", expand=True)

        if prompt_text:
            tk.Label(frame, text=prompt_text, bg=bg,
                     font=("Helvetica", 10), wraplength=400,
                     justify="left").pack(pady=(0, 8))

        tk.Label(frame, text=f"Notes for \"{prof_name}\":", bg=bg,
                 font=("Helvetica", 10)).pack(anchor="w", pady=(0, 4))

        text_frame = tk.Frame(frame)
        text_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        notes_text = tk.Text(text_frame, height=8, width=50,
                              font=("Helvetica", 10), wrap="word",
                              yscrollcommand=scrollbar.set)
        notes_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=notes_text.yview)

        current = profile.get('notes', '')
        notes_text.insert("1.0", current)
        notes_text.focus_set()

        def save():
            new_notes = notes_text.get("1.0", tk.END).strip()
            profile['notes'] = new_notes
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            # Refresh info display if profile dialog is open
            if hasattr(self, '_prof_info_label'):
                try:
                    self._toner_on_profile_selected()
                except Exception:
                    pass
            dlg.destroy()

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x", pady=(8, 0))
        tk.Button(btn_frame, text="Save", command=save, width=10).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy,
                  width=10).pack(side="left")

        dlg.protocol("WM_DELETE_WINDOW", dlg.destroy)
        dlg.minsize(400, 250)

    def _toner_build_profile_fields(self, frame, bg, fg):
        """Build profile form fields. Returns (fields_dict, lib_var, notes_text).

        Required fields are always shown. Optional fields respect
        the visible_profile_fields setting.
        """
        fields = {}
        vis = self.settings.get("visible_profile_fields", {})

        def add_field(label, key, default="", widget_type="entry",
                      optional_key=None):
            """Add a labeled field row. If optional_key is set, only show
            when that key is enabled in visible_profile_fields."""
            if optional_key and not vis.get(optional_key, False):
                return
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg=bg, fg=fg, width=14,
                     anchor="e", font=("Helvetica", 10)).pack(
                side="left", padx=(0, 8))
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
                 anchor="e", font=("Helvetica", 10)).pack(
            side="left", padx=(0, 8))
        existing_libs = [k for k in self._toner_profiles.keys()
                        if isinstance(self._toner_profiles[k], dict)]
        if not existing_libs:
            existing_libs = [DEFAULT_LIBRARY]
        lib_var = tk.StringVar(value=existing_libs[0])
        ttk.Combobox(lib_row, textvariable=lib_var,
                      values=existing_libs, width=20).pack(
            side="left", fill="x", expand=True)

        # Required fields
        add_field("Profile Name:", "name")
        add_field("Horn Type:", "horn_type", "Alto", widget_type="combo")
        add_field("Make:", "horn_make")
        add_field("Model:", "horn_model")
        add_field("Player:", "player")
        add_field("Mouthpiece:", "mouthpiece")

        # Optional fields
        add_field("Serial #:", "serial", optional_key="serial")
        add_field("Reed:", "reed", optional_key="reed")
        add_field("Ligature:", "ligature", optional_key="ligature")
        add_field("Room:", "room", optional_key="room")
        add_field("Preamp:", "preamp", optional_key="preamp")
        add_field("Mic Model:", "mic_model_field",
                  default=self.settings.get("mic_model", ""),
                  optional_key="mic_model")

        # Notes — multi-line (optional)
        notes_text = None
        if vis.get("notes", True):
            notes_row = tk.Frame(frame, bg=bg)
            notes_row.pack(fill="x", pady=2)
            tk.Label(notes_row, text="Notes:", bg=bg, fg=fg, width=14,
                     anchor="e", font=("Helvetica", 10)).pack(
                side="left", padx=(0, 8), anchor="n")
            notes_text = tk.Text(notes_row, height=3, width=25,
                                 font=("Helvetica", 10), wrap="word")
            notes_text.pack(side="left", fill="x", expand=True)

        return fields, lib_var, notes_text

    def _toner_validate_required_fields(self, fields, dlg):
        """Check required fields are filled. Returns True if valid."""
        required = [
            ("name", "Profile Name"),
            ("horn_make", "Make"),
            ("horn_model", "Model"),
            ("player", "Player"),
            ("mouthpiece", "Mouthpiece"),
        ]
        for key, label in required:
            if key in fields and not fields[key].get().strip():
                messagebox.showwarning("Required Field",
                    f"Please enter {label}.", parent=dlg)
                return False
        return True

    def _toner_collect_profile_data(self, fields, notes_text):
        """Collect profile data dict from form fields."""
        data = {
            'horn_type': fields.get("horn_type", tk.StringVar()).get(),
            'horn_make': fields.get("horn_make", tk.StringVar()).get().strip(),
            'horn_model': fields.get("horn_model", tk.StringVar()).get().strip(),
            'serial': fields.get("serial", tk.StringVar()).get().strip() if "serial" in fields else "",
            'player': fields.get("player", tk.StringVar()).get().strip(),
            'mouthpiece': fields.get("mouthpiece", tk.StringVar()).get().strip(),
            'reed': fields.get("reed", tk.StringVar()).get().strip() if "reed" in fields else "",
            'ligature': fields.get("ligature", tk.StringVar()).get().strip() if "ligature" in fields else "",
            'room': fields.get("room", tk.StringVar()).get().strip() if "room" in fields else "",
            'preamp': fields.get("preamp", tk.StringVar()).get().strip() if "preamp" in fields else "",
            'notes': notes_text.get("1.0", tk.END).strip() if notes_text else "",
            'created': time.strftime("%Y-%m-%d"),
            'sessions': [],
        }
        return data

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

        fields, lib_var, notes_text = self._toner_build_profile_fields(
            frame, bg, fg)

        def save():
            if not self._toner_validate_required_fields(fields, dlg):
                return
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if lib not in self._toner_profiles:
                self._toner_profiles[lib] = {}
            if name in self._toner_profiles[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            self._toner_profiles[lib][name] = self._toner_collect_profile_data(
                fields, notes_text)
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            self._toner_active_library = lib
            self._toner_active_profile = name
            self._toner_active_session = None
            self._toner_update_profile_label()
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
        """Toggle capture mode on/off."""
        if self._toner_capture_state is not None:
            # Stop capturing
            self._toner_stop_capture()
            return

        # Need a profile loaded
        if not self._toner_active_profile:
            messagebox.showinfo("No Profile Loaded",
                "Load a profile first to capture data.\n\n"
                "Use Load... on the control strip, or\n"
                "File > Profiles to create and manage profiles.")
            self._toner_open_profile_dialog()
            return

        # Start a new session if needed
        if not self._toner_active_session:
            self._toner_start_new_session_and_listen()
        else:
            self._toner_begin_listening()

    def _toner_begin_listening(self):
        """Enter listening mode (called after session is ready)."""
        self._toner_capture_mode = self._toner_mode_var.get()

        if self._toner_capture_mode == "calibration":
            # First-run tutorial
            if not self.settings.get("seen_calibration_tutorial"):
                self.settings["seen_calibration_tutorial"] = True
                save_settings(self.settings)
                messagebox.showinfo("Calibration Capture",
                    "Calibration walks through every note in the saxophone's "
                    "range, one at a time.\n\n"
                    "How it works:\n"
                    "  \u2022 A 10-second countdown gives you time to get ready\n"
                    "  \u2022 A large note name appears on screen\n"
                    "  \u2022 Play that note and hold it steady\n"
                    "  \u2022 The app records for 5 seconds, then waits for silence\n"
                    "  \u2022 Stop playing, and the next note appears\n"
                    "  \u2022 Work through all 32 notes at your own pace\n\n"
                    "Tips:\n"
                    "  \u2022 Use Pause if you need to adjust or take a break\n"
                    "  \u2022 Play at a comfortable mezzo-forte dynamic\n"
                    "  \u2022 If the app triggers on chair noise, raise the\n"
                    "    capture threshold in Options > Capture Threshold")

            # Build calibration note list, filtering out notes whose
            # concert frequency is below the engine's detection floor
            self._toner_cal_notes = [
                n for n in CALIBRATION_NOTES
                if note_to_freq(reverse_transpose_note(n, self._toner_sax_var.get())) >= MIN_FUNDAMENTAL_HZ
            ]
            skipped = len(CALIBRATION_NOTES) - len(self._toner_cal_notes)
            if skipped > 0:
                messagebox.showinfo("Note Range",
                    f"{skipped} low notes skipped — below detection range "
                    f"for {self._toner_sax_var.get()}.\n"
                    f"Calibration will cover {len(self._toner_cal_notes)} notes.")

            self._toner_cal_index = 0
            self._toner_cal_recording = False
            self._toner_capture_frames = []
            self._toner_capture_start = time.time()
            self._toner_capture_state = 'cal_countdown'
            label_text = f"Calibration starting in {CAL_COUNTDOWN_S:.0f}s... get ready"
        else:
            # Free mode: continuous auto-captures
            self._toner_stable_threshold = FREE_STABLE_FRAMES
            self._toner_capture_state = 'free_listening'
            self._toner_free_accumulator = []
            label_text = "Play anything..."

        self._toner_stable_note = ""
        self._toner_stable_count = 0
        self._toner_capture_btn.configure(text="Stop")

        self._toner_capture_frame.pack(fill="x", padx=5,
            before=self._toner_main_frame.winfo_children()[-1])
        self._toner_capture_label.configure(text=label_text)
        self._toner_capture_progress.configure(text="")


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
                 "mouthpiece.\nChange any variable? That's a new profile.",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(pady=(0, 10))

        fields, lib_var, notes_text = self._toner_build_profile_fields(
            frame, bg, fg)

        def save_and_start():
            if not self._toner_validate_required_fields(fields, dlg):
                return
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if lib not in self._toner_profiles:
                self._toner_profiles[lib] = {}
            if name in self._toner_profiles[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            self._toner_profiles[lib][name] = self._toner_collect_profile_data(
                fields, notes_text)
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

    def _toner_load_from_dialog(self, dlg):
        """Load the selected profile from the profile dialog."""
        sel = self._prof_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select Profile", "Select a profile first.")
            return
        key = self._prof_list_keys[sel[0]]
        if key is None:
            return
        lib_name, prof_name = key
        self._toner_active_library = lib_name
        self._toner_active_profile = prof_name
        self._toner_active_session = None
        self._toner_update_profile_label()

        # Sync sax type and lock selector
        prof = self._toner_profiles[lib_name][prof_name]
        sax_type = prof.get('horn_type', '')
        if sax_type:
            self._toner_sax_var.set(sax_type)
            if sax_type in self._toner_visible_sax:
                self._toner_sax_idx_var.set(
                    self._toner_visible_sax.index(sax_type))
            if self._toner_engine:
                self._toner_engine.set_sax_type(sax_type)
        self._toner_lock_sax_selector(True)
        dlg.destroy()

    def _toner_load_unload_toggle(self):
        """Toggle between Load and Unload based on current profile state."""
        if self._toner_active_profile:
            self._toner_unload_profile()
        else:
            self._toner_load_profile_quick()

    def _toner_update_load_unload_btn(self):
        """Update the Load/Unload button text based on profile state."""
        if hasattr(self, '_toner_load_unload_btn'):
            if self._toner_active_profile:
                self._toner_load_unload_btn.configure(text="Unload")
            else:
                self._toner_load_unload_btn.configure(text="Load...")

    def _toner_load_profile_quick(self):
        """Quick profile loader — shows list, loads selected."""
        all_profiles = []
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict):
                continue
            for prof_name, prof_data in lib_profiles.items():
                all_profiles.append((lib_name, prof_name, prof_data))

        if not all_profiles:
            messagebox.showinfo("No Profiles",
                "No profiles yet. Create one in File > Profiles.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Load Profile")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        frame = tk.Frame(dlg, bg=bg, padx=15, pady=10)
        frame.pack(fill="both", expand=True)

        listbox = tk.Listbox(frame, width=40, height=min(10, len(all_profiles)),
                              font=("Helvetica", 10))
        listbox.pack(fill="both", expand=True, pady=(0, 8))

        for lib_name, prof_name, _ in all_profiles:
            listbox.insert(tk.END, f"[{lib_name}] {prof_name}")

        def load():
            sel = listbox.curselection()
            if not sel:
                return
            lib_name, prof_name, _ = all_profiles[sel[0]]
            self._toner_active_library = lib_name
            self._toner_active_profile = prof_name
            self._toner_active_session = None
            self._toner_update_profile_label()

            # Sync sax type and lock selector
            prof = self._toner_profiles[lib_name][prof_name]
            sax_type = prof.get('horn_type', '')
            if sax_type:
                self._toner_sax_var.set(sax_type)
                if sax_type in self._toner_visible_sax:
                    self._toner_sax_idx_var.set(
                        self._toner_visible_sax.index(sax_type))
                if self._toner_engine:
                    self._toner_engine.set_sax_type(sax_type)
            self._toner_lock_sax_selector(True)
            dlg.destroy()

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Load", command=load).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="left")

    def _toner_unload_profile(self):
        """Unload the active profile."""
        self._toner_active_library = None
        self._toner_active_profile = None
        self._toner_active_session = None
        self._toner_update_profile_label()
        self._toner_lock_sax_selector(False)

    def _toner_signal_above_threshold(self, result):
        """Check if the signal level is above the capture threshold."""
        threshold = self.settings.get("capture_threshold", 50) / 100.0
        return (result.fundamental_freq > 0 and result.harmonics
                and result.signal_level > threshold)

    def _toner_update_cal_prompt(self, text, status=""):
        """Update the large calibration prompt on the spectrum canvas."""
        cv = self._toner_spectrum_canvas
        if hasattr(self, '_toner_cal_prompt'):
            if text:
                w = cv.winfo_width()
                cv.coords(self._toner_cal_prompt, w - 10, 10)
                cv.itemconfigure(self._toner_cal_prompt, text=text, state="normal")
                cv.coords(self._toner_cal_status, w - 10, 65)
                cv.itemconfigure(self._toner_cal_status,
                                 text=status, state="normal" if status else "hidden")
            else:
                cv.itemconfigure(self._toner_cal_prompt, state="hidden")
                cv.itemconfigure(self._toner_cal_status, state="hidden")

    def _toner_update_profile_label(self):
        """Update the active profile indicator and Load/Unload button."""
        if hasattr(self, '_toner_profile_label'):
            name = self._toner_active_profile or ""
            if name:
                display = name[:PROFILE_NAME_MAX_DISPLAY] + "..." if len(name) > PROFILE_NAME_MAX_DISPLAY else name
                self._toner_profile_label.configure(text=display, fg="#AAAAAA")
            else:
                self._toner_profile_label.configure(text="no profile", fg="#666666")
        self._toner_update_load_unload_btn()

    def _toner_start_new_session_and_listen(self):
        """Create a new session for the active profile and begin listening."""
        # Prompt for mic type if not set
        if not self.settings.get('mic_type'):
            messagebox.showinfo("Mic Type Required",
                "Please set your microphone type in Options \u2192 "
                "Input Device before capturing.\n\n"
                "A condenser mic is required for full harmonic "
                "analysis. Ribbon and dynamic mics can still be "
                "used but with reduced accuracy for upper harmonics.",
                parent=self.root)
            return

        self._toner_update_profile_label()
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
            'date': time.strftime("%Y-%m-%d %H:%M:%S"),
            'captures': [],
            'mic_type': self.settings.get('mic_type', ''),
            'mic_model': self.settings.get('mic_model', ''),
        }
        self._toner_rolloff_warned = False
        self._toner_begin_listening()

    def _toner_stop_capture(self):
        """Stop capture mode. Saves pending data, shows coverage summary."""
        # Save any accumulated free-mode frames before stopping
        if self._toner_capture_mode == 'free' and self._toner_free_accumulator:
            self._toner_free_save_micro_capture()

        # Compute and store recording quality metric for the session
        if self._toner_active_session and self._toner_active_session.get('captures'):
            rates = []
            for cap in self._toner_active_session['captures']:
                r = compute_rolloff_rate(cap.get('harmonics_db', []))
                if r is not None:
                    rates.append(r)
            if rates:
                self._toner_active_session['rolloff_rate'] = round(
                    sum(rates) / len(rates), 2)
                self._toner_save_active_session()

        self._toner_capture_state = None
        self._toner_capture_frames = []
        self._toner_free_accumulator = []
        self._toner_stable_note = ""
        self._toner_stable_count = 0
        self._toner_paused = False
        self._toner_paused_state = None
        self._toner_capture_frame.pack_forget()
        self._toner_update_cal_prompt("")
        if hasattr(self, '_toner_capture_btn'):
            self._toner_capture_btn.configure(text="Capture")
        if hasattr(self, '_toner_pause_btn'):
            self._toner_pause_btn.configure(text="Pause")

        # Show coverage summary if we have captures
        if (self._toner_active_session and
                self._toner_active_session.get('captures')):
            self._toner_show_coverage_summary()

    def _toner_show_coverage_summary(self):
        """Show a coverage summary with note distribution and resume option."""
        session = self._toner_active_session
        if not session:
            return

        captures = session.get('captures', [])
        if not captures:
            return

        # Get active profile for note transposition
        _cov_profile = None
        if self._toner_active_library and self._toner_active_profile:
            lib = self._toner_profiles.get(self._toner_active_library, {})
            _cov_profile = lib.get(self._toner_active_profile)

        dlg = tk.Toplevel(self.root)
        dlg.title("Capture Summary")
        dlg.resizable(False, False)
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        prof_name = self._toner_active_profile or "?"
        tk.Label(frame, text=f"Session: {prof_name}", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 5))

        # Count captures per note
        from collections import Counter
        note_counts = Counter(c.get('note', '') for c in captures)
        total = len(captures)
        unique = len(note_counts)

        tk.Label(frame, text=f"{total} captures across {unique} unique notes",
                 bg=bg, fg=fg, font=("Helvetica", 10)).pack(pady=(0, 10))

        # Build note distribution chart
        chart_cv = tk.Canvas(frame, bg="white", highlightthickness=1,
                              highlightbackground="#CCCCCC",
                              width=500, height=180)
        chart_cv.pack(pady=(0, 10))

        # Sort notes chromatically
        all_notes = sorted(note_counts.keys(),
                          key=lambda n: _note_sort_key(n))

        if all_notes:
            margin_l, margin_b, margin_t, margin_r = 35, 25, 10, 10
            cw = 500 - margin_l - margin_r
            ch = 180 - margin_t - margin_b
            max_count = max(note_counts.values())
            bar_w = max(4, min(20, cw / len(all_notes) - 2))
            total_bar_w = (bar_w + 2) * len(all_notes)
            start_x = margin_l + (cw - total_bar_w) / 2

            # Y axis
            for i in range(max_count + 1):
                y = margin_t + ch - (i / max(1, max_count)) * ch
                chart_cv.create_line(margin_l - 3, y, margin_l, y,
                                      fill="#888888")
                if i % max(1, max_count // 4) == 0 or i == max_count:
                    chart_cv.create_text(margin_l - 5, y, text=str(i),
                                          fill="#888888",
                                          font=("Helvetica", 7), anchor="e")

            # Bars
            # Color by register: low=blue, mid=green, high=orange
            for i, note in enumerate(all_notes):
                count = note_counts[note]
                x = start_x + i * (bar_w + 2)
                bar_h = (count / max(1, max_count)) * ch

                # Register coloring
                key = _note_sort_key(note)
                if key < 48:    # Below C4
                    color = "#4477CC"
                elif key < 72:  # C4 to B5
                    color = "#44AA44"
                else:           # C6+
                    color = "#CC7744"

                chart_cv.create_rectangle(
                    x, margin_t + ch - bar_h,
                    x + bar_w, margin_t + ch,
                    fill=color, outline="")

                # Note label (rotated text not supported, so abbreviated)
                if len(all_notes) <= 24 or i % 2 == 0:
                    disp_note = self._toner_display_note_for_profile(
                        note, _cov_profile)
                    chart_cv.create_text(
                        x + bar_w / 2, margin_t + ch + 10,
                        text=disp_note, fill="#444444",
                        font=("Helvetica", 6), angle=45 if len(all_notes) > 12 else 0)

            # Legend
            legend_y = margin_t + 5
            for label, color in [("Low", "#4477CC"), ("Mid", "#44AA44"), ("High", "#CC7744")]:
                chart_cv.create_rectangle(margin_l + 5, legend_y,
                                           margin_l + 15, legend_y + 8,
                                           fill=color, outline="")
                chart_cv.create_text(margin_l + 18, legend_y + 4, text=label,
                                      fill="#444444", font=("Helvetica", 7),
                                      anchor="w")
                legend_y += 12

        # Coverage assessment
        assessment = []
        low_count = sum(1 for n in note_counts if _note_sort_key(n) < 48)
        mid_count = sum(1 for n in note_counts if 48 <= _note_sort_key(n) < 72)
        high_count = sum(1 for n in note_counts if _note_sort_key(n) >= 72)

        if low_count == 0:
            assessment.append("No low register notes captured")
        if mid_count == 0:
            assessment.append("No mid register notes captured")
        if high_count == 0:
            assessment.append("No high register notes captured")
        if unique < MIN_PROFILE_NOTES:
            assessment.append(f"Need {MIN_PROFILE_NOTES - unique} more unique "
                            f"notes for fingerprint")

        if assessment:
            gaps = "\n".join(assessment)
            tk.Label(frame, text=gaps, bg=bg, fg="#884400",
                     font=("Helvetica", 9), justify="left").pack(pady=(0, 8))
        else:
            tk.Label(frame, text="Good coverage across registers!",
                     bg=bg, fg="#006600",
                     font=("Helvetica", 9)).pack(pady=(0, 8))

        # Buttons
        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")

        def resume():
            dlg.destroy()
            self._toner_begin_listening()

        def done_with_notes():
            dlg.destroy()
            # Clear session — it's saved, don't carry it to the next profile
            session_lib = self._toner_active_library
            session_prof = self._toner_active_profile
            self._toner_active_session = None
            # Prompt for notes after session
            lib = session_lib
            prof_name = session_prof
            if lib and prof_name and lib in self._toner_profiles:
                profile = self._toner_profiles[lib].get(prof_name)
                if profile:
                    self._toner_notes_dialog(prof_name, profile,
                        prompt_text="How did this horn sound to you? "
                        "Add your impressions \u2014 bright, dark, rich, "
                        "stuffy, free-blowing, anything you noticed.")

        def discard():
            if messagebox.askyesno("Discard Session",
                    "Discard all captures from this session?\n\n"
                    "This cannot be undone.", parent=dlg):
                # Remove this session's captures from the profile
                lib = self._toner_active_library
                prof_name = self._toner_active_profile
                if lib and prof_name and lib in self._toner_profiles:
                    profile = self._toner_profiles[lib].get(prof_name)
                    if profile and self._toner_active_session:
                        session_date = self._toner_active_session.get('date', '')
                        sessions = profile.get('sessions', [])
                        profile['sessions'] = [s for s in sessions
                                               if s.get('date') != session_date]
                        save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
                self._toner_active_session = None
                dlg.destroy()

        tk.Button(btn_frame, text="Resume Capturing",
                  command=resume).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Done",
                  command=done_with_notes).pack(side="left", padx=(0, 5))

        # Discard button — far right, separated
        tk.Button(btn_frame, text="Discard Session", fg="#CC0000",
                  font=("Helvetica", 8),
                  command=discard).pack(side="right")

    def _toner_toggle_pause(self):
        """Toggle pause during capture."""
        if self._toner_paused:
            # Resume
            self._toner_capture_state = self._toner_paused_state
            self._toner_paused = False
            self._toner_pause_btn.configure(text="Pause")
            self._toner_capture_label.configure(text="Resumed...")
        else:
            # Pause
            self._toner_paused_state = self._toner_capture_state
            self._toner_capture_state = 'paused'
            self._toner_paused = True
            self._toner_pause_btn.configure(text="Resume")
            self._toner_capture_label.configure(text="Paused")
            self._toner_update_cal_prompt("PAUSED")

    def _toner_process_capture_frame(self, result):
        """Called each animation frame. Drives the auto-capture state machine."""
        state = self._toner_capture_state
        if state is None or state == 'paused':
            return

        # Store concert pitch (what the mic heard). Transpose only for display.
        note = result.fundamental_note if result.fundamental_freq > 0 else ""

        # --- FREE LISTENING: continuous micro-capture mode ---
        if state == 'free_listening':
            # Accumulate frames while note is stable (skip attack transient)
            if note and result.harmonics:
                if note == self._toner_stable_note:
                    self._toner_stable_count += 1
                    # Skip first few frames (attack transient ~100ms)
                    if self._toner_stable_count > ATTACK_SKIP_FRAMES:
                        frame = {
                            'note': note,  # concert pitch (raw from engine)
                            'freq': result.fundamental_freq,
                            'harmonics_db': [h.magnitude_db for h in result.harmonics],
                            'harmonic_cents': [h.cents_deviation for h in result.harmonics],
                            'signal_level': result.signal_level,
                            'spectral_centroid': result.spectral_centroid,
                        }
                        self._toner_free_accumulator.append(frame)
                else:
                    # Note changed — save accumulated if enough, start new
                    self._toner_free_save_micro_capture()
                    self._toner_stable_note = note
                    self._toner_stable_count = 1
                    self._toner_free_accumulator = []
            else:
                # Silence — save what we have
                self._toner_free_save_micro_capture()
                self._toner_stable_note = ""
                self._toner_stable_count = 0
                self._toner_free_accumulator = []

            # Count unique notes captured so far
            notes_so_far = set()
            if self._toner_active_session:
                for cap in self._toner_active_session.get('captures', []):
                    notes_so_far.add(cap.get('note', ''))

            if self._toner_stable_note:
                disp = self._toner_transpose_note(self._toner_stable_note)
                self._toner_capture_label.configure(
                    text=f"Free: {disp} "
                         f"({self._toner_stable_count} frames)")
            else:
                self._toner_capture_label.configure(
                    text="Free mode \u2014 play anything...")
            self._toner_capture_progress.configure(
                text=f"({len(notes_so_far)} notes captured)")
            return

        # --- CAL COUNTDOWN: 10-second prep before calibration starts ---
        if state == 'cal_countdown':
            elapsed = time.time() - self._toner_capture_start
            remaining = CAL_COUNTDOWN_S - elapsed
            if remaining > 0:
                self._toner_capture_label.configure(
                    text=f"Calibration starting in {remaining:.0f}s... get ready")
                self._toner_update_cal_prompt(f"{remaining:.0f}s", "get ready...")
                return
            # Countdown done — start calibration
            self._toner_capture_state = 'calibration'
            cal_notes = self._toner_cal_notes
            first_note = cal_notes[0]
            self._toner_capture_label.configure(
                text=f"Play {first_note} (1/{len(cal_notes)})")
            self._toner_capture_progress.configure(text="waiting...")
            self._toner_update_cal_prompt(first_note)
            return

        # --- CALIBRATION: guided note-by-note capture ---
        if state == 'calibration':
            cal_notes = self._toner_cal_notes
            if self._toner_cal_index >= len(cal_notes):
                self._toner_update_cal_prompt("")
                self._toner_stop_capture()
                messagebox.showinfo("Calibration Complete",
                    f"Calibration capture finished!\n"
                    f"{len(cal_notes)} notes recorded.")
                return

            # Notes are already in written pitch — display directly
            display_note = cal_notes[self._toner_cal_index]
            note_num = self._toner_cal_index + 1
            total = len(cal_notes)
            has_signal = self._toner_signal_above_threshold(result)

            if not self._toner_cal_recording:
                # Waiting for player to start this note
                if has_signal:
                    self._toner_cal_recording = True
                    self._toner_capture_start = time.time()
                    self._toner_capture_frames = []
                    self._toner_capture_label.configure(
                        text=f"Recording {display_note}... ({note_num}/{total})")
                    self._toner_update_cal_prompt(display_note, "recording...")
                else:
                    self._toner_capture_label.configure(
                        text=f"Play {display_note} ({note_num}/{total})")
                    self._toner_capture_progress.configure(text="waiting...")
                    self._toner_update_cal_prompt(display_note, f"({note_num}/{total})")
                return

            # Recording this note (skip attack transient ~100ms)
            elapsed = time.time() - self._toner_capture_start

            # Concert note for this calibration step (written note reverse-transposed)
            concert_note = reverse_transpose_note(
                display_note, self._toner_sax_var.get())

            if has_signal and elapsed > CAL_ATTACK_SETTLE_S:
                self._toner_capture_frames.append({
                    'note': concert_note,
                    'detected_note': result.fundamental_note,
                    'freq': result.fundamental_freq,
                    'harmonics_db': [h.magnitude_db for h in result.harmonics],
                    'harmonic_cents': [h.cents_deviation for h in result.harmonics],
                    'signal_level': result.signal_level,
                    'spectral_centroid': result.spectral_centroid,
                })

            remaining = CALIBRATION_DURATION_S - elapsed
            n_frames = len(self._toner_capture_frames)
            self._toner_capture_label.configure(
                text=f"Recording {display_note}... {remaining:.1f}s ({note_num}/{total})")
            self._toner_capture_progress.configure(
                text=f"({n_frames} frames)")
            self._toner_update_cal_prompt(display_note, f"recording {remaining:.0f}s")

            if remaining <= 0:
                # Save this note's capture
                if self._toner_capture_frames:
                    averaged = average_captures(self._toner_capture_frames)
                    avg_freq = sum(f['freq'] for f in self._toner_capture_frames) / len(self._toner_capture_frames)

                    from collections import Counter
                    detected = Counter(
                        self._toner_transpose_note(f['detected_note'])
                        for f in self._toner_capture_frames)
                    top_detected = detected.most_common(1)[0][0] if detected else "?"

                    # Average signal level and spectral centroid from frames
                    _frames = self._toner_capture_frames
                    _sigs = [f.get('signal_level', 0) for f in _frames if f.get('signal_level')]
                    _avg_sig = round(sum(_sigs) / len(_sigs), 3) if _sigs else 0.0
                    _scs = [f.get('spectral_centroid', 0) for f in _frames if f.get('spectral_centroid')]
                    _avg_sc = round(sum(_scs) / len(_scs), 1) if _scs else 0.0

                    capture_entry = {
                        'note': concert_note,  # concert pitch (storage)
                        'written_note': display_note,  # written pitch (for reference)
                        'detected_as': top_detected,
                        'fundamental_freq': round(avg_freq, 2),
                        'harmonics_db': [round(db, 2) for db in averaged['harmonics_db']],
                        'harmonic_cents': [round(c, 2) for c in averaged.get('harmonic_cents', [])],
                        'timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
                        'n_frames': len(self._toner_capture_frames),
                        'method': 'calibration',
                        'signal_level': _avg_sig,
                        'spectral_centroid': _avg_sc,
                    }

                    if self._toner_active_session is not None:
                        self._toner_active_session['captures'].append(capture_entry)
                        self._toner_save_active_session()
                        self._toner_check_rolloff_warning()

                # Transition to pause — wait for silence before next note
                self._toner_cal_recording = False
                self._toner_capture_frames = []
                self._toner_capture_state = 'cal_pause'
                self._toner_capture_label.configure(
                    text=f"Captured {display_note} \u2713  stop playing...")
                self._toner_capture_progress.configure(text="")
                self._toner_update_cal_prompt("\u2713", "stop playing...")
            return

        # --- CAL_PAUSE: wait for silence between calibration notes ---
        if state == 'cal_pause':
            has_signal = self._toner_signal_above_threshold(result)
            if not has_signal:
                # Silence detected — advance to next note
                self._toner_cal_index += 1
                self._toner_capture_state = 'calibration'

                cal_notes = self._toner_cal_notes
                if self._toner_cal_index < len(cal_notes):
                    next_note = cal_notes[self._toner_cal_index]
                    note_num = self._toner_cal_index + 1
                    total = len(cal_notes)
                    self._toner_capture_label.configure(
                        text=f"Play {next_note} ({note_num}/{total})")
                    self._toner_capture_progress.configure(text="waiting...")
                    self._toner_update_cal_prompt(next_note, f"({note_num}/{total})")
            return

    def _toner_free_save_micro_capture(self):
        """Save accumulated free-mode frames as a micro-capture (if enough)."""
        frames = self._toner_free_accumulator
        if len(frames) < FREE_MIN_FRAMES:
            return

        from collections import Counter
        note_counts = Counter(f['note'] for f in frames)
        dominant_note = note_counts.most_common(1)[0][0]
        note_frames = [f for f in frames if f['note'] == dominant_note]

        if len(note_frames) < FREE_MIN_FRAMES:
            return

        averaged = average_captures(note_frames)
        avg_freq = sum(f['freq'] for f in note_frames) / len(note_frames)

        # Average signal level and spectral centroid from accumulated frames
        sig_levels = [f.get('signal_level', 0) for f in note_frames if f.get('signal_level')]
        avg_signal = round(sum(sig_levels) / len(sig_levels), 3) if sig_levels else 0.0
        sc_vals = [f.get('spectral_centroid', 0) for f in note_frames if f.get('spectral_centroid')]
        avg_sc = round(sum(sc_vals) / len(sc_vals), 1) if sc_vals else 0.0

        capture_entry = {
            'note': dominant_note,
            'fundamental_freq': round(avg_freq, 2),
            'harmonics_db': [round(db, 2) for db in averaged['harmonics_db']],
            'harmonic_cents': [round(c, 2) for c in averaged.get('harmonic_cents', [])],
            'timestamp': time.strftime("%Y-%m-%d %H:%M:%S"),
            'n_frames': len(note_frames),
            'method': 'free',
            'signal_level': avg_signal,
            'spectral_centroid': avg_sc,
        }

        if self._toner_active_session is not None:
            self._toner_active_session['captures'].append(capture_entry)
            self._toner_save_active_session()
            self._toner_check_rolloff_warning()

    def _toner_check_rolloff_warning(self):
        """Check harmonic rolloff rate and warn if mic/room quality is poor."""
        if getattr(self, '_toner_rolloff_warned', False):
            return
        session = self._toner_active_session
        if not session:
            return
        captures = session.get('captures', [])
        if len(captures) < ROLLOFF_MIN_CAPTURES:
            return

        rates = []
        for cap in captures:
            r = compute_rolloff_rate(cap.get('harmonics_db', []))
            if r is not None:
                rates.append(r)
        if len(rates) < ROLLOFF_MIN_CAPTURES:
            return

        avg_rate = sum(rates) / len(rates)
        if avg_rate > ROLLOFF_WARN_THRESHOLD:
            self._toner_rolloff_warned = True
            messagebox.showwarning(
                "Recording Quality",
                f"Upper harmonics are dropping off steeply "
                f"({avg_rate:.1f} dB/harmonic).\n\n"
                f"This usually means the microphone is too far from "
                f"the bell, or a built-in laptop mic is being used.\n\n"
                f"For best results:\n"
                f"  \u2022  Use an external condenser mic (e.g. AT2020)\n"
                f"  \u2022  Place it 2\u20133 feet from the bell\n"
                f"  \u2022  See Help \u2192 User Guide for details",
                parent=self.root)

    def _toner_save_active_session(self):
        """Save the active session to the active profile.

        Saves a deep copy of the session data so that subsequent
        captures on a different profile don't mutate this profile's
        stored data through shared references.
        """
        import copy
        lib = self._toner_active_library
        prof_name = self._toner_active_profile
        if not lib or not prof_name:
            return
        if lib not in self._toner_profiles:
            return
        lib_profiles = self._toner_profiles[lib]
        if prof_name not in lib_profiles:
            return

        profile = lib_profiles[prof_name]
        sessions = profile.setdefault('sessions', [])
        session_date = self._toner_active_session.get('date', '')

        # Deep copy to prevent shared references between profiles
        session_copy = copy.deepcopy(self._toner_active_session)

        found = False
        for i, s in enumerate(sessions):
            if s.get('date') == session_date:
                sessions[i] = session_copy
                found = True
                break
        if not found:
            sessions.append(session_copy)

        save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)

    # ------------------------------------------------------------------
    # COMPARISON
    # ------------------------------------------------------------------

    def _toner_open_analyze_dialog(self):
        """Open dialog to select profiles/sessions for analysis."""
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
                "No profiles with captures to analyze.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Analyze Tone Profiles")
        dlg.geometry("600x520")
        dlg.resizable(True, True)
        dlg.minsize(500, 400)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select one or more profiles to analyze, compare, or overlay.",
                 bg=bg, fg=fg, font=("Helvetica", 10),
                 justify="left").pack(pady=(0, 5))

        # Filter controls — Row 1: horn identity
        filter_row1 = tk.Frame(frame, bg=bg)
        filter_row1.pack(fill="x", pady=(0, 2))

        # Collect unique values for filters
        def _unique(field):
            return sorted(set(p.get(field, '') for _, _, p in all_profiles
                              if p.get(field)))
        all_types = _unique('horn_type')
        all_makes = _unique('horn_make')
        all_models = _unique('horn_model')
        all_players = _unique('player')
        all_mpcs = _unique('mouthpiece')

        # Mic types from sessions
        all_mic_types = set()
        for _, _, p in all_profiles:
            for s in p.get('sessions', []):
                mt = s.get('mic_type', '')
                if mt:
                    all_mic_types.add(mt.capitalize())
        all_mic_types = sorted(all_mic_types)

        filter_type = tk.StringVar(value="All")
        filter_make = tk.StringVar(value="All")
        filter_model = tk.StringVar(value="All")
        filter_player = tk.StringVar(value="All")
        filter_mpc = tk.StringVar(value="All")
        filter_mic_type = tk.StringVar(value="All")

        def _add_filters(parent, items):
            for label, var, values in items:
                if values:
                    tk.Label(parent, text=label, bg=bg, fg=fg,
                             font=("Helvetica", 8)).pack(side="left", padx=(0, 2))
                    cb = ttk.Combobox(parent, textvariable=var,
                                       values=["All"] + values,
                                       state="readonly", width=12)
                    cb.pack(side="left", padx=(0, 6))
                    cb.bind("<<ComboboxSelected>>", lambda e: refresh_list())

        _add_filters(filter_row1, [
            ("Type:", filter_type, all_types),
            ("Make:", filter_make, all_makes),
            ("Model:", filter_model, all_models),
        ])

        # Filter controls — Row 2: setup + search
        filter_row2 = tk.Frame(frame, bg=bg)
        filter_row2.pack(fill="x", pady=(0, 5))

        _add_filters(filter_row2, [
            ("Player:", filter_player, all_players),
            ("Mpc:", filter_mpc, all_mpcs),
            ("Mic:", filter_mic_type, all_mic_types),
        ])

        # Text search
        tk.Label(filter_row2, text="Search:", bg=bg, fg=fg,
                 font=("Helvetica", 8)).pack(side="left", padx=(0, 2))
        search_var = tk.StringVar()
        search_entry = tk.Entry(filter_row2, textvariable=search_var, width=16)
        search_entry.pack(side="left", padx=(0, 2))
        search_entry.bind("<KeyRelease>", lambda e: refresh_list())
        tk.Button(filter_row2, text="\u00d7", font=("Helvetica", 8), width=2,
                  command=lambda: (search_var.set(""), refresh_list())).pack(side="left")

        # Scrollable checkbox list
        list_outer = tk.Frame(frame, bg=bg)
        list_outer.pack(fill="both", expand=True, pady=(0, 10))

        list_canvas = tk.Canvas(list_outer, bg=bg, highlightthickness=0,
                                height=250)
        scrollbar = tk.Scrollbar(list_outer, orient="vertical",
                                 command=list_canvas.yview)
        list_inner = tk.Frame(list_canvas, bg=bg)

        list_inner.bind("<Configure>",
            lambda e: list_canvas.configure(scrollregion=list_canvas.bbox("all")))
        list_canvas.create_window((0, 0), window=list_inner, anchor="nw")
        list_canvas.configure(yscrollcommand=scrollbar.set)

        list_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Mousewheel scrolling
        def _on_mousewheel(event):
            if sys.platform == 'darwin':
                list_canvas.yview_scroll(int(-1 * event.delta), "units")
            else:
                list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        if sys.platform == 'linux':
            dlg.bind('<Button-4>', lambda e: list_canvas.yview_scroll(-1, "units"))
            dlg.bind('<Button-5>', lambda e: list_canvas.yview_scroll(1, "units"))
        else:
            dlg.bind('<MouseWheel>', _on_mousewheel)

        # Each item in check_items is a dict describing what the checkbox
        # represents: either an entire profile or a single session.
        # {'type': 'profile', 'lib': str, 'name': str, 'profile': dict}
        # {'type': 'session', 'lib': str, 'name': str, 'profile': dict,
        #  'session': dict, 'date': str}
        check_items = []
        check_vars = []

        def refresh_list():
            nonlocal check_items, check_vars
            for w in list_inner.winfo_children():
                w.destroy()
            check_items = []
            check_vars = []
            ft = filter_type.get()
            fmk = filter_make.get()
            fmd = filter_model.get()
            fp = filter_player.get()
            fm = filter_mpc.get()
            fmt = filter_mic_type.get()
            search = search_var.get().strip().lower()
            for lib_name, prof_name, prof in all_profiles:
                if ft != "All" and prof.get('horn_type', '') != ft:
                    continue
                if fmk != "All" and prof.get('horn_make', '') != fmk:
                    continue
                if fmd != "All" and prof.get('horn_model', '') != fmd:
                    continue
                if fp != "All" and prof.get('player', '') != fp:
                    continue
                if fm != "All" and prof.get('mouthpiece', '') != fm:
                    continue
                if fmt != "All":
                    prof_mic_types = set(
                        s.get('mic_type', '').capitalize()
                        for s in prof.get('sessions', [])
                        if s.get('mic_type'))
                    if fmt not in prof_mic_types:
                        continue
                if search:
                    haystack = " ".join([
                        prof_name,
                        prof.get('horn_make', ''),
                        prof.get('horn_model', ''),
                        prof.get('serial', ''),
                        prof.get('player', ''),
                        prof.get('mouthpiece', ''),
                        prof.get('reed', ''),
                        prof.get('ligature', ''),
                        prof.get('room', ''),
                        prof.get('preamp', ''),
                        prof.get('notes', ''),
                    ]).lower()
                    # Also search session-level mic info
                    for s in prof.get('sessions', []):
                        haystack += " " + s.get('mic_type', '')
                        haystack += " " + s.get('mic_model', '')
                    if search not in haystack:
                        continue
                sessions = [s for s in prof.get('sessions', [])
                            if s.get('captures')]
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                if not notes:
                    continue
                status = f"{len(notes)} notes"
                if len(notes) >= MIN_PROFILE_NOTES:
                    status += " \u2713"

                # Profile-level checkbox (all sessions)
                var = tk.BooleanVar(value=False)
                check_vars.append(var)
                check_items.append({'type': 'profile', 'lib': lib_name,
                                    'name': prof_name, 'profile': prof})
                session_vars = []  # track child session vars

                def _make_profile_toggle(pvar, svars):
                    def _toggle():
                        val = pvar.get()
                        for sv in svars:
                            sv.set(val)
                    return _toggle

                tk.Checkbutton(
                    list_inner,
                    text=f"[{lib_name}] {prof_name}  ({status})",
                    variable=var, bg=bg, fg=fg,
                    selectcolor=bg, activebackground=bg,
                    anchor="w", font=("Helvetica", 10),
                    command=_make_profile_toggle(var, session_vars),
                ).pack(fill="x", anchor="w")

                # Session-level checkboxes (indented, only if 2+ sessions)
                if len(sessions) >= 2:
                    for sess in sessions:
                        date = sess.get('date', '?')[:10]
                        n_caps = len(sess.get('captures', []))
                        s_notes = set(c.get('note', '')
                                      for c in sess.get('captures', []))
                        svar = tk.BooleanVar(value=False)
                        check_vars.append(svar)
                        session_vars.append(svar)
                        check_items.append({
                            'type': 'session', 'lib': lib_name,
                            'name': prof_name, 'profile': prof,
                            'session': sess, 'date': date})
                        tk.Checkbutton(
                            list_inner,
                            text=f"{date}  ({n_caps} caps, "
                                 f"{len(s_notes)} notes)",
                            variable=svar, bg=bg, fg="#555555",
                            selectcolor=bg, activebackground=bg,
                            anchor="w", font=("Helvetica", 9),
                        ).pack(fill="x", anchor="w", padx=(25, 0))

        refresh_list()

        sel_frame = tk.Frame(frame, bg=bg)
        sel_frame.pack(fill="x", pady=(2, 5))
        def _select_all():
            for v in check_vars:
                v.set(True)
        def _deselect_all():
            for v in check_vars:
                v.set(False)
        tk.Button(sel_frame, text="Select All", font=("Helvetica", 8),
                  command=_select_all).pack(side="left", padx=(0, 4))
        tk.Button(sel_frame, text="Deselect All", font=("Helvetica", 8),
                  command=_deselect_all).pack(side="left")

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")

        def get_selected():
            """Return fingerprints for checked items (profiles or sessions)."""
            results = []
            for i, var in enumerate(check_vars):
                if var.get():
                    item = check_items[i]
                    sax = item['profile'].get('horn_type', 'Tenor')
                    if item['type'] == 'profile':
                        fp_val = compute_fingerprint(
                            item['profile'].get('sessions', []), sax)
                        fp_val['_name'] = item['name']
                        fp_val['_profile'] = item['profile']
                    else:
                        fp_val = compute_session_fingerprint(
                            item['session'], sax)
                        if not fp_val:
                            continue
                        fp_val['_name'] = (f"{item['name']} \u2014 "
                                           f"{item['date']}")
                        fp_val['_profile'] = item['profile']
                    results.append(fp_val)
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
            self._toner_show_delta_gauges(True)
            dlg.destroy()

        def analyze_selected():
            """Open analysis window for selected profiles/sessions."""
            selected = get_selected()
            if not selected:
                messagebox.showinfo("Select",
                    "Select at least one profile or session.", parent=dlg)
                return
            dlg.destroy()
            self._toner_show_analysis(selected)

        def clear_comparison():
            self._toner_comparison = None
            if hasattr(self, '_toner_compare_label'):
                self._toner_spectrum_canvas.itemconfigure(
                    self._toner_compare_label, text="")
            for g in self._toner_ghost_markers:
                self._toner_spectrum_canvas.itemconfigure(g, state="hidden")
            self._toner_show_delta_gauges(False)
            dlg.destroy()

        def average_selected():
            """Compute group average of selected profiles (profile-level only)."""
            profile_list = []
            for i, var in enumerate(check_vars):
                if var.get() and check_items[i]['type'] == 'profile':
                    item = check_items[i]
                    profile_list.append((item['name'], item['profile']))
            if len(profile_list) < 2:
                messagebox.showinfo("Select More",
                    "Select at least 2 profiles to average.\n"
                    "(Individual sessions are not included in group averages.)",
                    parent=dlg)
                return
            dlg.destroy()
            self._toner_show_group_report(profile_list)

        tk.Button(btn_frame, text="Analyze Selected",
                  command=analyze_selected).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Average Selected",
                  command=average_selected).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Overlay on Spectrum",
                  command=load_overlay).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Clear",
                  command=clear_comparison).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel",
                  command=dlg.destroy).pack(side="right")

    def _toner_show_analysis(self, fingerprints):
        """Show analysis window for one or more profiles/sessions.

        Single selection: profile view with chart, descriptors, per-note.
        Multiple selections: comparison with delta analysis.
        """
        is_single = len(fingerprints) == 1

        dlg = tk.Toplevel(self.root)
        if is_single:
            dlg.title(f"Analyze \u2014 {fingerprints[0].get('_name', '?')}")
        else:
            dlg.title("Analyze \u2014 Comparison")
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

        # Difference view toggle (only for 2-profile comparisons)
        chart_mode = tk.StringVar(value="overlay")
        diff_frame = tk.Frame(toggle_frame, bg=bg)
        if len(fingerprints) == 2:
            diff_frame.pack(side="right", padx=(10, 0))
            tk.Label(diff_frame, text="Chart:", bg=bg, fg=fg,
                     font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
            tk.Radiobutton(diff_frame, text="Overlay", variable=chart_mode,
                            value="overlay", bg=bg, fg=fg, selectcolor=bg,
                            font=("Helvetica", 9),
                            command=lambda: refresh_all()).pack(side="left")
            tk.Radiobutton(diff_frame, text="Difference", variable=chart_mode,
                            value="difference", bg=bg, fg=fg, selectcolor=bg,
                            font=("Helvetica", 9),
                            command=lambda: refresh_all()).pack(side="left")

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
            is_diff = (chart_mode.get() == "difference"
                       and len(fingerprints) == 2)

            margin_l, margin_r, margin_t, margin_b = 40, 10, 10, 25
            cw = w - margin_l - margin_r
            ch = h - margin_t - margin_b

            all_db = [d.get('harmonics_db', []) for d in data]
            max_h = max((len(db) for db in all_db), default=2)
            max_h = max(max_h, 2)

            if is_diff:
                # Difference mode: show delta between two profiles
                db_min, db_max = -20.0, 20.0
                # Grid with zero line
                for db in range(-20, 21, 5):
                    y = margin_t + ch * (1.0 - (db - db_min) / (db_max - db_min))
                    color = "#AAAAAA" if db == 0 else "#EEEEEE"
                    width = 2 if db == 0 else 1
                    chart_cv.create_line(margin_l, y, w - margin_r, y,
                                          fill=color, width=width)
                    chart_cv.create_text(margin_l - 4, y,
                                          text=f"{db:+d}" if db != 0 else "0",
                                          fill="#888888", font=("Helvetica", 7),
                                          anchor="e")

                # Axis label
                chart_cv.create_text(8, h // 2, text="\u0394 dB",
                                      fill="#888888", font=("Helvetica", 7),
                                      angle=90)
            else:
                db_min, db_max = -60.0, 0.0
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

            if is_diff:
                # Draw single difference curve
                db_a = data[0].get('harmonics_db', [])
                db_b = data[1].get('harmonics_db', [])
                n = min(len(db_a), len(db_b))
                if n >= 2:
                    points = []
                    for hi in range(n):
                        delta = db_a[hi] - db_b[hi]
                        x = margin_l + cw * hi / (max_h - 1)
                        clamped = max(db_min, min(db_max, delta))
                        y = margin_t + ch * (1.0 - (clamped - db_min) / (db_max - db_min))
                        points.extend([x, y])
                    if len(points) >= 4:
                        chart_cv.create_line(*points, fill="#FF5722",
                                              width=2, smooth=True)
                        # Color dots by direction
                        for j in range(0, len(points), 2):
                            hi = j // 2
                            delta = db_a[hi] - db_b[hi]
                            dot_color = "#4CAF50" if delta >= 0 else "#2196F3"
                            chart_cv.create_oval(
                                points[j] - 4, points[j + 1] - 4,
                                points[j] + 4, points[j + 1] + 4,
                                fill=dot_color, outline="")

                    # Legend
                    n1 = fingerprints[0]['_name'][:20]
                    n2 = fingerprints[1]['_name'][:20]
                    chart_cv.create_text(
                        margin_l + 5, margin_t + 8,
                        text=f"\u2191 {n1} stronger   \u2193 {n2} stronger",
                        fill="#666666", font=("Helvetica", 7), anchor="w")
            else:
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
        table_label = "Descriptors" if is_single else "Descriptor Comparison"
        table_frame = tk.LabelFrame(main, text=table_label, bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        table_frame.pack(fill="x", pady=(0, 6))
        table_inner = tk.Frame(table_frame, bg=bg)
        table_inner.pack(fill="x", padx=5, pady=5)

        desc_labels = [
            ("Complexity", "richness"),
            ("Warmth", "warmth"),
            ("Core Tone", "core_tone"),
            ("Even/Odd", "even_odd"),
            ("Rolloff Shape", "rolloff_shape"),
        ]

        # Rolloff rates and mic types for each profile (for table and mismatch check)
        _rolloff_rates = [fp.get('rolloff_rate') for fp in fingerprints]
        _mic_types = [fp.get('mic_type', '') for fp in fingerprints]

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

            # Rolloff rate and mic type rows (only in horn average view)
            if view_mode.get() == "average":
                row = tk.Frame(table_inner, bg=bg)
                row.pack(fill="x")
                tk.Label(row, text="Rec. Quality", width=14, bg=bg, fg=fg,
                         font=("Helvetica", 9), anchor="w").pack(side="left")
                for rate in _rolloff_rates:
                    if rate is None:
                        text = "\u2014"
                        val_fg = fg
                    else:
                        text = f"{rate:.1f} dB/H"
                        val_fg = "#880000" if rate > ROLLOFF_WARN_THRESHOLD else fg
                    tk.Label(row, text=text, width=12, bg=bg, fg=val_fg,
                             font=("Helvetica", 9), anchor="center").pack(side="left")

                row2 = tk.Frame(table_inner, bg=bg)
                row2.pack(fill="x")
                tk.Label(row2, text="Mic Type", width=14, bg=bg, fg=fg,
                         font=("Helvetica", 9), anchor="w").pack(side="left")
                for mt in _mic_types:
                    text = mt.capitalize() if mt else "\u2014"
                    tk.Label(row2, text=text, width=12, bg=bg, fg=fg,
                             font=("Helvetica", 9), anchor="center").pack(side="left")

        def rebuild_analysis():
            data = get_data_for_view()
            analysis_text.configure(state="normal")
            analysis_text.delete("1.0", tk.END)

            mode = view_mode.get()
            prefix = ""
            if mode == "per_note":
                prefix = f"For {note_var.get()}: "

            # Helper: session/capture context for a profile
            def _prof_ctx(fp):
                p = fp.get('_profile', {})
                n_sess = len([s for s in p.get('sessions', [])
                              if s.get('captures')]) if p else 0
                return f"{fp['capture_count']} caps, {n_sess} sess"

            lines = []

            if is_single:
                # Single profile/session view
                fp = fingerprints[0]
                d = data[0].get('descriptors', {})
                lines.append(f"{fp['_name']} ({_prof_ctx(fp)})")
                lines.append("")
                desc_parts = []
                for label, key in desc_labels:
                    val = d.get(key, 0)
                    if val > 0:
                        desc_parts.append(f"{label}: {val:.0%}")
                if desc_parts:
                    lines.append(prefix + ", ".join(desc_parts))
                rr = fp.get('rolloff_rate')
                mt = fp.get('mic_type', '')
                if rr is not None or mt:
                    parts = []
                    if mt:
                        parts.append(f"Mic: {mt}")
                    if rr is not None:
                        parts.append(f"Rolloff: {rr:.1f} dB/H")
                    lines.append(" | ".join(parts))

            else:
                # Context header for multi-profile
                ctx_parts = [f"{fp['_name']} ({_prof_ctx(fp)})"
                             for fp in fingerprints]
                lines.append(f"Comparing: {', '.join(ctx_parts)}")
                lines.append("")

            if len(fingerprints) == 2:
                n1 = fingerprints[0]['_name']
                n2 = fingerprints[1]['_name']
                da = data[0].get('descriptors', {})
                db_d = data[1].get('descriptors', {})
                h_a = data[0].get('harmonics_db', [])
                h_b = data[1].get('harmonics_db', [])

                if not da and not db_d:
                    lines.append(f"{prefix}No data for this note in either profile.")
                else:
                    # Descriptor deltas
                    delta_parts = []
                    for label, key in desc_labels:
                        va = da.get(key, 0)
                        vb = db_d.get(key, 0)
                        diff = va - vb
                        if abs(diff) > 0.05:
                            sign = "+" if diff > 0 else ""
                            delta_parts.append(f"{label.lower()} {sign}{diff:.0%}")
                    if delta_parts:
                        lines.append(f"{prefix}{n1} \u2192 {n2}: "
                                     + ", ".join(delta_parts))
                    else:
                        lines.append(f"{prefix}Descriptors are very similar.")

                    # Harmonic shift summary with component interpretation
                    n = min(len(h_a), len(h_b))
                    if n >= 2:
                        deltas = [(i, h_a[i] - h_b[i]) for i in range(1, n)]
                        biggest = max(deltas, key=lambda x: abs(x[1]))
                        if abs(biggest[1]) > 2.0:
                            big = [i + 1 for i, d in deltas if abs(d) > 2.0]
                            if len(big) >= 2:
                                hrange = f"H{big[0]}\u2013H{big[-1]}"
                            else:
                                hrange = f"H{big[0]}"
                            direction = n1 if biggest[1] > 0 else n2
                            lines.append(
                                f"{prefix}Biggest harmonic shifts at {hrange} "
                                f"({direction} stronger by up to "
                                f"{abs(biggest[1]):.1f} dB)")

                        # Component interpretation based on where shifts concentrate
                        low_shift = sum(abs(d) for i, d in deltas if i < 6) / max(1, min(5, n - 1))
                        mid_shift = sum(abs(d) for i, d in deltas if 6 <= i < 12) / max(1, len([d for i, d in deltas if 6 <= i < 12]))
                        hi_shift = sum(abs(d) for i, d in deltas if i >= 12) / max(1, len([d for i, d in deltas if i >= 12])) if n > 12 else 0

                        if mid_shift > 3.0 and mid_shift > low_shift * 1.5:
                            # Shifts concentrated in upper harmonics
                            if low_shift < 2.0:
                                lines.append(
                                    f"{prefix}Shifts concentrated in upper harmonics "
                                    "\u2014 consistent with neck or mouthpiece differences. "
                                    "Low harmonics are similar (body/bore is comparable).")
                            else:
                                lines.append(
                                    f"{prefix}Broadband shifts across H2\u2013H12 "
                                    "\u2014 consistent with a mouthpiece or player difference.")
                        elif low_shift > 2.0 and low_shift > mid_shift:
                            lines.append(
                                f"{prefix}Shifts concentrated in low harmonics (H2\u2013H4) "
                                "\u2014 this range is shaped primarily by the bore.")

                    # Player context
                    p1 = fingerprints[0].get('_profile', {}).get('player', '')
                    p2 = fingerprints[1].get('_profile', {}).get('player', '')
                    if p1 and p2:
                        if p1.lower() == p2.lower():
                            lines.append(
                                f"\nSame player ({p1}) \u2014 differences reflect "
                                "horn, neck, mouthpiece, or reed, not embouchure.")
                        else:
                            lines.append(
                                f"\nDifferent players ({p1} vs {p2}) \u2014 "
                                "cannot fully separate horn effect from "
                                "player/mouthpiece effect. Use Core Tone (H2\u2013H4) "
                                "for the most player-independent comparison.")
            elif len(fingerprints) > 2:
                for label, key in desc_labels:
                    values = [(fingerprints[i]['_name'],
                              d.get('descriptors', {}).get(key, 0))
                             for i, d in enumerate(data)]
                    values.sort(key=lambda x: x[1], reverse=True)
                    spread = values[0][1] - values[-1][1]
                    if spread > 0.1:
                        lines.append(
                            f"{prefix}{label} spread: {spread:.0%} "
                            f"({values[0][0]} highest, "
                            f"{values[-1][0]} lowest)")

            # Check for mic type and rolloff mismatches
            if mode == "average":
                # Mic type mismatch
                known_types = [mt for mt in _mic_types if mt]
                if len(set(known_types)) > 1:
                    type_list = ", ".join(
                        f"{fp['_name']}: {mt.capitalize()}"
                        for fp, mt in zip(fingerprints, _mic_types) if mt)
                    lines.append("")
                    lines.append(
                        f"\u26a0 Mic types differ ({type_list}). "
                        "Differences in complexity may partly reflect "
                        "the mic's frequency response rather than the horn.")

                # Rolloff mismatch (still useful even with same mic type)
                valid_rates = [r for r in _rolloff_rates if r is not None]
                if len(valid_rates) >= 2:
                    rate_spread = max(valid_rates) - min(valid_rates)
                    if rate_spread > 1.0:
                        lines.append("")
                        lines.append(
                            "\u26a0 Recording quality differs significantly "
                            f"(rolloff spread: {rate_spread:.1f} dB/H). "
                            "Complexity comparison may reflect mic/room "
                            "differences rather than horn differences.")

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
    # GROUP REPORT
    # ------------------------------------------------------------------

    def _toner_show_group_report(self, profile_list):
        """Show an aggregated report across multiple profiles.

        Args:
            profile_list: list of (name, profile_data) tuples
        """
        grp = compute_group_fingerprint(profile_list)
        if grp['profile_count'] == 0:
            messagebox.showinfo("No Data",
                "Selected profiles have no captures.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Group Average")
        dlg.geometry("720x650")
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"

        main = tk.Frame(dlg, bg=bg)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Header ---
        tk.Label(main, text=f"Group Average: {grp['profile_count']} profiles",
                 bg=bg, fg=fg,
                 font=("Helvetica", 13, "bold")).pack(anchor="w")
        names = [n for n, _ in profile_list]
        tk.Label(main, text=", ".join(names), bg=bg, fg=fg,
                 font=("Helvetica", 9), wraplength=680,
                 justify="left").pack(anchor="w")
        tk.Label(main, text=f"{grp['total_captures']} total captures",
                 bg=bg, fg=fg, font=("Helvetica", 9)).pack(anchor="w", pady=(0, 8))

        # --- Descriptors with ± stdev ---
        desc_frame = tk.LabelFrame(main, text="Group Descriptors",
                                    bg=bg, fg=fg,
                                    font=("Helvetica", 10, "bold"))
        desc_frame.pack(fill="x", pady=(0, 8))

        desc_row = tk.Frame(desc_frame, bg=bg)
        desc_row.pack(fill="x", padx=10, pady=8)

        gd = grp['descriptors']
        gs = grp['descriptor_stats']
        for label, key in [("Complexity", "richness"),
                           ("Warmth", "warmth")]:
            val = gd.get(key, 0)
            col = tk.Frame(desc_row, bg=bg)
            col.pack(side="left", expand=True)
            tk.Label(col, text=f"{val:.0%}", bg=bg, fg=fg,
                     font=("Helvetica", 16, "bold")).pack()
            if key in gs and gs[key]['n'] >= 2:
                sd = gs[key]['stdev']
                tk.Label(col, text=f"\u00b1{sd:.0%}", bg=bg, fg="#666666",
                         font=("Helvetica", 8)).pack()
            tk.Label(col, text=label, bg=bg, fg=fg,
                     font=("Helvetica", 8)).pack()

        # --- Harmonic chart: group average + individual profiles ---
        chart_frame = tk.LabelFrame(main, text="Harmonic Profiles",
                                     bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        chart_frame.pack(fill="both", expand=True, pady=(0, 8))

        chart_colors = ["#2196F3", "#FF5722", "#4CAF50", "#FF9800",
                        "#9C27B0", "#00BCD4", "#E91E63", "#8BC34A"]
        chart_cv = tk.Canvas(chart_frame, bg="white", highlightthickness=0,
                              height=150)
        chart_cv.pack(fill="both", expand=True, padx=5, pady=5)

        def draw_group_chart(event=None):
            chart_cv.delete("all")
            w = chart_cv.winfo_width()
            h = chart_cv.winfo_height()
            if w < 50 or h < 50:
                return

            margin_l, margin_r, margin_t, margin_b = 40, 100, 10, 25
            cw = w - margin_l - margin_r
            ch = h - margin_t - margin_b
            db_min, db_max = -60.0, 5.0

            # Grid
            for db in range(-60, 6, 10):
                y = margin_t + ch * (1.0 - (db - db_min) / (db_max - db_min))
                chart_cv.create_line(margin_l, y, w - margin_r, y,
                                      fill="#DDDDDD", width=1)
                chart_cv.create_text(margin_l - 4, y, text=f"{db}",
                                      fill="#888888", font=("Helvetica", 7),
                                      anchor="e")

            max_h = max((len(fp.get('harmonics_db', []))
                         for _, fp in grp['per_profile']),
                        default=0)
            max_h = max(max_h, len(grp.get('harmonics_db', [])))
            if max_h < 2:
                return

            for hi in range(max_h):
                x = margin_l + cw * hi / (max_h - 1)
                chart_cv.create_text(x, h - 5, text=f"H{hi+1}",
                                      fill="#888888", font=("Helvetica", 7))

            # Individual profiles (thin lines)
            for idx, (name, fp) in enumerate(grp['per_profile']):
                color = chart_colors[idx % len(chart_colors)]
                hdb = fp.get('harmonics_db', [])
                points = []
                for hi, db in enumerate(hdb):
                    x = margin_l + cw * hi / (max_h - 1)
                    clamped = max(db_min, min(db_max, db))
                    y = margin_t + ch * (1.0 - (clamped - db_min) / (db_max - db_min))
                    points.extend([x, y])
                if len(points) >= 4:
                    chart_cv.create_line(*points, fill=color, width=1,
                                          smooth=True, dash=(3, 3))

            # Group average (thick line)
            hdb = grp.get('harmonics_db', [])
            points = []
            for hi, db in enumerate(hdb):
                x = margin_l + cw * hi / (max_h - 1)
                clamped = max(db_min, min(db_max, db))
                y = margin_t + ch * (1.0 - (clamped - db_min) / (db_max - db_min))
                points.extend([x, y])
            if len(points) >= 4:
                chart_cv.create_line(*points, fill="#000000", width=3,
                                      smooth=True)

            # Legend
            legend_x = w - margin_r + 8
            legend_y = margin_t + 5
            chart_cv.create_rectangle(legend_x - 2, legend_y - 2,
                                       legend_x + 10, legend_y + 10,
                                       fill="#000000", outline="")
            chart_cv.create_text(legend_x + 14, legend_y + 4,
                                  text="Average", fill="#000000",
                                  font=("Helvetica", 7), anchor="w")
            legend_y += 16
            for idx, (name, _) in enumerate(grp['per_profile']):
                color = chart_colors[idx % len(chart_colors)]
                chart_cv.create_rectangle(legend_x - 2, legend_y - 2,
                                           legend_x + 10, legend_y + 10,
                                           fill=color, outline="")
                disp = name[:18] + "\u2026" if len(name) > 18 else name
                chart_cv.create_text(legend_x + 14, legend_y + 4,
                                      text=disp, fill=color,
                                      font=("Helvetica", 7), anchor="w")
                legend_y += 14

        chart_cv.bind("<Configure>", draw_group_chart)

        # --- Per-profile breakdown table ---
        tbl_frame = tk.LabelFrame(main, text="Per-Profile Breakdown",
                                   bg=bg, fg=fg,
                                   font=("Helvetica", 10, "bold"))
        tbl_frame.pack(fill="x", pady=(0, 8))

        tbl_cv = tk.Canvas(tbl_frame, bg=bg, highlightthickness=0, height=100)
        tbl_sb = tk.Scrollbar(tbl_frame, orient="vertical", command=tbl_cv.yview)
        tbl_inner = tk.Frame(tbl_cv, bg=bg)
        tbl_inner.bind("<Configure>",
            lambda e: tbl_cv.configure(scrollregion=tbl_cv.bbox("all")))
        tbl_cv.create_window((0, 0), window=tbl_inner, anchor="nw")
        tbl_cv.configure(yscrollcommand=tbl_sb.set)
        tbl_cv.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        tbl_sb.pack(side="right", fill="y")

        # Header
        thdr = tk.Frame(tbl_inner, bg=bg)
        thdr.pack(fill="x")
        for text, w in [("Profile", 20), ("Sess", 5), ("Caps", 5),
                        ("Notes", 5), ("Cmpx", 5), ("Warm", 5)]:
            tk.Label(thdr, text=text, width=w, bg=bg, fg=fg,
                     font=("Helvetica", 8, "bold")).pack(side="left")

        for name, fp in grp['per_profile']:
            trow = tk.Frame(tbl_inner, bg=bg)
            trow.pack(fill="x")
            # Find session count from profile_list
            n_sess = 0
            for pn, pd in profile_list:
                if pn == name:
                    n_sess = len([s for s in pd.get('sessions', [])
                                  if s.get('captures')])
                    break
            disp_name = name[:20] + "\u2026" if len(name) > 20 else name
            td = fp.get('descriptors', {})
            tk.Label(trow, text=disp_name, width=20, bg=bg, fg=fg,
                     font=("Helvetica", 8), anchor="w").pack(side="left")
            tk.Label(trow, text=str(n_sess), width=5, bg=bg, fg=fg,
                     font=("Helvetica", 8)).pack(side="left")
            tk.Label(trow, text=str(fp['capture_count']), width=5, bg=bg,
                     fg=fg, font=("Helvetica", 8)).pack(side="left")
            tk.Label(trow, text=str(fp['note_count']), width=5, bg=bg,
                     fg=fg, font=("Helvetica", 8)).pack(side="left")
            for key in ['richness', 'warmth']:
                tk.Label(trow, text=f"{td.get(key, 0):.0%}", width=5,
                         bg=bg, fg=fg,
                         font=("Helvetica", 8)).pack(side="left")

        # Context text
        ctx = (f"Group average across {grp['profile_count']} profiles "
               f"({grp['total_captures']} total captures). "
               f"\u00b1 values reflect variation across profiles, "
               f"not measurement noise within a profile.")
        tk.Label(main, text=ctx, bg=bg, fg="#888888",
                 font=("Helvetica", 7), wraplength=680,
                 justify="left").pack(anchor="w", pady=(0, 4))

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

        # Check for built-in mic on first start
        if not getattr(self, '_toner_mic_checked', False):
            self._toner_mic_checked = True
            is_builtin, dev_name = check_mic_quality()
            if is_builtin:
                messagebox.showwarning("Microphone Notice",
                    f"Your input device appears to be a built-in microphone "
                    f"(\"{dev_name}\").\n\n"
                    "Built-in mics often have poor low-frequency response, "
                    "which can cause inaccurate readings \u2014 especially "
                    "in the low register.\n\n"
                    "For best results, use a condenser mic such as the "
                    "Audio-Technica AT2020 USB.")

        device = self.settings.get("audio_input_device")
        if sys.platform == 'linux':
            device = None  # Linux/PulseAudio: device selection unreliable, use system default
        success, err = self._toner_engine.start(device=device)
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

        # Check if the engine died with an error
        if self._toner_engine and self._toner_engine.last_error:
            self._toner_show_stream_error(self._toner_engine.last_error)
            return

        if self._toner_engine and self._toner_engine.is_running:
            result = self._toner_engine.analyze()

            # Re-check after analyze() — stream may have just died
            if self._toner_engine.last_error:
                self._toner_show_stream_error(self._toner_engine.last_error)
                return

            # One-time spectral quality check after enough audio
            if (not self._toner_engine._mic_quality_warned and
                    self._toner_engine._spectral_check_frames >= SPECTRAL_CHECK_FRAME_COUNT):
                if self._toner_engine.check_spectral_quality():
                    self._toner_engine._mic_quality_warned = True
                    messagebox.showwarning("Microphone Quality",
                        "Your microphone appears to have poor low-frequency "
                        "response. Low register readings may be inaccurate.\n\n"
                        "For best results, use a condenser mic such as the "
                        "Audio-Technica AT2020 USB.")

            # Update session lamp
            self._toner_update_session_lamp()

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

            # Track low harmonic data condition
            if d.get('low_harmonic_data', False) and result.fundamental_freq > 0:
                self._toner_low_data_frames += 1
            else:
                self._toner_low_data_frames = 0

            # Show/hide low-data overlay on descriptor gauges
            if (self._toner_low_data_frames >= LOW_DATA_FRAME_THRESHOLD
                    and not self._toner_low_data_shown):
                self._toner_set_low_data_overlay(True)
            elif (self._toner_low_data_frames == 0
                    and self._toner_low_data_shown):
                self._toner_set_low_data_overlay(False)

            # Delta mode: compare live vs loaded baseline per-note
            if (self._toner_delta_mode.get() and self._toner_comparison
                    and result.fundamental_freq > 0):
                note = result.fundamental_note
                fp = self._toner_comparison
                pn = fp.get('per_note', {})
                baseline = pn.get(note) if note else None

                if baseline and baseline.get('descriptors'):
                    bd = baseline['descriptors']
                    bh = baseline.get('harmonics_db', [])
                    # Build live harmonics_db from result
                    live_h = [h.magnitude_db for h in result.harmonics
                              ] if result.harmonics else []
                    deltas = compute_delta_descriptors(
                        live_h, bh, d, bd)
                    if deltas:
                        # Existing gauges show delta (centered: 0.5 + delta/2)
                        r_delta = deltas['richness_delta']
                        w_delta = deltas['warmth_delta']
                        self._toner_update_gauge(
                            'richness', 0.5 + r_delta * 0.5)
                        self._toner_update_gauge(
                            'warmth', 0.5 + w_delta * 0.5)
                        # Comparison-only gauges
                        self._toner_update_delta_gauge(
                            'spectral_tilt', deltas.get('spectral_tilt'))
                        self._toner_update_delta_gauge(
                            'mid_harmonic', deltas.get('mid_harmonic'))
                    else:
                        self._toner_update_gauge('richness', 0.5)
                        self._toner_update_gauge('warmth', 0.5)
                        self._toner_update_delta_gauge('spectral_tilt', None)
                        self._toner_update_delta_gauge('mid_harmonic', None)
                else:
                    # No baseline data for this note
                    self._toner_update_gauge('richness', 0.5)
                    self._toner_update_gauge('warmth', 0.5)
                    self._toner_update_delta_gauge('spectral_tilt', None)
                    self._toner_update_delta_gauge('mid_harmonic', None)
            else:
                self._toner_update_gauge('richness', d.get('richness', 0.0))
                self._toner_update_gauge('warmth', d.get('warmth', 0.0))
                if self._toner_delta_mode.get():
                    # No signal — center delta gauges
                    self._toner_update_delta_gauge('spectral_tilt', None)
                    self._toner_update_delta_gauge('mid_harmonic', None)

            # Process capture if active
            self._toner_process_capture_frame(result)

        interval = FRAME_RATES.get(self._toner_fps_var.get(), 33)
        try:
            self._toner_anim_id = self.root.after(interval, self._toner_animate)
        except tk.TclError:
            pass  # Root window destroyed during shutdown

    def _toner_show_stream_error(self, error_msg):
        """Show audio stream error on the spectrum canvas with a retry option."""
        self._toner_running = False
        if hasattr(self, '_toner_spectrum_canvas'):
            c = self._toner_spectrum_canvas
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
                       lambda e: self._toner_retry())

    def _toner_retry(self):
        """Retry starting the toner after a stream error."""
        if self._toner_engine:
            self._toner_engine.last_error = None
        self._toner_start()

    # ------------------------------------------------------------------
    # SETTINGS SAVE/RESTORE
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # IMPORT / EXPORT
    # ------------------------------------------------------------------

    def _toner_import_audio_file(self):
        """Import an audio file (WAV) — full guided flow with profile setup."""
        from tkinter import filedialog

        if not self._toner_engine:
            messagebox.showerror("Error", "Toner engine not available.")
            return

        # Step 1: Select the audio file first
        filepath = filedialog.askopenfilename(
            title="Select Audio File to Analyze",
            filetypes=(("WAV files", "*.wav"), ("All files", "*.*"))
        )
        if not filepath:
            return

        filename = os.path.basename(filepath)

        # Step 2: Profile setup dialog (create or select, with source info)
        dlg = tk.Toplevel(self.root)
        dlg.title("Import Audio File")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text=f"Import: {filename}", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text="Select a profile for this recording, or create one.\n"
                 "Add notes about the source (who, when, where recorded).",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(pady=(0, 10))

        # Profile selector
        all_profiles = []
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict):
                continue
            for prof_name in lib_profiles:
                all_profiles.append((lib_name, prof_name))

        profile_frame = tk.Frame(frame, bg=bg)
        profile_frame.pack(fill="x", pady=(0, 5))

        use_existing = tk.BooleanVar(value=bool(all_profiles))

        if all_profiles:
            tk.Radiobutton(profile_frame, text="Existing profile:",
                           variable=use_existing, value=True,
                           bg=bg, fg=fg, font=("Helvetica", 10)).pack(
                               anchor="w")
            profile_var = tk.StringVar(
                value=f"[{all_profiles[0][0]}] {all_profiles[0][1]}" if all_profiles else "")
            profile_list = [f"[{lib}] {name}" for lib, name in all_profiles]
            prof_combo = ttk.Combobox(profile_frame, textvariable=profile_var,
                                      values=profile_list, state="readonly",
                                      width=40)
            prof_combo.pack(anchor="w", padx=(20, 0), pady=(0, 5))

        tk.Radiobutton(profile_frame, text="Create new profile...",
                       variable=use_existing, value=False,
                       bg=bg, fg=fg, font=("Helvetica", 10)).pack(anchor="w")

        # Source notes
        tk.Label(frame, text="Source notes:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(anchor="w", pady=(5, 2))
        source_notes_var = tk.StringVar(value=f"Imported from {filename}")
        tk.Entry(frame, textvariable=source_notes_var, width=45).pack(
            fill="x", pady=(0, 10))

        result_holder = [None]  # [filepath] if confirmed

        def do_import():
            source_notes = source_notes_var.get().strip()

            if use_existing.get() and all_profiles:
                # Use selected profile
                sel_text = profile_var.get()
                selected = None
                for lib, name in all_profiles:
                    if f"[{lib}] {name}" == sel_text:
                        selected = (lib, name)
                        break
                if not selected:
                    messagebox.showinfo("Select Profile",
                        "Select a profile from the list.", parent=dlg)
                    return

                self._toner_active_library = selected[0]
                self._toner_active_profile = selected[1]

                # Sync sax type
                prof = self._toner_profiles[selected[0]][selected[1]]
                sax_type = prof.get('horn_type', '')
                if sax_type and self._toner_engine:
                    self._toner_engine.set_sax_type(sax_type)
                    if hasattr(self, '_toner_sax_var'):
                        self._toner_sax_var.set(sax_type)

                # Create a file import session
                self._toner_active_session = {
                    'date': time.strftime("%Y-%m-%d %H:%M"),
                    'captures': [],
                    'source_notes': source_notes,
                    'method': 'file',
                    'mic_type': self.settings.get('mic_type', ''),
                    'mic_model': self.settings.get('mic_model', ''),
                }

                dlg.destroy()
                self._toner_do_file_import(filepath, source_notes)
            else:
                # Need to create new profile first
                dlg.destroy()
                # Store filepath for after profile creation
                self._toner_pending_file_import = (filepath, source_notes)
                self._toner_new_profile_flow_then_import()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(5, 0))
        tk.Button(btn_row, text="Import && Analyze",
                  command=do_import).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel",
                  command=dlg.destroy).pack(side="left")

    def _toner_new_profile_flow_then_import(self):
        """Create a new profile, then import the pending file."""
        # Reuse the existing new profile flow but override the callback
        dlg = tk.Toplevel(self.root)
        dlg.title("New Tone Profile")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Profile for Audio Import", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text="Describe the horn and setup heard in the recording.",
                 bg=bg, fg=fg, font=("Helvetica", 9)).pack(pady=(0, 10))

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
        ttk.Combobox(lib_row, textvariable=lib_var,
                      values=existing_libs, width=20).pack(
            side="left", fill="x", expand=True)

        add_field("Profile Name:", "name")
        add_field("Horn Type:", "horn_type", "Alto", widget_type="combo")
        add_field("Make:", "horn_make")
        add_field("Model:", "horn_model")
        add_field("Serial #:", "serial")
        add_field("Player:", "player")
        add_field("Mouthpiece:", "mouthpiece")
        add_field("Reed:", "reed")

        # Notes — multi-line
        notes_row = tk.Frame(frame, bg=bg)
        notes_row.pack(fill="x", pady=2)
        tk.Label(notes_row, text="Notes:", bg=bg, fg=fg, width=14,
                 anchor="e", font=("Helvetica", 10)).pack(side="left", padx=(0, 8), anchor="n")
        notes_text = tk.Text(notes_row, height=3, width=25, font=("Helvetica", 10), wrap="word")
        notes_text.pack(side="left", fill="x", expand=True)

        def save_and_import():
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if not name:
                messagebox.showwarning("Name Required",
                    "Please enter a profile name.", parent=dlg)
                return

            if lib not in self._toner_profiles:
                self._toner_profiles[lib] = {}
            if name in self._toner_profiles[lib]:
                messagebox.showwarning("Duplicate",
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
                'notes': notes_text.get("1.0", tk.END).strip(),
                'created': time.strftime("%Y-%m-%d"),
                'sessions': [],
            }
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
            self._toner_active_library = lib
            self._toner_active_profile = name
            self._toner_active_session = None  # Clear old session
            self._toner_update_profile_label()

            # Set sax type
            sax_type = fields["horn_type"].get()
            if self._toner_engine and sax_type:
                self._toner_engine.set_sax_type(sax_type)
                if hasattr(self, '_toner_sax_var'):
                    self._toner_sax_var.set(sax_type)

            dlg.destroy()

            # Now do the file import
            if hasattr(self, '_toner_pending_file_import'):
                filepath, source_notes = self._toner_pending_file_import
                del self._toner_pending_file_import
                self._toner_active_session = {
                    'date': time.strftime("%Y-%m-%d %H:%M"),
                    'captures': [],
                    'source_notes': source_notes,
                    'method': 'file',
                }
                self._toner_do_file_import(filepath, source_notes)

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Create && Import",
                  command=save_and_import).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel",
                  command=dlg.destroy).pack(side="left")

    def _toner_do_file_import(self, filepath, source_notes=""):
        """Actually run the file analysis and save captures."""
        filename = os.path.basename(filepath)

        # Show progress
        self._toner_capture_frame.pack(fill="x", padx=5,
            before=self._toner_main_frame.winfo_children()[-1])
        self._toner_capture_label.configure(
            text=f"Analyzing {filename}...")
        self._toner_capture_progress.configure(text="")
        self.root.update_idletasks()

        def on_progress(current, total):
            pct = current * 100 // total
            self._toner_capture_progress.configure(text=f"{pct}%")
            self.root.update_idletasks()

        try:
            captures = analyze_audio_file(filepath, self._toner_engine,
                                          progress_cb=on_progress)
        except Exception as e:
            self._toner_capture_frame.pack_forget()
            messagebox.showerror("Import Error", f"Could not analyze file:\n{e}")
            return

        self._toner_capture_frame.pack_forget()

        if not captures:
            messagebox.showinfo("No Data",
                "No stable note segments found in the file.\n"
                "The file may be too short, too quiet, or contain no "
                "sustained tones.")
            return

        for cap in captures:
            # Engine returns concert pitch — store as-is
            cap['timestamp'] = time.strftime("%Y-%m-%d %H:%M:%S")
            cap['source_file'] = filename
            if source_notes:
                cap['source_notes'] = source_notes

        self._toner_active_session['captures'].extend(captures)
        # Compute and store rolloff for file import session
        rates = [compute_rolloff_rate(c.get('harmonics_db', []))
                 for c in self._toner_active_session['captures']]
        rates = [r for r in rates if r is not None]
        if rates:
            self._toner_active_session['rolloff_rate'] = round(
                sum(rates) / len(rates), 2)
        self._toner_save_active_session()

        notes = set(c['note'] for c in captures)
        total_notes = set()
        for cap in self._toner_active_session['captures']:
            total_notes.add(cap.get('note', ''))

        # Transpose note names for display
        display_notes = sorted(
            self._toner_transpose_note(n) for n in notes)
        messagebox.showinfo("File Imported",
            f"Extracted {len(captures)} note segments from '{filename}'.\n"
            f"Notes found: {', '.join(display_notes)}\n\n"
            f"Profile now has {len(total_notes)} unique notes total.")

    def _toner_export_profiles(self):
        """Export selected tone profiles to a JSON file."""
        from tkinter import filedialog

        # Build list of all profiles
        all_profiles = []
        for lib_name, lib_profiles in self._toner_profiles.items():
            if not isinstance(lib_profiles, dict):
                continue
            for prof_name, prof_data in lib_profiles.items():
                all_profiles.append((lib_name, prof_name, prof_data))

        if not all_profiles:
            messagebox.showinfo("Nothing to Export", "No tone profiles to export.")
            return

        # Selection dialog
        dlg = tk.Toplevel(self.root)
        dlg.title("Export Tone Profiles")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select profiles to export:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(pady=(0, 5))

        # Checkboxes
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        check_vars = []
        for lib_name, prof_name, prof_data in all_profiles:
            sessions = prof_data.get('sessions', [])
            caps = sum(len(s.get('captures', [])) for s in sessions)
            var = tk.BooleanVar(value=True)
            tk.Checkbutton(list_frame,
                text=f"[{lib_name}] {prof_name} ({caps} captures)",
                variable=var, bg=bg, font=("Helvetica", 9),
                anchor="w").pack(fill="x")
            check_vars.append(var)

        # Select all / none
        sel_frame = tk.Frame(frame, bg=bg)
        sel_frame.pack(fill="x", pady=(0, 10))
        tk.Button(sel_frame, text="All", font=("Helvetica", 8),
                  command=lambda: [v.set(True) for v in check_vars]).pack(
                      side="left", padx=(0, 5))
        tk.Button(sel_frame, text="None", font=("Helvetica", 8),
                  command=lambda: [v.set(False) for v in check_vars]).pack(
                      side="left")

        def do_export():
            # Build export data from selected profiles
            export = {}
            count = 0
            for (lib, name, data), var in zip(all_profiles, check_vars):
                if var.get():
                    if lib not in export:
                        export[lib] = {}
                    export[lib][name] = data
                    count += 1

            if not export:
                messagebox.showinfo("Nothing Selected",
                    "Select at least one profile to export.", parent=dlg)
                return

            dlg.destroy()

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
                    json.dump(export, f, indent=2)
                messagebox.showinfo("Export Successful",
                    f"Exported {count} profiles to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Could not export:\n{e}")

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Export", command=do_export).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(
            side="left")

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
            "sax_type": self._toner_sax_var.get() if hasattr(self, '_toner_sax_var') else "Alto",
            "concert_pitch": self._toner_concert_pitch.get() if hasattr(self, '_toner_concert_pitch') else False,
        }
        # Also save any pending session
        if self._toner_active_session and self._toner_active_profile:
            save_tone_profiles(self._toner_profiles, TONE_PROFILES_FILE)
