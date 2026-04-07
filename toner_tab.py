"""
Toner tab mixin for Stohrer Sax Shop Companion.

Harmonic tone analyzer for saxophone. Shows a live spectrum analyzer
(full FFT or harmonic-only bars) on the left and VU-style tone
descriptor gauges on the right. Auto-detects fundamental pitch and
analyzes harmonic content for real-time tone quality feedback.

Includes a tone preset system: guided capture sessions build up
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
        MIN_PRESET_NOTES, SAX_TYPES, MIN_FUNDAMENTAL_HZ,
        SAX_TRANSPOSITIONS, FREE_STABLE_FRAMES, FREE_MIN_FRAMES,
        ATTACK_SKIP_FRAMES, CALIBRATION_NOTES, CALIBRATION_DURATION_S,
        DEFAULT_LIBRARY, average_captures, compute_fingerprint,
        compute_session_fingerprint, compute_session_variation,
        compute_group_fingerprint, compute_rolloff_rate,

        ROLLOFF_WARN_THRESHOLD, ROLLOFF_MIN_CAPTURES,
        get_rolloff_threshold,
        load_tone_presets, save_tone_presets,
        analyze_audio_file, check_mic_quality,
        transpose_note, reverse_transpose_note, note_to_freq,
        compute_population_stats, percentile_rank,
        find_out_of_range_notes, note_name_to_midi,
    )
    _TONER_IMPORTS_OK = True
except ImportError:
    _TONER_IMPORTS_OK = False
    AUDIO_AVAILABLE = False
    PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

try:
    from config import TONER_DATA_FILE, save_settings
except ImportError:
    TONER_DATA_FILE = "toner_data.json"


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
GAUGE_INTONATION_NEEDLE_OFFSET = 16
GAUGE_INTONATION_DAMPING = 0.18       # lerp factor per frame for intonation needle

# Animation / display thresholds
IN_TUNE_CENTS_THRESHOLD = 4.0         # cents — below this, "in tune" lamp lights
INTONATION_RANGE_CENTS = 50.0         # cents — full deflection of intonation gauge
SPECTRUM_MAX_FREQ = 8000.0            # Hz — right edge of spectrum display
SPECTRAL_CHECK_FRAME_COUNT = 25       # frames before mic quality check fires
FRAME_DURATION_S = 0.033              # approximate duration of one frame at 30 fps

# Capture state machine
CAL_COUNTDOWN_S = 10.0                # seconds of countdown before calibration starts
CAL_ATTACK_SETTLE_S = 0.1             # seconds of attack transient to skip in calibration
DEFAULT_STABLE_THRESHOLD = 25         # frames — initial stable-note threshold
PRESET_NAME_MAX_DISPLAY = 20         # characters before truncating preset name


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


def _format_preset_info(p):
    """Format a preset dict into a readable info string."""
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
    mic_parts = []
    if p.get('mic_type'):
        mic_parts.append(p['mic_type'].capitalize())
    if p.get('mic_model'):
        mic_parts.append(p['mic_model'])
    if mic_parts:
        info += f"\nMic: {' — '.join(mic_parts)}"
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
    if 0 < len(notes) < MIN_PRESET_NOTES:
        info += f" (need {MIN_PRESET_NOTES - len(notes)} more)"
    elif len(notes) >= MIN_PRESET_NOTES:
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
        # Profile system state — nested: {library: {preset_name: data}}
        self._toner_presets = {}
        self._toner_active_library = None   # Library name
        self._toner_active_preset = None   # Profile name within library
        self._toner_active_session = None   # Current capture session dict
        self._toner_captured_notes = set()  # Cached note set (avoids per-frame iteration)
        self._toner_save_pending = False    # Deferred save flag
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

        self._toner_presets = load_tone_presets(TONER_DATA_FILE)

        # Migrate: backfill mic_type/mic_model from sessions into presets
        _migrated = False
        for _lib in self._toner_presets.values():
            if not isinstance(_lib, dict):
                continue
            for _preset in _lib.values():
                if not isinstance(_preset, dict) or _preset.get('mic_type'):
                    continue
                # Find most common mic_type from sessions
                from collections import Counter
                _mics = [s.get('mic_type', '') for s in _preset.get('sessions', [])
                         if s.get('mic_type')]
                if _mics:
                    _preset['mic_type'] = Counter(_mics).most_common(1)[0][0]
                    _models = [s.get('mic_model', '') for s in _preset.get('sessions', [])
                               if s.get('mic_model')]
                    _preset['mic_model'] = _models[0] if _models else ''
                    _migrated = True
        if _migrated:
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)

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
            gauge_frame, bg=bg, highlightthickness=0,
            width=GAUGE_WIDTH, height=GAUGE_HEIGHT)
        self._toner_intonation_canvas._dark_canvas = True
        self._toner_intonation_canvas.grid(row=gauge_row, column=0, pady=(2, 0))
        self._toner_intonation_gauge = self._toner_build_intonation_gauge(
            self._toner_intonation_canvas)
        self._toner_smooth_cents = 0.0

        # Note display + in-tune lamp to the right of intonation gauge
        note_frame = tk.Frame(gauge_frame, bg=bg)
        note_frame._skip_theme = True
        note_frame.grid(row=gauge_row, column=1, sticky="w", padx=(4, 2))

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

        # Live descriptor gauges (Pure↔Complex, Thin↔Warm) were removed
        # 2026-04-06 — current data is too noisy to make absolute single-
        # preset readouts meaningful (mic position alone shifts complexity
        # 10-20%; mouthpiece dominates the signal). Comparison descriptors
        # still live in the Analyze tool where deltas cancel the confounders.

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

        # Sandbox checkbox (below sax selector, visible only when enabled)
        self._toner_sandbox_var = tk.BooleanVar(value=False)

        def _on_sandbox_toggled(*args):
            if self._toner_sandbox_var.get():
                self._toner_concert_pitch.set(True)
                self._toner_update_pitch_mode_label()

        self._toner_sandbox_cb = tk.Checkbutton(
            sax_inner, text="sandbox (concert pitch)",
            variable=self._toner_sandbox_var,
            command=_on_sandbox_toggled,
            bg=ctrl_bg, fg="#CC8800", activebackground=ctrl_bg,
            activeforeground="#CC8800", selectcolor=ctrl_bg,
            font=("Helvetica", 7, "bold"))
        self._toner_sandbox_cb._skip_theme = True
        if self.settings.get("toner_sandbox_enabled"):
            self._toner_sandbox_cb.pack(pady=(2, 0))

        # Lock sax selector when preset is loaded
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
                 font=("Helvetica", 7, "bold")).pack(side="left", padx=(3, 6))

        # Mic status indicator (shows loaded preset's mic)
        self._toner_mic_label = tk.Label(
            right_col, text="no preset", bg=ctrl_bg, fg="#666666",
            font=("Helvetica", 7))
        self._toner_mic_label._skip_theme = True
        self._toner_mic_label.pack(side="left", padx=(0, 10))

        # Profile indicator + Load/Unload toggle
        self._toner_preset_label = tk.Label(
            right_col, text="no preset", bg=ctrl_bg, fg="#666666",
            font=("Helvetica", 8))
        self._toner_preset_label.pack(side="left", padx=(0, 4))

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

    def _toner_set_overlay(self, fingerprint, name):
        """Set a fingerprint as the live spectrum ghost overlay."""
        self._toner_comparison = fingerprint
        if hasattr(self, '_toner_compare_label'):
            note_count = fingerprint.get('note_count', 0)
            captures = fingerprint.get('capture_count',
                                       fingerprint.get('total_captures', 0))
            detail = f"{note_count} notes" if note_count else f"{captures} captures"
            self._toner_spectrum_canvas.itemconfigure(
                self._toner_compare_label,
                text=f"Overlay: {name} ({detail})")

    def _toner_clear_overlay(self):
        """Remove the live spectrum ghost overlay."""
        self._toner_comparison = None
        if hasattr(self, '_toner_compare_label'):
            self._toner_spectrum_canvas.itemconfigure(
                self._toner_compare_label, text="")
        for g in self._toner_ghost_markers:
            self._toner_spectrum_canvas.itemconfigure(g, state="hidden")

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

        # Ghost overlay markers (for comparison preset)
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
        """Lock or unlock the sax selector (locked when preset is loaded)."""
        if hasattr(self, '_toner_sax_scale_widget'):
            state = "disabled" if locked else "normal"
            self._toner_sax_scale_widget.configure(state=state)

    def _toner_open_settings(self):
        """Open consolidated toner settings dialog."""
        from config import get_input_devices, save_settings
        from tkinter import filedialog

        dlg = tk.Toplevel(self.root)
        dlg.title("Tone Analyzer Settings")
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"

        notebook = ttk.Notebook(dlg)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        # ==================================================================
        # GENERAL TAB
        # ==================================================================
        gen_frame = tk.Frame(notebook, bg=bg, padx=15, pady=10)
        notebook.add(gen_frame, text="General")

        # --- Input Device ---
        input_frame = tk.LabelFrame(gen_frame, text="Input Device", bg=bg,
                                     fg=fg, padx=10, pady=8)
        input_frame.pack(fill="x", pady=(0, 8))

        devices = get_input_devices()
        dev_indices = [None]
        dev_names = ["System Default"]
        mic_var = tk.StringVar(value="System Default")

        if sys.platform == 'linux':
            tk.Label(input_frame, text="System Default (set in system audio settings)",
                     bg=bg, fg="#888888", font=("Helvetica", 9)).pack(anchor="w")
        elif devices:
            dev_names += [name for _, name in devices]
            dev_indices += [idx for idx, _ in devices]

            current_dev = self.settings.get("audio_input_device")
            if current_dev is not None:
                for idx, name in devices:
                    if idx == current_dev:
                        mic_var.set(name)
                        break

            listbox = tk.Listbox(input_frame, height=min(6, len(dev_names)),
                                  width=45, font=("Helvetica", 9))
            listbox.pack(fill="x", pady=(0, 6))
            for name in dev_names:
                listbox.insert(tk.END, name)
            current_idx = 0
            if current_dev is not None:
                for i, (idx, _) in enumerate(devices):
                    if idx == current_dev:
                        current_idx = i + 1
                        break
            listbox.selection_set(current_idx)
            listbox.see(current_idx)
        else:
            tk.Label(input_frame, text="No audio input devices found.",
                     bg=bg, fg="#888888", font=("Helvetica", 9)).pack(anchor="w")

        tk.Label(input_frame, text="Mic type and model are set per preset\n"
                 "in File \u2192 Presets.",
                 bg=bg, fg="#888888", font=("Helvetica", 8)).pack(anchor="w", pady=(4, 0))

        # --- Recording ---
        rec_frame = tk.LabelFrame(gen_frame, text="Recording", bg=bg,
                                   fg=fg, padx=10, pady=8)
        rec_frame.pack(fill="x", pady=(0, 8))

        record_var = tk.BooleanVar(value=self.settings.get("toner_record_wav", True))
        wav_warning = tk.Label(rec_frame,
            text="\u26a0 Without WAV recording, post-capture analysis "
                 "will be roughly half as accurate.",
            bg=bg, fg="#CC6600", font=("Helvetica", 8),
            wraplength=380, justify="left")

        def _on_record_toggled():
            if record_var.get():
                wav_warning.pack_forget()
            else:
                wav_warning.pack(anchor="w", padx=(16, 0), pady=(0, 2))

        tk.Checkbutton(rec_frame, text="Record WAV during capture",
                       variable=record_var, bg=bg, fg=fg,
                       command=_on_record_toggled,
                       font=("Helvetica", 9)).pack(anchor="w")
        if not record_var.get():
            wav_warning.pack(anchor="w", padx=(16, 0), pady=(0, 2))

        auto_delete_var = tk.BooleanVar(value=self.settings.get("toner_wav_auto_delete", False))
        tk.Checkbutton(rec_frame, text="Delete WAV after analysis",
                       variable=auto_delete_var, bg=bg, fg=fg,
                       font=("Helvetica", 9)).pack(anchor="w", padx=(16, 0))

        folder_frame = tk.Frame(rec_frame, bg=bg)
        folder_frame.pack(fill="x", pady=(2, 0))
        current_dir = self.settings.get("toner_recording_dir", "") or self._toner_get_recording_dir()
        folder_label = tk.Label(folder_frame, text=current_dir, bg=bg, fg="#666666",
                                font=("Helvetica", 8), anchor="w")
        def choose_folder():
            chosen = filedialog.askdirectory(
                title="Choose Recording Folder",
                initialdir=current_dir if os.path.isdir(current_dir) else os.path.expanduser("~"),
                parent=dlg)
            if chosen:
                folder_label.configure(text=chosen)
        tk.Button(folder_frame, text="Folder...", font=("Helvetica", 8),
                  command=choose_folder).pack(side="left", padx=(0, 5))
        folder_label.pack(side="left", fill="x")

        # --- Pitch ---
        pitch_frame = tk.LabelFrame(gen_frame, text="Pitch", bg=bg,
                                     fg=fg, padx=10, pady=8)
        pitch_frame.pack(fill="x", pady=(0, 8))

        pitch_row = tk.Frame(pitch_frame, bg=bg)
        pitch_row.pack(fill="x", pady=(0, 4))
        tk.Label(pitch_row, text="Reference pitch  A =", bg=bg, fg=fg,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 5))
        ref_pitch_var = tk.DoubleVar(value=self._toner_pitch_var.get())
        tk.Entry(pitch_row, textvariable=ref_pitch_var, width=6,
                 font=("Helvetica", 9)).pack(side="left")
        tk.Label(pitch_row, text="Hz", bg=bg, fg=fg,
                 font=("Helvetica", 9)).pack(side="left", padx=(3, 0))

        display_pitch_var = tk.StringVar(
            value="concert" if self._toner_concert_pitch.get() else "written")
        tk.Radiobutton(pitch_frame, text="Written pitch (what the player fingers)",
                       variable=display_pitch_var, value="written", bg=bg, fg=fg,
                       font=("Helvetica", 9)).pack(anchor="w")
        tk.Radiobutton(pitch_frame, text="Concert pitch (actual sounding frequency)",
                       variable=display_pitch_var, value="concert", bg=bg, fg=fg,
                       font=("Helvetica", 9)).pack(anchor="w")

        # ==================================================================
        # ANALYSIS TAB
        # ==================================================================
        ana_frame = tk.Frame(notebook, bg=bg, padx=15, pady=10)
        notebook.add(ana_frame, text="Analysis")

        # --- Preset Fields ---
        pf_frame = tk.LabelFrame(ana_frame, text="Preset Fields", bg=bg,
                                  fg=fg, padx=10, pady=8)
        pf_frame.pack(fill="x", pady=(0, 8))

        tk.Label(pf_frame, text="Show these optional fields when creating presets.",
                 bg=bg, fg=fg, font=("Helvetica", 9)).pack(anchor="w", pady=(0, 4))

        field_labels = [
            ("serial", "Serial #"),
            ("reed", "Reed"),
            ("ligature", "Ligature"),
            ("mic_position", "Mic Position"),
            ("room", "Room / Environment"),
            ("preamp", "Preamp / Interface"),
            ("notes", "Notes"),
        ]
        vis = self.settings.get("visible_preset_fields", {})
        field_vars = {}
        for key, label in field_labels:
            var = tk.BooleanVar(value=vis.get(key, False))
            tk.Checkbutton(pf_frame, text=label, variable=var, bg=bg, fg=fg,
                           font=("Helvetica", 9)).pack(anchor="w")
            field_vars[key] = var

        # Easter egg
        _ns_var = tk.BooleanVar(value=False)
        _ns_cb = tk.Checkbutton(pf_frame, text="Heavy Mass Neck Screw",
                                variable=_ns_var, bg=bg, fg=fg,
                                font=("Helvetica", 9))
        _ns_cb.pack(anchor="w")
        _ns_msgs = ["No.", "Nope.", "Uh uh.", "Jabroni: I refuse.",
                     "Forget it.", "Stop.", "Dude."]
        _ns_idx = [0]

        def _ns_remove():
            pw = tk.Toplevel(dlg)
            pw.title("Processing...")
            pw.resizable(False, False)
            pw.transient(dlg)
            pw.grab_set()
            pf = tk.Frame(pw, bg=bg, padx=20, pady=15)
            pf.pack(fill="both", expand=True)
            tk.Label(pf, text="Removing option...", bg=bg,
                     font=("Helvetica", 10)).pack(pady=(0, 8))
            pbar = ttk.Progressbar(pf, orient="horizontal",
                                   length=250, mode="determinate")
            pbar.pack()

            def _tick(val):
                pbar['value'] = val
                if val < 100:
                    if val < 70:
                        delay = 40
                    elif val < 90:
                        delay = 200
                    elif val < 99:
                        delay = 500
                    else:
                        delay = 3000
                    pw.after(delay, _tick, val + 1)
                else:
                    pw.after(300, lambda: (pw.destroy(), _ns_cb.pack_forget()))
            _tick(0)

        def _ns_click():
            _ns_var.set(False)
            if _ns_idx[0] < len(_ns_msgs):
                messagebox.showinfo("", _ns_msgs[_ns_idx[0]], parent=dlg)
                _ns_idx[0] += 1
            if _ns_idx[0] >= len(_ns_msgs):
                _ns_remove()
        _ns_cb.configure(command=_ns_click)

        # --- Analysis Descriptors ---
        desc_frame = tk.LabelFrame(ana_frame, text="Analysis Descriptors", bg=bg,
                                    fg=fg, padx=10, pady=8)
        desc_frame.pack(fill="x", pady=(0, 8))

        tk.Label(desc_frame, text="Choose which descriptors appear in the Analyze tool.",
                 bg=bg, fg=fg, font=("Helvetica", 9)).pack(anchor="w", pady=(0, 4))

        toner_settings = self.settings.get("toner_settings", {})
        analysis_desc = toner_settings.get("analysis_descriptors", {})

        descriptor_info = [
            ("richness", "Harmonic Complexity", None,
             "Spectral flatness \u2014 how evenly energy is spread "
             "across the harmonic series.\n\n"
             "Higher values = many harmonics at similar strength "
             "(complex, rich tone).\n"
             "Lower values = energy concentrated in fewer harmonics "
             "(purer, simpler tone).\n\n"
             "Note: Ribbon mics roll off upper harmonics and cannot "
             "accurately measure this descriptor."),
            ("warmth", "Warmth", None,
             "Strength of the second harmonic (H2, one octave above "
             "the fundamental) relative to the fundamental.\n\n"
             "Higher values = strong octave harmonic (warm, full tone).\n"
             "Lower values = weak octave harmonic (thinner tone)."),
            ("even_odd", "Even/Odd Harmonic Balance", "beta",
             "Measures the ratio of even harmonic energy (H2, H4, H6...) "
             "to odd harmonic energy (H3, H5, H7...).\n\n"
             "Higher values = more even-dominant (rounder quality).\n"
             "Lower values = more odd-dominant (edgier, hollower quality).\n\n"
             "This descriptor shows good differentiation across horns but "
             "is not an established measurement in saxophone acoustics. "
             "Treat it as experimental."),
            ("rolloff_shape", "Rolloff Shape", "beta",
             "Measures how smoothly harmonics decrease in strength "
             "vs having bumps or peaks in the harmonic series.\n\n"
             "Higher values = more spectral peaks that stick out.\n"
             "Lower values = smooth, even rolloff.\n\n"
             "This is a signal processing metric, not an established "
             "measurement in saxophone acoustics. Treat it as experimental."),
        ]

        desc_vars = {}
        for key, label, badge, info_text in descriptor_info:
            row = tk.Frame(desc_frame, bg=bg)
            row.pack(fill="x", pady=2)

            var = tk.BooleanVar(value=analysis_desc.get(key, key in ("richness", "warmth", "even_odd")))
            desc_vars[key] = var
            tk.Checkbutton(row, variable=var, bg=bg).pack(side="left")

            display = label
            if badge:
                display += f"  [{badge}]"
            tk.Label(row, text=display, bg=bg, fg=fg,
                     font=("Helvetica", 9)).pack(side="left")

            def make_info_cmd(title=label, text=info_text):
                return lambda: messagebox.showinfo(title, text, parent=dlg)
            tk.Button(row, text="?", width=2, font=("Helvetica", 8),
                      command=make_info_cmd()).pack(side="left", padx=(6, 0))

        # --- Sandbox Mode ---
        sandbox_frame = tk.LabelFrame(ana_frame, text="Sandbox Mode", bg=bg,
                                       fg=fg, padx=10, pady=8)
        sandbox_frame.pack(fill="x", pady=(0, 8))

        sandbox_var = tk.BooleanVar(value=self.settings.get("toner_sandbox_enabled", False))
        tk.Checkbutton(sandbox_frame, text="Allow sandbox mode",
                       variable=sandbox_var, bg=bg, fg=fg,
                       font=("Helvetica", 9)).pack(anchor="w")
        tk.Label(sandbox_frame,
                 text="Sandbox presets can capture any pitched sound\n"
                      "without requiring mic type or instrument fields.\n"
                      "Good for experiments, non-sax instruments, or\n"
                      "unconventional setups like contact mics.",
                 bg=bg, fg="#888888", font=("Helvetica", 8),
                 justify="left").pack(anchor="w", pady=(2, 0))

        # ==================================================================
        # OK / CANCEL
        # ==================================================================
        btn_frame = tk.Frame(dlg, bg=bg)
        btn_frame.pack(fill="x", padx=10, pady=(5, 10))

        def save():
            # Input device
            if devices and sys.platform != 'linux':
                sel = listbox.curselection()
                if sel:
                    self.settings["audio_input_device"] = dev_indices[sel[0]]

            # Recording
            rec = record_var.get()
            self.settings["toner_record_wav"] = rec
            self.settings["toner_wav_reanalyze"] = rec  # always reanalyze when recording
            self.settings["toner_wav_auto_delete"] = auto_delete_var.get()
            folder = folder_label.cget("text")
            if folder and folder != current_dir:
                self.settings["toner_recording_dir"] = folder

            # Pitch
            try:
                new_pitch = ref_pitch_var.get()
                if 420 <= new_pitch <= 460:
                    self._toner_pitch_var.set(new_pitch)
                    self._toner_on_pitch_changed()
            except (tk.TclError, ValueError):
                pass
            self._toner_concert_pitch.set(display_pitch_var.get() == "concert")
            self._toner_update_pitch_mode_label()

            # Preset fields
            for key, var in field_vars.items():
                vis[key] = var.get()
            self.settings["visible_preset_fields"] = vis

            # Analysis descriptors
            ts = self.settings.setdefault("toner_settings", {})
            ts["analysis_descriptors"] = {k: v.get() for k, v in desc_vars.items()}

            # Sandbox
            self.settings["toner_sandbox_enabled"] = sandbox_var.get()

            save_settings(self.settings)

            # Restart audio engine if device changed
            if devices and sys.platform != 'linux':
                sel = listbox.curselection()
                if sel:
                    new_dev = dev_indices[sel[0]]
                    if new_dev != self.settings.get("_prev_audio_device"):
                        if hasattr(self, '_toner_engine') and self._toner_engine and self._toner_engine.is_running:
                            self._toner_stop()
                            self._toner_start()

            dlg.destroy()

        tk.Button(btn_frame, text="OK", width=10, command=save).pack(side="right", padx=(5, 0))
        tk.Button(btn_frame, text="Cancel", width=10, command=dlg.destroy).pack(side="right")

    def _toner_transpose_note(self, concert_note):
        """Transpose a concert pitch note name to written pitch for display.

        Returns the original note if concert pitch display is on.
        Uses the engine's pure transpose_note() function.
        """
        if not concert_note or self._toner_concert_pitch.get():
            return concert_note
        return transpose_note(concert_note, self._toner_sax_var.get())

    def _toner_display_note_for_preset(self, concert_note, preset=None):
        """Transpose a concert note for display using a preset's horn type.

        Used in reports and comparisons where the active sax selector might
        not match the preset being viewed.
        """
        if not concert_note or self._toner_concert_pitch.get():
            return concert_note
        if preset:
            sax_type = preset.get('horn_type', self._toner_sax_var.get())
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

    def _toner_open_preset_dialog(self):
        """Open the preset management dialog — central hub for all preset operations."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Tone Presets")
        dlg.geometry("550x580")
        dlg.minsize(400, 450)
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Tone Presets", bg=bg, fg=fg,
                 font=("Helvetica", 14, "bold")).pack(pady=(0, 10))

        # Profile list
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(fill="both", expand=True, pady=(0, 5))

        self._preset_listbox = tk.Listbox(list_frame, width=55, height=12,
                                         font=("Helvetica", 10))
        self._preset_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame, command=self._preset_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._preset_listbox.config(yscrollcommand=scrollbar.set)

        self._toner_refresh_preset_list()

        # Info display (fixed height so long notes don't push buttons off screen)
        info_frame = tk.Frame(frame, bg=bg)
        info_frame.pack(fill="both", expand=True, pady=(0, 10))
        self._preset_info_text = tk.Text(info_frame, height=8, wrap="word",
                                          font=("Helvetica", 9), bg=bg, fg=fg,
                                          relief="flat", padx=4, pady=4,
                                          state="disabled")
        info_scroll = tk.Scrollbar(info_frame, command=self._preset_info_text.yview)
        self._preset_info_text.configure(yscrollcommand=info_scroll.set)
        info_scroll.pack(side="right", fill="y")
        self._preset_info_text.pack(side="left", fill="both", expand=True)

        self._preset_listbox.bind("<<ListboxSelect>>",
                                lambda e: self._toner_on_preset_selected())

        # --- Action buttons: two rows ---
        # Row 1: preset operations (need selection)
        row1 = tk.Frame(frame, bg=bg)
        row1.pack(fill="x", pady=(0, 5))

        btn_load = tk.Button(row1, text="Load for Capture", state="disabled",
                  command=lambda: self._toner_load_from_dialog(dlg))
        btn_load.pack(side="left", padx=(0, 5))
        btn_analyze = tk.Button(row1, text="Analyze...", state="disabled",
                  command=lambda: [dlg.destroy(),
                                   self._toner_open_analyze_dialog()])
        btn_analyze.pack(side="left", padx=(0, 5))
        btn_notes = tk.Button(row1, text="Edit Notes...", state="disabled",
                  command=self._toner_edit_preset_notes)
        btn_notes.pack(side="left", padx=(0, 5))

        # Row 2: create, import, delete, close
        row2 = tk.Frame(frame, bg=bg)
        row2.pack(fill="x")

        tk.Button(row2, text="New Preset...",
                  command=lambda: self._toner_new_preset(dlg)).pack(
                      side="left", padx=(0, 5))
        btn_mutate = tk.Button(row2, text="Mutate...", state="disabled",
                  command=self._toner_mutate_preset)
        btn_mutate.pack(side="left", padx=(0, 5))
        tk.Button(row2, text="Import Audio File...",
                  command=lambda: [dlg.destroy(),
                                   self._toner_import_audio_file()]).pack(
                      side="left", padx=(0, 5))
        btn_delete = tk.Button(row2, text="Delete", state="disabled",
                  command=self._toner_delete_preset)
        btn_delete.pack(side="left", padx=(0, 5))
        tk.Button(row2, text="Close",
                  command=dlg.destroy).pack(side="right")

        self._preset_selection_btns = [btn_load, btn_analyze, btn_notes,
                                        btn_mutate, btn_delete]

    def _toner_refresh_preset_list(self):
        """Refresh the preset listbox (nested library structure)."""
        if not hasattr(self, '_preset_listbox'):
            return
        self._preset_listbox.delete(0, tk.END)
        self._preset_list_keys = []  # Track (lib, name) for selection lookup
        for lib_name, lib_presets in self._toner_presets.items():
            if not isinstance(lib_presets, dict) or not lib_presets:
                continue
            self._preset_listbox.insert(tk.END, f"\u2500\u2500 {lib_name} \u2500\u2500")
            self._preset_list_keys.append(None)  # Header, not selectable
            for preset_name, preset in lib_presets.items():
                sessions = preset.get('sessions', [])
                total_caps = sum(len(s.get('captures', []))
                               for s in sessions)
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                status = f"{len(notes)} notes" if total_caps > 0 else "empty"
                if len(notes) >= MIN_PRESET_NOTES:
                    status += " \u2713"
                sb_tag = "[sandbox] " if preset.get('sandbox') else ""
                horn = preset.get('horn_type', '?') or '?'
                self._preset_listbox.insert(tk.END,
                    f"  {sb_tag}{preset_name}  ({horn}, {status})")
                self._preset_list_keys.append((lib_name, preset_name))

    def _toner_on_preset_selected(self):
        """Update info label when a preset is selected."""
        sel = self._preset_listbox.curselection()
        if not sel:
            return
        key = self._preset_list_keys[sel[0]]
        if key is None:
            self._preset_info_text.configure(state="normal")
            self._preset_info_text.delete("1.0", tk.END)
            self._preset_info_text.configure(state="disabled")
            for btn in self._preset_selection_btns:
                btn.configure(state="disabled")
            return  # Library header
        lib_name, preset_name = key
        preset = self._toner_presets[lib_name][preset_name]
        self._preset_info_text.configure(state="normal")
        self._preset_info_text.delete("1.0", tk.END)
        self._preset_info_text.insert("1.0", _format_preset_info(preset))
        self._preset_info_text.configure(state="disabled")
        for btn in self._preset_selection_btns:
            btn.configure(state="normal")

    def _toner_edit_preset_notes(self):
        """Edit the notes field of the selected preset."""
        sel = self._preset_listbox.curselection()
        if not sel:
            return
        key = self._preset_list_keys[sel[0]]
        if key is None:
            return
        lib_name, preset_name = key
        preset = self._toner_presets[lib_name][preset_name]
        self._toner_notes_dialog(preset_name, preset)

    def _toner_notes_dialog(self, preset_name, preset, prompt_text=None):
        """Open a multi-line notes editor for a preset."""
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Notes \u2014 {preset_name}")
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

        tk.Label(frame, text=f"Notes for \"{preset_name}\":", bg=bg,
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

        current = preset.get('notes', '')
        notes_text.insert("1.0", current)
        notes_text.focus_set()

        def save():
            new_notes = notes_text.get("1.0", tk.END).strip()
            preset['notes'] = new_notes
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            # Refresh info display if preset dialog is open
            if hasattr(self, '_preset_info_text'):
                try:
                    self._toner_on_preset_selected()
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

    def _toner_build_preset_fields(self, frame, bg, fg, defaults=None,
                                    default_lib=None):
        """Build preset form fields. Returns (fields_dict, lib_var, notes_text).

        Required fields are always shown. Optional fields respect
        the visible_preset_fields setting.

        Args:
            defaults: dict of preset data to pre-fill fields from (for mutate)
            default_lib: library name to pre-select
        """
        fields = {}
        vis = self.settings.get("visible_preset_fields", {})
        d = defaults or {}

        MIC_TYPES = ["Condenser", "Ribbon", "Dynamic", "Other"]

        def add_field(label, key, default="", widget_type="entry",
                      optional_key=None, data_key=None):
            """Add a labeled field row. If optional_key is set, only show
            when that key is enabled in visible_preset_fields."""
            if optional_key and not vis.get(optional_key, False):
                return
            val = d.get(data_key or key, default)
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, bg=bg, fg=fg, width=14,
                     anchor="e", font=("Helvetica", 10)).pack(
                side="left", padx=(0, 8))
            if widget_type == "combo":
                var = tk.StringVar(value=val)
                ttk.Combobox(row, textvariable=var, values=SAX_TYPES,
                             state="readonly", width=20).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var
            elif widget_type == "mic_combo":
                var = tk.StringVar(value=val.capitalize() if val else "")
                ttk.Combobox(row, textvariable=var, values=MIC_TYPES,
                             state="readonly", width=20).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var
            else:
                var = tk.StringVar(value=val)
                tk.Entry(row, textvariable=var, width=25).pack(
                    side="left", fill="x", expand=True)
                fields[key] = var

        # Library selector
        lib_row = tk.Frame(frame, bg=bg)
        lib_row.pack(fill="x", pady=2)
        tk.Label(lib_row, text="Library:", bg=bg, fg=fg, width=14,
                 anchor="e", font=("Helvetica", 10)).pack(
            side="left", padx=(0, 8))
        existing_libs = [k for k in self._toner_presets.keys()
                        if isinstance(self._toner_presets[k], dict)]
        if not existing_libs:
            existing_libs = [DEFAULT_LIBRARY]
        lib_val = default_lib if default_lib and default_lib in existing_libs else existing_libs[0]
        lib_var = tk.StringVar(value=lib_val)
        ttk.Combobox(lib_row, textvariable=lib_var,
                      values=existing_libs, width=20).pack(
            side="left", fill="x", expand=True)

        # Required fields
        add_field("Preset Name:", "name")
        add_field("Horn Type:", "horn_type", "Alto", widget_type="combo")
        add_field("Make:", "horn_make")
        add_field("Model:", "horn_model")
        add_field("Player:", "player")
        add_field("Mouthpiece:", "mouthpiece")
        add_field("Mic Type:", "mic_type", "", widget_type="mic_combo")
        add_field("Mic Model:", "mic_model", "")

        # Optional fields (shown based on visible_preset_fields settings)
        add_field("Serial #:", "serial", optional_key="serial")
        add_field("Reed:", "reed", optional_key="reed")
        add_field("Ligature:", "ligature", optional_key="ligature")
        add_field("Mic Position:", "mic_position", optional_key="mic_position")
        add_field("Room:", "room", optional_key="room")
        add_field("Preamp:", "preamp", optional_key="preamp")

        # Sandbox checkbox (only when enabled in settings)
        sandbox_var = None
        if self.settings.get("toner_sandbox_enabled"):
            sb_row = tk.Frame(frame, bg=bg)
            sb_row.pack(fill="x", pady=(6, 2))
            sandbox_var = tk.BooleanVar(value=d.get('sandbox', False))
            sb_cb = tk.Checkbutton(sb_row, text="Sandbox",
                                   variable=sandbox_var, bg=bg, fg=fg,
                                   font=("Helvetica", 10, "bold"))
            sb_cb.pack(side="left", padx=(0, 8))
            tk.Label(sb_row, text="(any sound, no required fields)",
                     bg=bg, fg="#888888", font=("Helvetica", 8)).pack(
                side="left")
            # Disable checkbox if editing an existing preset (sandbox is immutable)
            if d.get('sandbox') is not None and d.get('sessions'):
                sb_cb.configure(state="disabled")

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
            if d.get('notes'):
                notes_text.insert("1.0", d['notes'])

        return fields, lib_var, notes_text, sandbox_var

    def _toner_validate_required_fields(self, fields, dlg, sandbox=False):
        """Check required fields are filled. Returns True if valid.

        In sandbox mode, only the preset name is required.
        """
        if sandbox:
            required = [("name", "Preset Name")]
        else:
            required = [
                ("name", "Preset Name"),
                ("horn_make", "Make"),
                ("horn_model", "Model"),
                ("player", "Player"),
                ("mouthpiece", "Mouthpiece"),
                ("mic_type", "Mic Type"),
                ("mic_model", "Mic Model"),
            ]
        for key, label in required:
            if key in fields and not fields[key].get().strip():
                messagebox.showwarning("Required Field",
                    f"Please enter {label}.", parent=dlg)
                return False
        return True

    def _toner_collect_preset_data(self, fields, notes_text):
        """Collect preset data dict from form fields."""
        data = {
            'horn_type': fields.get("horn_type", tk.StringVar()).get(),
            'horn_make': fields.get("horn_make", tk.StringVar()).get().strip(),
            'horn_model': fields.get("horn_model", tk.StringVar()).get().strip(),
            'serial': fields.get("serial", tk.StringVar()).get().strip() if "serial" in fields else "",
            'player': fields.get("player", tk.StringVar()).get().strip(),
            'mouthpiece': fields.get("mouthpiece", tk.StringVar()).get().strip(),
            'reed': fields.get("reed", tk.StringVar()).get().strip() if "reed" in fields else "",
            'ligature': fields.get("ligature", tk.StringVar()).get().strip() if "ligature" in fields else "",
            'mic_position': fields.get("mic_position", tk.StringVar()).get().strip() if "mic_position" in fields else "",
            'room': fields.get("room", tk.StringVar()).get().strip() if "room" in fields else "",
            'preamp': fields.get("preamp", tk.StringVar()).get().strip() if "preamp" in fields else "",
            'mic_type': fields.get("mic_type", tk.StringVar()).get().lower(),
            'mic_model': fields.get("mic_model", tk.StringVar()).get().strip(),
            'notes': notes_text.get("1.0", tk.END).strip() if notes_text else "",
            'created': time.strftime("%Y-%m-%d"),
            'sessions': [],
        }
        return data

    def _toner_new_preset(self, parent_dlg):
        """Create a new horn preset via guided dialog."""
        dlg = tk.Toplevel(parent_dlg)
        dlg.title("New Tone Preset")
        dlg.resizable(False, False)
        dlg.transient(parent_dlg)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Preset", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 10))

        fields, lib_var, notes_text, sandbox_var = self._toner_build_preset_fields(
            frame, bg, fg)

        def save():
            is_sandbox = sandbox_var and sandbox_var.get()
            if not self._toner_validate_required_fields(fields, dlg,
                                                         sandbox=is_sandbox):
                return
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if lib not in self._toner_presets:
                self._toner_presets[lib] = {}
            if name in self._toner_presets[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            data = self._toner_collect_preset_data(fields, notes_text)
            if is_sandbox:
                data['sandbox'] = True
            self._toner_presets[lib][name] = data
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            self._toner_active_library = lib
            self._toner_active_preset = name
            self._toner_active_session = None
            self._toner_update_preset_label()
            self._toner_refresh_preset_list()
            dlg.destroy()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Create", command=save).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="left")

    def _toner_mutate_preset(self):
        """Duplicate selected preset with editable fields, save as new."""
        sel = self._preset_listbox.curselection()
        if not sel:
            return
        key = self._preset_list_keys[sel[0]]
        if key is None:
            return
        lib_name, preset_name = key
        source = self._toner_presets[lib_name][preset_name]

        parent = self._preset_listbox.winfo_toplevel()
        dlg = tk.Toplevel(parent)
        dlg.title("Mutate Preset")
        dlg.resizable(False, False)
        dlg.transient(parent)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Mutate Preset", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 2))
        tk.Label(frame, text="Change what you need, save as a new preset.",
                 bg=bg, fg="#666666", font=("Helvetica", 9)).pack(pady=(0, 10))

        # Pre-fill from source, clear the name so they must enter a new one
        defaults = dict(source)
        defaults['name'] = ""

        fields, lib_var, notes_text, sandbox_var = self._toner_build_preset_fields(
            frame, bg, fg, defaults=defaults, default_lib=lib_name)

        def save():
            is_sandbox = sandbox_var and sandbox_var.get()
            if not self._toner_validate_required_fields(fields, dlg,
                                                         sandbox=is_sandbox):
                return
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if lib not in self._toner_presets:
                self._toner_presets[lib] = {}
            if name in self._toner_presets[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            data = self._toner_collect_preset_data(fields, notes_text)
            if is_sandbox:
                data['sandbox'] = True
            self._toner_presets[lib][name] = data
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            self._toner_active_library = lib
            self._toner_active_preset = name
            self._toner_active_session = None
            self._toner_update_preset_label()
            self._toner_refresh_preset_list()
            dlg.destroy()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Save as New", command=save).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel", command=dlg.destroy).pack(side="left")

    def _toner_delete_preset(self):
        """Delete the selected preset."""
        sel = self._preset_listbox.curselection()
        if not sel:
            return
        key = self._preset_list_keys[sel[0]]
        if key is None:
            return  # Library header
        lib_name, preset_name = key
        if messagebox.askyesno("Delete Preset",
                f"Delete preset '{preset_name}' from '{lib_name}'?"):
            del self._toner_presets[lib_name][preset_name]
            # Remove empty libraries
            if not self._toner_presets[lib_name]:
                del self._toner_presets[lib_name]
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            if (self._toner_active_library == lib_name and
                    self._toner_active_preset == preset_name):
                self._toner_active_library = None
                self._toner_active_preset = None
                self._toner_active_session = None
            self._toner_refresh_preset_list()
            self._preset_info_text.configure(state="normal")
            self._preset_info_text.delete("1.0", tk.END)
            self._preset_info_text.configure(state="disabled")

    # ------------------------------------------------------------------
    # CAPTURE SYSTEM (auto-detect stable tones)
    # ------------------------------------------------------------------

    def _toner_toggle_capture(self):
        """Toggle capture mode on/off."""
        if self._toner_capture_state is not None:
            # Stop capturing
            self._toner_stop_capture()
            return

        # Need a preset loaded
        if not self._toner_active_preset:
            messagebox.showinfo("No Preset Loaded",
                "Load a preset first to capture data.\n\n"
                "Use Load... on the control strip, or\n"
                "File > Presets to create and manage presets.")
            self._toner_open_preset_dialog()
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
            is_sb = (self._toner_active_session or {}).get('sandbox')
            label_text = "[sandbox] Play anything..." if is_sb else "Play anything..."

        self._toner_stable_note = ""
        self._toner_stable_count = 0
        self._toner_capture_btn.configure(text="Stop")

        self._toner_capture_frame.pack(fill="x", padx=5,
            before=self._toner_main_frame.winfo_children()[-1])
        self._toner_capture_label.configure(text=label_text)
        self._toner_capture_progress.configure(text="")


    def _toner_new_preset_flow(self):
        """Create a new preset (complete setup identity), then start capturing."""
        dlg = tk.Toplevel(self.root)
        dlg.title("New Tone Preset")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Preset", bg=bg, fg=fg,
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 5))
        tk.Label(frame, text="A preset saves your setup details: horn + player + "
                 "mouthpiece.\nIt pre-fills session metadata for quick capture start.",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(pady=(0, 10))

        fields, lib_var, notes_text, sandbox_var = self._toner_build_preset_fields(
            frame, bg, fg)

        def save_and_start():
            is_sandbox = sandbox_var and sandbox_var.get()
            if not self._toner_validate_required_fields(fields, dlg,
                                                         sandbox=is_sandbox):
                return
            name = fields["name"].get().strip()
            lib = lib_var.get().strip() or DEFAULT_LIBRARY
            if lib not in self._toner_presets:
                self._toner_presets[lib] = {}
            if name in self._toner_presets[lib]:
                messagebox.showwarning("Duplicate Name",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            data = self._toner_collect_preset_data(fields, notes_text)
            if is_sandbox:
                data['sandbox'] = True
            self._toner_presets[lib][name] = data
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            self._toner_active_library = lib
            self._toner_active_preset = name
            dlg.destroy()
            self._toner_start_new_session_and_listen()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(10, 0))
        tk.Button(btn_row, text="Create && Start Capturing",
                  command=save_and_start).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel",
                  command=dlg.destroy).pack(side="left")

    def _toner_load_from_dialog(self, dlg):
        """Load the selected preset from the preset dialog."""
        sel = self._preset_listbox.curselection()
        if not sel:
            messagebox.showinfo("Select Preset", "Select a preset first.")
            return
        key = self._preset_list_keys[sel[0]]
        if key is None:
            return
        lib_name, preset_name = key
        self._toner_active_library = lib_name
        self._toner_active_preset = preset_name
        self._toner_active_session = None
        self._toner_update_preset_label()

        # Sync sax type and lock selector
        preset = self._toner_presets[lib_name][preset_name]
        sax_type = preset.get('horn_type', '')
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
        """Toggle between Load and Unload based on current preset state."""
        if self._toner_active_preset:
            self._toner_unload_preset()
        else:
            self._toner_load_preset_quick()

    def _toner_update_load_unload_btn(self):
        """Update the Load/Unload button text based on preset state."""
        if hasattr(self, '_toner_load_unload_btn'):
            if self._toner_active_preset:
                self._toner_load_unload_btn.configure(text="Unload")
            else:
                self._toner_load_unload_btn.configure(text="Load...")

    def _toner_load_preset_quick(self):
        """Quick preset loader — shows list, loads selected."""
        all_presets = []
        for lib_name, lib_presets in self._toner_presets.items():
            if not isinstance(lib_presets, dict):
                continue
            for preset_name, preset_data in lib_presets.items():
                all_presets.append((lib_name, preset_name, preset_data))

        if not all_presets:
            messagebox.showinfo("No Presets",
                "No presets yet. Create one in File > Presets.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Load Preset")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        frame = tk.Frame(dlg, bg=bg, padx=15, pady=10)
        frame.pack(fill="both", expand=True)

        listbox = tk.Listbox(frame, width=40, height=min(10, len(all_presets)),
                              font=("Helvetica", 10))
        listbox.pack(fill="both", expand=True, pady=(0, 8))

        for lib_name, preset_name, _ in all_presets:
            listbox.insert(tk.END, f"[{lib_name}] {preset_name}")

        def load():
            sel = listbox.curselection()
            if not sel:
                return
            lib_name, preset_name, _ = all_presets[sel[0]]
            self._toner_active_library = lib_name
            self._toner_active_preset = preset_name
            self._toner_active_session = None
            self._toner_update_preset_label()

            # Sync sax type and lock selector
            preset = self._toner_presets[lib_name][preset_name]
            sax_type = preset.get('horn_type', '')
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

    def _toner_unload_preset(self):
        """Unload the active preset."""
        self._toner_active_library = None
        self._toner_active_preset = None
        self._toner_active_session = None
        self._toner_update_preset_label()
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

    def _toner_update_preset_label(self):
        """Update the active preset indicator, mic label, and Load/Unload button."""
        if hasattr(self, '_toner_preset_label'):
            name = self._toner_active_preset or ""
            if name:
                display = name[:PRESET_NAME_MAX_DISPLAY] + "..." if len(name) > PRESET_NAME_MAX_DISPLAY else name
                self._toner_preset_label.configure(text=display, fg="#AAAAAA")
            else:
                self._toner_preset_label.configure(text="no preset", fg="#666666")
        if hasattr(self, '_toner_mic_label'):
            preset = {}
            if self._toner_active_library and self._toner_active_preset:
                lib = self._toner_presets.get(self._toner_active_library, {})
                preset = lib.get(self._toner_active_preset, {})
            mt = preset.get('mic_type', '')
            mm = preset.get('mic_model', '')
            if mt:
                mic_text = mm if mm else mt.capitalize()
                self._toner_mic_label.configure(text=mic_text, fg="#AAAAAA")
            elif preset.get('sandbox'):
                self._toner_mic_label.configure(text="sandbox", fg="#CC8800")
            elif self._toner_active_preset:
                self._toner_mic_label.configure(text="mic not set", fg="#FF8800")
            else:
                self._toner_mic_label.configure(text="no preset", fg="#666666")
        # Sync sandbox checkbox with loaded preset
        if hasattr(self, '_toner_sandbox_var'):
            is_sb = bool(preset.get('sandbox')) if preset else False
            self._toner_sandbox_var.set(is_sb)
            if is_sb:
                self._toner_concert_pitch.set(True)
                self._toner_update_pitch_mode_label()
        self._toner_update_load_unload_btn()

    def _toner_start_new_session_and_listen(self):
        """Create a new session for the active preset and begin listening."""
        # Prompt for mic type if not set on preset (skip for sandbox)
        preset_check = {}
        if self._toner_active_library and self._toner_active_preset:
            lib = self._toner_presets.get(self._toner_active_library, {})
            preset_check = lib.get(self._toner_active_preset, {})
        if not preset_check.get('sandbox') and not preset_check.get('mic_type'):
            messagebox.showinfo("Mic Type Required",
                "This preset has no mic type set.\n\n"
                "Use Mutate in File \u2192 Presets to add mic info,\n"
                "or create a new preset with mic type specified.",
                parent=self.root)
            return

        self._toner_update_preset_label()
        # Sync sax type from preset to selector and engine
        if self._toner_active_library and self._toner_active_preset:
            lib = self._toner_presets.get(self._toner_active_library, {})
            preset = lib.get(self._toner_active_preset, {})
            sax_type = preset.get('horn_type', '')
            if sax_type:
                if hasattr(self, '_toner_sax_var'):
                    self._toner_sax_var.set(sax_type)
                if self._toner_engine:
                    self._toner_engine.set_sax_type(sax_type)

        # Copy preset metadata into session so each session owns its data
        preset = {}
        if self._toner_active_library and self._toner_active_preset:
            lib = self._toner_presets.get(self._toner_active_library, {})
            preset = lib.get(self._toner_active_preset, {})

        is_sandbox = preset.get('sandbox', False)
        self._toner_active_session = {
            'date': time.strftime("%Y-%m-%d %H:%M:%S"),
            'captures': [],
            'mic_type': preset.get('mic_type', ''),
            'mic_model': preset.get('mic_model', ''),
            'horn_type': preset.get('horn_type', ''),
            'horn_make': preset.get('horn_make', ''),
            'horn_model': preset.get('horn_model', ''),
            'serial': preset.get('serial', ''),
            'player': preset.get('player', ''),
            'mouthpiece': preset.get('mouthpiece', ''),
            'reed': preset.get('reed', ''),
            'mic_position': preset.get('mic_position', ''),
        }
        if is_sandbox:
            self._toner_active_session['sandbox'] = True
        self._toner_rolloff_warned = False
        self._toner_captured_notes = set()
        self._toner_save_pending = False

        # Start WAV recording if enabled
        if self.settings.get('toner_record_wav') and self._toner_engine:
            # First time: require user to pick a folder
            if not self.settings.get('toner_recording_dir'):
                from tkinter import filedialog
                messagebox.showinfo("Choose Recording Folder",
                    "WAV recording is enabled but no folder has been chosen.\n\n"
                    "Please select a folder where recordings will be saved.",
                    parent=self.root)
                folder = filedialog.askdirectory(
                    title="Choose Recording Folder",
                    initialdir=os.path.expanduser("~"),
                    parent=self.root)
                if folder:
                    self.settings['toner_recording_dir'] = folder
                    save_settings(self.settings)
                else:
                    # User cancelled — disable WAV recording for this session
                    self.settings['toner_record_wav'] = False
                    save_settings(self.settings)
            if self.settings.get('toner_record_wav'):
                self._toner_engine.start_recording()

        self._toner_begin_listening()

    def _toner_stop_capture(self):
        """Stop capture mode. Saves pending data, shows coverage summary."""
        # Flush any deferred save
        self._toner_flush_save()
        # Save any accumulated free-mode frames before stopping
        if self._toner_capture_mode == 'free' and self._toner_free_accumulator:
            self._toner_free_save_micro_capture()

        # Save WAV recording if active
        wav_filepath = None
        if self._toner_engine:
            chunks = self._toner_engine.stop_recording()
            if chunks and self._toner_active_session:
                wav_filepath = self._toner_save_wav_recording(chunks)

        # Reanalyze from WAV for max accuracy (replaces live captures)
        if (wav_filepath and self._toner_active_session
                and self.settings.get('toner_wav_reanalyze')):
            self._toner_reanalyze_from_wav(wav_filepath)

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

    def _toner_get_recording_dir(self):
        """Return the directory for WAV recordings, creating it if needed."""
        rec_dir = self.settings.get('toner_recording_dir', '')
        if not rec_dir:
            # Default: Music/StohrerSaxShopCompanion or Documents fallback
            music = os.path.join(os.path.expanduser('~'), 'Music')
            if not os.path.isdir(music):
                music = os.path.join(os.path.expanduser('~'), 'Documents')
            rec_dir = os.path.join(music, 'StohrerSaxShopCompanion')
        os.makedirs(rec_dir, exist_ok=True)
        return rec_dir

    def _toner_save_wav_recording(self, chunks):
        """Save recorded audio chunks to a WAV file. Returns filepath or None."""
        try:
            rec_dir = self._toner_get_recording_dir()
            # Build filename from preset name and current time (unique per save)
            preset_name = ''
            if self._toner_active_preset:
                preset_name = self._toner_active_preset
            # Sanitize for filesystem
            safe_name = "".join(c if c.isalnum() or c in ' -_' else '_'
                                for c in preset_name).strip() or 'session'
            timestamp = time.strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"{safe_name}_{timestamp}.wav"
            filepath = os.path.join(rec_dir, filename)

            from toner_engine import TonerEngine
            TonerEngine.save_recording(chunks, filepath)

            # Verify file was created
            if not os.path.isfile(filepath) or os.path.getsize(filepath) == 0:
                messagebox.showwarning("Recording",
                    "WAV recording failed — file was not created.")
                return None

            # Store reference in session
            self._toner_active_session['recording_file'] = filename
            self._toner_save_active_session()

            if not self.settings.get('toner_wav_reanalyze'):
                messagebox.showinfo("Recording Saved",
                    f"WAV saved to:\n{filepath}")
            return filepath
        except Exception as e:
            messagebox.showwarning("Recording",
                f"Could not save WAV recording:\n{e}")
            return None

    def _toner_reanalyze_from_wav(self, wav_filepath):
        """Replace live captures with offline WAV analysis for better accuracy."""
        if not self._toner_active_session or not self._toner_engine:
            return
        live_count = len(self._toner_active_session.get('captures', []))

        # Show progress in the capture frame
        self._toner_capture_label.configure(text="Performing WAV analysis...")
        self._toner_capture_progress.configure(text="")
        self.root.update_idletasks()

        def on_progress(current, total):
            pct = current * 100 // total
            self._toner_capture_progress.configure(text=f"{pct}%")
            self.root.update_idletasks()

        try:
            captures = analyze_audio_file(wav_filepath, self._toner_engine,
                                          progress_cb=on_progress)
        except Exception as e:
            messagebox.showwarning("WAV Analysis",
                f"WAV analysis failed, keeping live captures:\n{e}")
            return

        if not captures:
            messagebox.showinfo("WAV Analysis",
                "WAV analysis found no stable segments.\n"
                "Keeping live captures.")
            return

        # Stamp captures with metadata
        for cap in captures:
            cap['timestamp'] = time.strftime("%Y-%m-%d %H:%M:%S")
            cap['method'] = 'file'

        self._toner_active_session['captures'] = captures
        self._toner_save_active_session()

        wav_notes = len(set(c['note'] for c in captures))
        msg = (f"WAV analysis: {len(captures)} captures "
               f"across {wav_notes} notes\n"
               f"(replaced {live_count} live captures)")

        # Auto-delete WAV if enabled
        if self.settings.get('toner_wav_auto_delete'):
            try:
                os.remove(wav_filepath)
                self._toner_active_session.pop('recording_file', None)
                self._toner_save_active_session()
                msg += "\nWAV file deleted."
            except OSError:
                msg += "\nCould not delete WAV file."

        messagebox.showinfo("WAV Analysis Complete", msg)

    def _toner_show_coverage_summary(self):
        """Show a coverage summary with note distribution and resume option."""
        session = self._toner_active_session
        if not session:
            return

        captures = session.get('captures', [])
        if not captures:
            return

        # Get active preset for note transposition
        _cov_preset = None
        if self._toner_active_library and self._toner_active_preset:
            lib = self._toner_presets.get(self._toner_active_library, {})
            _cov_preset = lib.get(self._toner_active_preset)

        dlg = tk.Toplevel(self.root)
        dlg.title("Capture Summary")
        dlg.resizable(False, False)
        dlg.transient(self.root)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        preset_name = self._toner_active_preset or "?"
        tk.Label(frame, text=f"Session: {preset_name}", bg=bg, fg=fg,
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
                    disp_note = self._toner_display_note_for_preset(
                        note, _cov_preset)
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
        if unique < MIN_PRESET_NOTES:
            assessment.append(f"Need {MIN_PRESET_NOTES - unique} more unique "
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
            # Clear session — it's saved, don't carry it to the next preset
            session_lib = self._toner_active_library
            session_prof = self._toner_active_preset
            self._toner_active_session = None
            # Prompt for notes after session
            lib = session_lib
            preset_name = session_prof
            if lib and preset_name and lib in self._toner_presets:
                preset = self._toner_presets[lib].get(preset_name)
                if preset:
                    self._toner_notes_dialog(preset_name, preset,
                        prompt_text="How did this horn sound to you? "
                        "Add your impressions \u2014 bright, dark, rich, "
                        "stuffy, free-blowing, anything you noticed.")

        def discard():
            if messagebox.askyesno("Discard Session",
                    "Discard all captures from this session?\n\n"
                    "This cannot be undone.", parent=dlg):
                # Remove this session's captures from the preset
                lib = self._toner_active_library
                preset_name = self._toner_active_preset
                if lib and preset_name and lib in self._toner_presets:
                    preset = self._toner_presets[lib].get(preset_name)
                    if preset and self._toner_active_session:
                        session_date = self._toner_active_session.get('date', '')
                        sessions = preset.get('sessions', [])
                        preset['sessions'] = [s for s in sessions
                                               if s.get('date') != session_date]
                        save_tone_presets(self._toner_presets, TONER_DATA_FILE)
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

            # Count unique notes captured so far (cached set)
            notes_so_far = self._toner_captured_notes

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
                        self._toner_captured_notes.add(concert_note)
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
            self._toner_captured_notes.add(dominant_note)
            self._toner_schedule_save()
            self._toner_check_rolloff_warning()

    def _toner_check_rolloff_warning(self):
        """Check harmonic rolloff rate and warn if mic/room quality is poor."""
        if getattr(self, '_toner_rolloff_warned', False):
            return
        session = self._toner_active_session
        if not session:
            return
        # Check if preset has suppressed this warning
        if self._toner_active_library and self._toner_active_preset:
            lib = self._toner_presets.get(self._toner_active_library, {})
            preset = lib.get(self._toner_active_preset, {})
            if preset.get('suppress_rolloff_warning'):
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
        # Mic-class-aware threshold: ribbon and dynamic mics legitimately
        # read higher rolloff than condensers due to physics, not bad
        # recording quality.
        threshold = get_rolloff_threshold(session.get('mic_type', ''))
        if avg_rate > threshold:
            self._toner_rolloff_warned = True
            self._toner_show_rolloff_warning(avg_rate)

    def _toner_show_rolloff_warning(self, avg_rate):
        """Show rolloff warning with option to suppress for this preset."""
        warn_dlg = tk.Toplevel(self.root)
        warn_dlg.title("Recording Quality")
        warn_dlg.transient(self.root)
        warn_dlg.grab_set()
        warn_dlg.resizable(False, False)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(warn_dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="\u26a0 Recording Quality", bg=bg, fg=fg,
                 font=("Helvetica", 11, "bold")).pack(anchor="w", pady=(0, 8))

        tk.Label(frame,
                 text=f"Upper harmonics are dropping off steeply "
                      f"({avg_rate:.1f} dB/harmonic).\n\n"
                      f"This can be caused by:\n"
                      f"  \u2022  Microphone too far from the bell\n"
                      f"  \u2022  Built-in laptop mic\n"
                      f"  \u2022  Playing technique (subtone, soft dynamics)\n\n"
                      f"For best results, use an external condenser mic\n"
                      f"2\u20133 feet from the bell.",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(anchor="w", pady=(0, 10))

        suppress_var = tk.BooleanVar(value=False)
        tk.Checkbutton(frame,
                       text="Don't show this again for this preset",
                       variable=suppress_var, bg=bg, fg=fg,
                       font=("Helvetica", 9)).pack(anchor="w", pady=(0, 10))

        def _dismiss():
            if suppress_var.get():
                if self._toner_active_library and self._toner_active_preset:
                    lib = self._toner_presets.get(
                        self._toner_active_library, {})
                    preset = lib.get(self._toner_active_preset, {})
                    preset['suppress_rolloff_warning'] = True
                    save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            warn_dlg.destroy()

        tk.Button(frame, text="OK", command=_dismiss,
                  width=10).pack(anchor="e")

    def _toner_schedule_save(self):
        """Schedule a deferred save (batches rapid captures)."""
        if not self._toner_save_pending:
            self._toner_save_pending = True
            try:
                self.root.after(10000, self._toner_flush_save)
            except tk.TclError:
                pass

    def _toner_flush_save(self):
        """Execute a deferred save if one is pending."""
        if self._toner_save_pending:
            self._toner_save_pending = False
            self._toner_save_active_session()

    def _toner_save_active_session(self):
        """Save the active session to the active preset.

        Saves a deep copy of the session data so that subsequent
        captures on a different preset don't mutate this preset's
        stored data through shared references.
        """
        import copy
        lib = self._toner_active_library
        preset_name = self._toner_active_preset
        if not lib or not preset_name:
            return
        if lib not in self._toner_presets:
            return
        lib_presets = self._toner_presets[lib]
        if preset_name not in lib_presets:
            return

        preset = lib_presets[preset_name]
        sessions = preset.setdefault('sessions', [])
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

        save_tone_presets(self._toner_presets, TONER_DATA_FILE)

    # ------------------------------------------------------------------
    # COMPARISON
    # ------------------------------------------------------------------

    def _toner_open_analyze_dialog(self):
        """Open dialog to select profiles/sessions for analysis."""
        all_presets = []
        for lib_name, lib_presets in self._toner_presets.items():
            if not isinstance(lib_presets, dict):
                continue
            for preset_name, preset_data in lib_presets.items():
                sessions = preset_data.get('sessions', [])
                total_caps = sum(len(s.get('captures', [])) for s in sessions)
                if total_caps > 0:
                    all_presets.append((lib_name, preset_name, preset_data))

        if not all_presets:
            messagebox.showinfo("No Presets",
                "No presets with captures to analyze.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Analyze Tone Data")
        dlg.geometry("600x520")
        dlg.resizable(True, True)
        dlg.minsize(500, 400)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select one or more presets or sessions to analyze, compare, or overlay.",
                 bg=bg, fg=fg, font=("Helvetica", 10),
                 justify="left").pack(pady=(0, 5))

        # Filter controls — Row 1: horn identity
        filter_row1 = tk.Frame(frame, bg=bg)
        filter_row1.pack(fill="x", pady=(0, 2))

        # Collect unique values for filters
        def _unique(field):
            return sorted(set(p.get(field, '') for _, _, p in all_presets
                              if p.get(field)))
        all_types = _unique('horn_type')
        all_makes = _unique('horn_make')
        all_models = _unique('horn_model')
        all_players = _unique('player')
        all_mpcs = _unique('mouthpiece')

        # Mic types from sessions
        all_mic_types = set()
        for _, _, p in all_presets:
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
        # represents: either an entire preset or a single session.
        # {'type': 'preset', 'lib': str, 'name': str, 'preset': dict}
        # {'type': 'session', 'lib': str, 'name': str, 'preset': dict,
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
            for lib_name, preset_name, preset in all_presets:
                if ft != "All" and preset.get('horn_type', '') != ft:
                    continue
                if fmk != "All" and preset.get('horn_make', '') != fmk:
                    continue
                if fmd != "All" and preset.get('horn_model', '') != fmd:
                    continue
                if fp != "All" and preset.get('player', '') != fp:
                    continue
                if fm != "All" and preset.get('mouthpiece', '') != fm:
                    continue
                if fmt != "All":
                    prof_mic_types = set(
                        s.get('mic_type', '').capitalize()
                        for s in preset.get('sessions', [])
                        if s.get('mic_type'))
                    if fmt not in prof_mic_types:
                        continue
                if search:
                    haystack = " ".join([
                        preset_name,
                        preset.get('horn_make', ''),
                        preset.get('horn_model', ''),
                        preset.get('serial', ''),
                        preset.get('player', ''),
                        preset.get('mouthpiece', ''),
                        preset.get('reed', ''),
                        preset.get('ligature', ''),
                        preset.get('mic_position', ''),
                        preset.get('room', ''),
                        preset.get('preamp', ''),
                        preset.get('notes', ''),
                    ]).lower()
                    # Also search session-level mic info
                    for s in preset.get('sessions', []):
                        haystack += " " + s.get('mic_type', '')
                        haystack += " " + s.get('mic_model', '')
                    if search not in haystack:
                        continue
                sessions = [s for s in preset.get('sessions', [])
                            if s.get('captures')]
                notes = set()
                for s in sessions:
                    for c in s.get('captures', []):
                        notes.add(c.get('note', ''))
                if not notes:
                    continue
                status = f"{len(notes)} notes"
                if len(notes) >= MIN_PRESET_NOTES:
                    status += " \u2713"

                # Profile-level checkbox (all sessions)
                var = tk.BooleanVar(value=False)
                check_vars.append(var)
                check_items.append({'type': 'preset', 'lib': lib_name,
                                    'name': preset_name, 'preset': preset})
                session_vars = []  # track child session vars

                def _make_preset_toggle(pvar, svars):
                    def _toggle():
                        val = pvar.get()
                        for sv in svars:
                            sv.set(val)
                    return _toggle

                tk.Checkbutton(
                    list_inner,
                    text=f"[{lib_name}] {preset_name}  ({status})",
                    variable=var, bg=bg, fg=fg,
                    selectcolor=bg, activebackground=bg,
                    anchor="w", font=("Helvetica", 10),
                    command=_make_preset_toggle(var, session_vars),
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
                            'name': preset_name, 'preset': preset,
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
                    sax = item['preset'].get('horn_type', 'Tenor')
                    if item['type'] == 'preset':
                        fp_val = compute_fingerprint(
                            item['preset'].get('sessions', []), sax)
                        fp_val['_name'] = item['name']
                        fp_val['_preset'] = item['preset']
                    else:
                        fp_val = compute_session_fingerprint(
                            item['session'], sax)
                        if not fp_val:
                            continue
                        fp_val['_name'] = (f"{item['name']} \u2014 "
                                           f"{item['date']}")
                        fp_val['_preset'] = item['preset']
                    results.append(fp_val)
            return results

        def analyze_selected():
            """Open analysis window for selected profiles/sessions."""
            selected = get_selected()
            if not selected:
                messagebox.showinfo("Select",
                    "Select at least one preset or session.", parent=dlg)
                return
            # Cross-sax-type warning. Mixing alto + tenor + bari + soprano
            # in the same comparison is allowed but the readings get
            # confounded by physics: lower-pitched horns intrinsically read
            # warmer and brighter on the H4-H5 measure regardless of
            # mouthpiece, so cross-type comparisons mostly reflect
            # fundamental frequency rather than tonal character.
            sax_types_in_selection = sorted(set(
                fp.get('_preset', {}).get('horn_type', '') or '?'
                for fp in selected))
            if len(sax_types_in_selection) > 1:
                proceed = messagebox.askyesno(
                    "Mixed Sax Types",
                    "You're comparing presets across different sax types "
                    f"({', '.join(sax_types_in_selection)}).\n\n"
                    "Lower-pitched horns intrinsically read warmer and "
                    "brighter on most descriptors, regardless of "
                    "mouthpiece or player. Cross-type comparisons mostly "
                    "reflect that physics, not the tonal character of "
                    "the instruments.\n\n"
                    "The Character Map and descriptor table are most "
                    "reliable when comparing within a single sax type.\n\n"
                    "Continue anyway?",
                    parent=dlg, icon='warning', default='no')
                if not proceed:
                    return
            # Determine primary sax type for population stats
            sax_type = selected[0].get('_preset', {}).get(
                'horn_type', 'Tenor')
            pop_stats = compute_population_stats(all_presets, sax_type)
            dlg.destroy()
            self._toner_show_analysis(selected, population_stats=pop_stats)

        def clear_comparison():
            self._toner_clear_overlay()
            dlg.destroy()

        def average_selected():
            """Compute group average of selected profiles (preset-level only)."""
            preset_list = []
            for i, var in enumerate(check_vars):
                if var.get() and check_items[i]['type'] == 'preset':
                    item = check_items[i]
                    preset_list.append((item['name'], item['preset']))
            if len(preset_list) < 2:
                messagebox.showinfo("Select More",
                    "Select at least 2 presets to average.\n"
                    "(Individual sessions are not included in group averages.)",
                    parent=dlg)
                return
            # Cross-sax-type warning — averaging across sax types is even
            # less meaningful than comparing them, since the result has no
            # natural physical interpretation.
            sax_types_in_selection = sorted(set(
                p.get('horn_type', '') or '?' for _, p in preset_list))
            if len(sax_types_in_selection) > 1:
                proceed = messagebox.askyesno(
                    "Mixed Sax Types",
                    "You're averaging presets across different sax types "
                    f"({', '.join(sax_types_in_selection)}).\n\n"
                    "Lower-pitched horns intrinsically read warmer and "
                    "brighter than higher-pitched horns. Averaging across "
                    "sax types mixes that physics into a single number "
                    "with no clear physical interpretation.\n\n"
                    "Group averages are most meaningful within a single "
                    "sax type.\n\n"
                    "Continue anyway?",
                    parent=dlg, icon='warning', default='no')
                if not proceed:
                    return
            dlg.destroy()
            self._toner_show_group_report(preset_list)

        tk.Button(btn_frame, text="Analyze Selected",
                  command=analyze_selected).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Average Selected",
                  command=average_selected).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Clear Overlay",
                  command=clear_comparison).pack(side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel",
                  command=dlg.destroy).pack(side="right")

    def _toner_show_analysis(self, fingerprints, population_stats=None):
        """Show analysis window for one or more profiles/sessions.

        Single selection: preset view with chart, descriptors, per-note.
        Multiple selections: comparison with delta analysis.
        """
        is_single = len(fingerprints) == 1

        dlg = tk.Toplevel(self.root)
        if is_single:
            dlg.title(f"Analyze \u2014 {fingerprints[0].get('_name', '?')}")
        else:
            dlg.title("Analyze \u2014 Comparison")
        # Single-preset stays compact; comparison view needs room for the
        # extra Character Map chart between the harmonic chart and table.
        dlg.geometry("720x950" if len(fingerprints) >= 2 else "720x750")
        # Intentionally NOT transient — on Windows, transient toplevels lose
        # the maximize button. The Analyze window is non-modal and content-
        # heavy; users want to maximize it, alt-tab to it, and leave it open
        # while doing other work. It should behave as an independent window.
        # Allow the user to resize / maximize / fullscreen the window —
        # there's a lot of information here and bigger is often better.
        # The window content is taller than any viewport when there's a lot
        # of comparison data, so the whole thing scrolls vertically.
        # minsize stays modest because scrolling handles the overflow.
        dlg.resizable(True, True)
        dlg.minsize(640, 500)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"

        chart_colors = ["#2196F3", "#FF5722", "#4CAF50", "#FF9800",
                        "#9C27B0", "#00BCD4", "#E91E63", "#8BC34A"]

        # --- Scrollable main content ---
        # Outer holds the scrollbar + the scroll canvas. The "main" frame
        # used by the rest of this method lives INSIDE the scroll canvas
        # via create_window so the entire UI scrolls as one piece.
        outer = tk.Frame(dlg, bg=bg)
        outer.pack(fill="both", expand=True)

        scroll_vbar = tk.Scrollbar(outer, orient="vertical")
        scroll_vbar.pack(side="right", fill="y")

        scroll_canvas = tk.Canvas(outer, bg=bg, highlightthickness=0,
                                   yscrollcommand=scroll_vbar.set)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll_vbar.config(command=scroll_canvas.yview)

        main = tk.Frame(scroll_canvas, bg=bg)
        scroll_window = scroll_canvas.create_window(
            (0, 0), window=main, anchor="nw", tags="scroll_inner")

        # Update the scrollregion whenever the inner frame's natural size
        # changes (e.g. when widgets are added or text is rebuilt).
        def _update_scrollregion(event=None):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        main.bind("<Configure>", _update_scrollregion)

        # Make the inner frame's width track the scroll canvas's width so
        # children that pack(fill="x") fill the visible area instead of
        # collapsing to their natural width.
        def _resize_inner(event):
            scroll_canvas.itemconfig("scroll_inner", width=event.width)
        scroll_canvas.bind("<Configure>", _resize_inner)

        # Mouse wheel scrolling bound at the dialog level so it works
        # regardless of which child widget the cursor is over. Bound only
        # to this dialog so it doesn't interfere with other windows.
        def _on_mousewheel(event):
            scroll_canvas.yview_scroll(
                int(-1 * (event.delta / 120)), "units")
        dlg.bind("<MouseWheel>", _on_mousewheel)
        # Linux uses Button-4/5 instead of MouseWheel
        dlg.bind("<Button-4>",
                 lambda e: scroll_canvas.yview_scroll(-1, "units"))
        dlg.bind("<Button-5>",
                 lambda e: scroll_canvas.yview_scroll(1, "units"))

        # Inner padding lives on a sub-frame so it doesn't fight with the
        # scroll canvas geometry math.
        main_pad = tk.Frame(main, bg=bg)
        main_pad.pack(fill="both", expand=True, padx=10, pady=10)
        main = main_pad  # rest of method uses `main` as before

        # Sandbox banner if any preset is sandbox
        any_sandbox = any(fp.get('_preset', {}).get('sandbox')
                          for fp in fingerprints)
        if any_sandbox:
            sb_banner = tk.Label(main,
                text="\u26a0 Sandbox preset \u2014 non-standard setup, "
                     "compare with caution",
                bg="#443300", fg="#FFCC00", font=("Helvetica", 9, "bold"),
                padx=8, pady=4)
            sb_banner.pack(fill="x", pady=(0, 5))

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

        # Chart mode toggle (Overlay always; Bars when 2+; Difference only for 2)
        chart_mode = tk.StringVar(value="overlay")
        # Y-axis scale toggle. dB is the underlying storage format and works
        # well for showing wide dynamic range; linear amplitude
        # (= 10^(dB/20)) is closer to perceived loudness and prevents the
        # visual compression where -6 dB (50% amplitude) appears at 90% bar
        # height. Difference mode always stays in dB because deltas in dB
        # are easier to interpret than deltas in raw amplitude.
        scale_mode = tk.StringVar(value="dB")
        diff_frame = tk.Frame(toggle_frame, bg=bg)
        if len(fingerprints) >= 2:
            diff_frame.pack(side="right", padx=(10, 0))
            tk.Label(diff_frame, text="Chart:", bg=bg, fg=fg,
                     font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
            tk.Radiobutton(diff_frame, text="Overlay", variable=chart_mode,
                            value="overlay", bg=bg, fg=fg, selectcolor=bg,
                            font=("Helvetica", 9),
                            command=lambda: refresh_all()).pack(side="left")
            tk.Radiobutton(diff_frame, text="Bars", variable=chart_mode,
                            value="bars", bg=bg, fg=fg, selectcolor=bg,
                            font=("Helvetica", 9),
                            command=lambda: refresh_all()).pack(side="left")
            if len(fingerprints) == 2:
                tk.Radiobutton(diff_frame, text="Difference",
                                variable=chart_mode, value="difference",
                                bg=bg, fg=fg, selectcolor=bg,
                                font=("Helvetica", 9),
                                command=lambda: refresh_all()).pack(side="left")

        # Scale toggle — visible for single AND multi-preset views.
        scale_frame = tk.Frame(toggle_frame, bg=bg)
        scale_frame.pack(side="right", padx=(10, 0))
        tk.Label(scale_frame, text="Scale:", bg=bg, fg=fg,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
        tk.Radiobutton(scale_frame, text="dB", variable=scale_mode,
                        value="dB", bg=bg, fg=fg, selectcolor=bg,
                        font=("Helvetica", 9),
                        command=lambda: refresh_all()).pack(side="left")
        tk.Radiobutton(scale_frame, text="Linear", variable=scale_mode,
                        value="linear", bg=bg, fg=fg, selectcolor=bg,
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
        chart_frame = tk.LabelFrame(main, text="Harmonic Data", bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        chart_frame.pack(fill="both", expand=True, pady=(0, 6))

        chart_cv = tk.Canvas(chart_frame, bg="white", highlightthickness=0,
                              height=180)
        chart_cv.pack(fill="both", expand=True, padx=5, pady=5)

        legend_frame = tk.Frame(chart_frame, bg=bg)
        legend_frame.pack(fill="x", padx=5, pady=(0, 3))

        # Detail label shown when a legend entry is clicked
        detail_var = tk.StringVar(value="")
        detail_label = tk.Label(chart_frame, textvariable=detail_var,
                                bg=bg, fg="#666666", font=("Helvetica", 8),
                                anchor="w", justify="left")

        # Cross-widget "selected preset" state. Clicking a chart line, a quad
        # dot, or a legend entry selects that preset; clicking the same item
        # again toggles the selection off. The real set_selected is installed
        # later, once draw_chart and _quad_redraw exist.
        selected_idx = [None]
        legend_labels = []
        legend_default_bg = bg
        set_selected = lambda idx: None  # placeholder

        def _show_preset_detail(fp):
            p = fp.get('_preset', {})
            parts = []
            horn = " ".join(filter(None, [p.get('horn_make', ''),
                                          p.get('horn_model', '')]))
            if horn:
                parts.append(horn)
            serial = p.get('serial', '')
            if serial:
                parts.append(f"#{serial}")
            ht = p.get('horn_type', '')
            if ht:
                parts.append(f"({ht})")
            player = p.get('player', '')
            if player:
                parts.append(f"\u2014 {player}")
            mpc = p.get('mouthpiece', '')
            if mpc:
                parts.append(f"| mpc: {mpc}")
            reed = p.get('reed', '')
            if reed:
                parts.append(f"| reed: {reed}")
            mic_t = p.get('mic_type', '')
            mic_m = p.get('mic_model', '')
            mic = mic_m or mic_t
            if mic:
                parts.append(f"| mic: {mic}")
            mic_pos = p.get('mic_position', '')
            if mic_pos:
                parts.append(f"@ {mic_pos}")
            detail_var.set(" ".join(parts) if parts else fp.get('_name', ''))
            detail_label.pack(fill="x", padx=5, pady=(0, 3))
            # Highlight this preset across the legend, chart, and quad map.
            # set_selected is installed later but Python looks it up at call
            # time, so this works as long as the user can't click before the
            # real implementation is bound (which they can't).
            try:
                idx = fingerprints.index(fp)
            except ValueError:
                idx = None
            set_selected(idx)

        # Legend wraps to multiple rows when there are many presets so long
        # names don't get clipped on the right edge of the chart frame.
        legend_per_row = 3 if len(fingerprints) > 3 else len(fingerprints)
        legend_per_row = max(1, legend_per_row)
        legend_row = None
        for i, fp in enumerate(fingerprints):
            if i % legend_per_row == 0:
                legend_row = tk.Frame(legend_frame, bg=bg)
                legend_row.pack(fill="x")
            color = chart_colors[i % len(chart_colors)]
            tk.Label(legend_row, text="\u25a0", fg=color, bg=bg,
                     font=("Helvetica", 12)).pack(side="left")
            name = fp['_name']
            if len(name) > 30:
                name = name[:29] + "\u2026"
            lbl = tk.Label(legend_row, text=name, bg=legend_default_bg, fg=fg,
                           font=("Helvetica", 9), cursor="hand2",
                           padx=4)
            lbl.pack(side="left", padx=(0, 12))
            lbl.bind("<Button-1>", lambda e, f=fp: _show_preset_detail(f))
            legend_labels.append(lbl)

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
            mode = chart_mode.get()
            is_diff = (mode == "difference" and len(fingerprints) == 2)
            is_bars = (mode == "bars" and len(fingerprints) >= 2)
            # Difference mode always renders in dB regardless of scale toggle
            # because deltas in dB are easier to interpret than deltas in
            # raw amplitude (where the same dB shift looks very different
            # at the loud end vs the quiet end).
            is_linear = (scale_mode.get() == "linear" and not is_diff)

            margin_l, margin_r, margin_t, margin_b = 40, 10, 10, 25
            cw = w - margin_l - margin_r
            ch = h - margin_t - margin_b

            all_db = [d.get('harmonics_db', []) for d in data]
            max_h = max((len(db) for db in all_db), default=2)
            max_h = max(max_h, 2)

            # X position of harmonic index hi. Bars mode uses centered groups
            # so the first group has space on its left; line/diff modes anchor
            # H1 at the left edge so curves use the full width.
            group_width = cw / max_h if max_h > 0 else cw
            def harmonic_x(hi):
                if is_bars:
                    return margin_l + group_width * (hi + 0.5)
                if max_h > 1:
                    return margin_l + cw * hi / (max_h - 1)
                return margin_l + cw / 2

            # Y position helper. In dB mode the chart spans -60..0 dB
            # linearly. In linear amplitude mode the chart spans 0..1
            # amplitude (= 10^(dB/20)), which gives a closer match to
            # perceived loudness — a -6 dB harmonic (50% amplitude) sits
            # at 50% bar height instead of 90%. Default args capture the
            # range so the later db_min/db_max reassignment for difference
            # mode (below) can't leak into this closure.
            db_min, db_max = -60.0, 0.0
            def db_to_y(db_val, _mn=db_min, _mx=db_max, _lin=is_linear):
                clamped = max(_mn, min(_mx, db_val))
                if _lin:
                    amp = 10.0 ** (clamped / 20.0)
                    return margin_t + ch * (1.0 - amp)
                return margin_t + ch * (1.0 - (clamped - _mn) / (_mx - _mn))

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
            elif is_linear:
                # Linear amplitude grid at 0%, 25%, 50%, 75%, 100%
                for amp_pct in (0, 25, 50, 75, 100):
                    amp = amp_pct / 100.0
                    y = margin_t + ch * (1.0 - amp)
                    chart_cv.create_line(margin_l, y, w - margin_r, y,
                                          fill="#DDDDDD", width=1)
                    chart_cv.create_text(margin_l - 4, y,
                                          text=f"{amp_pct}%",
                                          fill="#888888",
                                          font=("Helvetica", 7), anchor="e")
                # Axis label so users know what's plotted
                chart_cv.create_text(8, h // 2, text="amp",
                                      fill="#888888",
                                      font=("Helvetica", 7), angle=90)
            else:
                for db in range(-60, 1, 10):
                    y = margin_t + ch * (1.0 - (db - db_min) / (db_max - db_min))
                    chart_cv.create_line(margin_l, y, w - margin_r, y,
                                          fill="#DDDDDD", width=1)
                    chart_cv.create_text(margin_l - 4, y, text=f"{db}",
                                          fill="#888888", font=("Helvetica", 7),
                                          anchor="e")

            for hi in range(max_h):
                x = harmonic_x(hi)
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
                        x = harmonic_x(hi)
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
            elif is_bars:
                # Grouped bars at each harmonic position, one bar per preset.
                # Useful for 3+ preset comparisons where overlay lines tangle.
                n_presets = len(data)
                if n_presets > 0:
                    # 85% of group goes to bars, 15% gap between groups
                    bar_w = max(1.0, (group_width * 0.85) / n_presets)
                    chart_bottom_y = margin_t + ch
                    for hi in range(max_h):
                        center_x = harmonic_x(hi)
                        total_w = bar_w * n_presets
                        start_x = center_x - total_w / 2
                        for i, d in enumerate(data):
                            color = chart_colors[i % len(chart_colors)]
                            db_list = d.get('harmonics_db', [])
                            if hi >= len(db_list):
                                continue
                            bar_top_y = db_to_y(db_list[hi])
                            x_left = start_x + i * bar_w
                            x_right = x_left + bar_w * 0.92  # tiny gap inside group
                            tag = f"fp_{i}"
                            is_selected = (selected_idx[0] == i)
                            outline_color = "#000000" if is_selected else ""
                            outline_w = 2 if is_selected else 0
                            chart_cv.create_rectangle(
                                x_left, bar_top_y, x_right, chart_bottom_y,
                                fill=color, outline=outline_color,
                                width=outline_w, tags=(tag,))
            else:
                # Draw the selected line LAST so it sits on top of the others.
                draw_order = sorted(
                    range(len(data)),
                    key=lambda i: 1 if i == selected_idx[0] else 0)
                for i in draw_order:
                    d = data[i]
                    color = chart_colors[i % len(chart_colors)]
                    db_list = d.get('harmonics_db', [])
                    if len(db_list) < 2:
                        continue
                    tag = f"fp_{i}"
                    is_selected = (selected_idx[0] == i)
                    line_width = 4 if is_selected else 2
                    dot_r = 5 if is_selected else 3

                    points = []
                    for hi, db in enumerate(db_list):
                        points.extend([harmonic_x(hi), db_to_y(db)])

                    if len(points) >= 4:
                        chart_cv.create_line(*points, fill=color,
                                              width=line_width,
                                              smooth=True, tags=(tag,))
                        for j in range(0, len(points), 2):
                            chart_cv.create_oval(
                                points[j] - dot_r, points[j + 1] - dot_r,
                                points[j] + dot_r, points[j + 1] + dot_r,
                                fill=color, outline="", tags=(tag,))

        def _on_chart_click(event):
            items = chart_cv.find_overlapping(
                event.x - 6, event.y - 6, event.x + 6, event.y + 6)
            for item_id in items:
                for tag in chart_cv.gettags(item_id):
                    if tag.startswith("fp_"):
                        idx = int(tag[3:])
                        if idx < len(fingerprints):
                            _show_preset_detail(fingerprints[idx])
                            return

        chart_cv.bind("<Configure>", draw_chart)
        chart_cv.bind("<Button-1>", _on_chart_click)

        # --- Character Map (Warmth × Brightness 2D scatter) ---
        # Only meaningful with 2+ presets; a single dot tells you nothing.
        # Reassigned inside the conditional below so refresh_all() can call
        # _quad_redraw() unconditionally.
        _quad_redraw = lambda *_: None

        if len(fingerprints) >= 2:
            quad_frame = tk.LabelFrame(
                main, text="Character Map (Warmth \u00d7 Brightness)",
                bg=bg, fg=fg, font=("Helvetica", 10, "bold"))
            # expand=True so the chart grows when the user enlarges the
            # dialog. The canvas itself also expands.
            quad_frame.pack(fill="both", expand=True, pady=(0, 6))

            quad_top = tk.Frame(quad_frame, bg=bg)
            quad_top.pack(fill="x", padx=5, pady=(2, 0))
            tk.Label(quad_top, text="Y axis (brightness):",
                      bg=bg, fg=fg, font=("Helvetica", 8)).pack(side="left")

            # Default to Complexity. H4-H5 mean is much more correlated
            # with Warmth (r ~ 0.83 within tenors) because both metrics
            # share a denominator (both measured in dB relative to the
            # fundamental). Complexity (spectral flatness) is the most
            # independent of Warmth (r ~ 0.16 within altos), so it makes
            # the chart actually function as a 2D character space rather
            # than a diagonal trend.
            quad_brightness_var = tk.StringVar(value="Complexity")
            quad_brightness_combo = ttk.Combobox(
                quad_top, textvariable=quad_brightness_var,
                values=["Complexity", "H4-H5 mean", "Rolloff (inverted)"],
                state="readonly", width=18, font=("Helvetica", 8))
            quad_brightness_combo.pack(side="left", padx=(4, 0))
            quad_brightness_combo.bind(
                "<<ComboboxSelected>>", lambda e: draw_quadrant())

            # height is the initial preferred size; expand=True lets it grow
            # to fill any extra vertical space when the dialog is enlarged.
            quad_cv = tk.Canvas(quad_frame, bg="white",
                                 highlightthickness=0, height=190)
            quad_cv.pack(fill="both", expand=True, padx=5, pady=5)

            def get_brightness(fp):
                """Y-axis value for a fingerprint based on selected measure."""
                measure = quad_brightness_var.get()
                if measure == "Complexity":
                    return fp.get('descriptors', {}).get('richness', 0.0)
                if measure == "Rolloff (inverted)":
                    r = fp.get('rolloff_rate')
                    if r is None:
                        return None
                    # Map ~1.0 dB/H (very bright) → 1.0,
                    # ~4.0 dB/H (very dark) → 0.0
                    return max(0.0, min(1.0, (4.0 - r) / 3.0))
                # Default: average of H4 and H5 in dB (raw signal)
                hd = fp.get('harmonics_db', [])
                if len(hd) >= 5:
                    return (hd[3] + hd[4]) / 2.0
                return None

            def draw_quadrant(event=None):
                quad_cv.delete("all")
                w = quad_cv.winfo_width()
                h = quad_cv.winfo_height()
                if w < 80 or h < 80:
                    return

                margin_l, margin_r, margin_t, margin_b = 56, 14, 14, 36
                cw = w - margin_l - margin_r
                ch = h - margin_t - margin_b
                chart_left = margin_l
                chart_right = margin_l + cw
                chart_top = margin_t
                chart_bottom = margin_t + ch

                # Compute Y values for all presets up front
                y_values = [get_brightness(fp) for fp in fingerprints]
                valid_ys = [v for v in y_values if v is not None]
                if not valid_ys:
                    quad_cv.create_text(
                        w // 2, h // 2,
                        text="No brightness data for selected measure",
                        fill="#888888", font=("Helvetica", 9))
                    return

                # Y-axis range
                measure = quad_brightness_var.get()
                if measure in ("Complexity", "Rolloff (inverted)"):
                    y_min, y_max = 0.0, 1.0
                else:
                    # H4-H5 mean: auto-scale with padding
                    y_min_data = min(valid_ys)
                    y_max_data = max(valid_ys)
                    pad = max(1.5, (y_max_data - y_min_data) * 0.25)
                    y_min = y_min_data - pad
                    y_max = y_max_data + pad
                    if y_max - y_min < 4.0:
                        center = (y_max + y_min) / 2
                        y_min = center - 2.0
                        y_max = center + 2.0

                # Outer border
                quad_cv.create_rectangle(
                    chart_left, chart_top, chart_right, chart_bottom,
                    outline="#CCCCCC", width=1)

                # Quadrant midlines (warmth=0.5, brightness midpoint)
                x_mid_pixel = chart_left + cw * 0.5
                y_mid_pixel = chart_top + ch * 0.5
                quad_cv.create_line(
                    x_mid_pixel, chart_top, x_mid_pixel, chart_bottom,
                    fill="#BBBBBB", width=1, dash=(3, 3))
                quad_cv.create_line(
                    chart_left, y_mid_pixel, chart_right, y_mid_pixel,
                    fill="#BBBBBB", width=1, dash=(3, 3))

                # Quadrant corner labels (faded)
                quad_cv.create_text(
                    chart_right - 4, chart_top + 4,
                    text="warm + bright", fill="#BBBBBB",
                    font=("Helvetica", 7, "italic"), anchor="ne")
                quad_cv.create_text(
                    chart_left + 4, chart_top + 4,
                    text="thin + bright", fill="#BBBBBB",
                    font=("Helvetica", 7, "italic"), anchor="nw")
                quad_cv.create_text(
                    chart_right - 4, chart_bottom - 4,
                    text="warm + dark", fill="#BBBBBB",
                    font=("Helvetica", 7, "italic"), anchor="se")
                quad_cv.create_text(
                    chart_left + 4, chart_bottom - 4,
                    text="thin + dark", fill="#BBBBBB",
                    font=("Helvetica", 7, "italic"), anchor="sw")

                # X-axis ticks and label
                for px, lbl in [(0.0, "0%"), (0.5, "50%"), (1.0, "100%")]:
                    x = chart_left + cw * px
                    quad_cv.create_text(
                        x, chart_bottom + 10, text=lbl,
                        fill="#888888", font=("Helvetica", 7))
                quad_cv.create_text(
                    chart_left + cw / 2, chart_bottom + 24,
                    text="\u2190 thin    Warmth    warm \u2192",
                    fill="#666666", font=("Helvetica", 8))

                # Y-axis ticks and label
                for frac in (0.0, 0.5, 1.0):
                    val = y_min + (y_max - y_min) * frac
                    y = chart_top + ch * (1.0 - frac)
                    if measure == "H4-H5 mean":
                        tick_text = f"{val:+.1f}"
                    else:
                        tick_text = f"{val:.2f}"
                    quad_cv.create_text(
                        chart_left - 4, y, text=tick_text,
                        fill="#888888", font=("Helvetica", 7), anchor="e")
                if measure == "H4-H5 mean":
                    y_title = "\u2190 dark   H4\u2013H5 dB   bright \u2192"
                elif measure == "Complexity":
                    y_title = "\u2190 pure   Complexity   complex \u2192"
                else:
                    y_title = "\u2190 dark   Rolloff (inv)   bright \u2192"
                quad_cv.create_text(
                    14, chart_top + ch / 2, text=y_title,
                    fill="#666666", font=("Helvetica", 8), angle=90)

                # Plot dots — colors match the harmonic chart legend just
                # above this frame, so no inline labels are needed. Click any
                # dot to see preset detail in the detail label and highlight
                # the matching preset across the legend, harmonic chart, and
                # this Character Map. Selected dot draws last so it sits on
                # top of any overlapping unselected dots.
                draw_order = sorted(
                    range(len(fingerprints)),
                    key=lambda i: 1 if i == selected_idx[0] else 0)
                for i in draw_order:
                    fp = fingerprints[i]
                    yv = y_values[i]
                    if yv is None:
                        continue
                    warmth = fp.get('descriptors', {}).get('warmth', 0.0)
                    color = chart_colors[i % len(chart_colors)]
                    is_selected = (selected_idx[0] == i)
                    r = 9 if is_selected else 6
                    outline_color = "#000000" if is_selected else "white"
                    outline_w = 2 if is_selected else 2
                    x = chart_left + cw * max(0.0, min(1.0, warmth))
                    y = chart_top + ch * (
                        1.0 - max(0.0, min(1.0, (yv - y_min) / (y_max - y_min))))
                    # Keep dots fully inside the chart frame
                    x = max(chart_left + r, min(chart_right - r, x))
                    y = max(chart_top + r, min(chart_bottom - r, y))
                    tag = f"qm_{i}"  # quadrant marker tag
                    quad_cv.create_oval(
                        x - r, y - r, x + r, y + r,
                        fill=color, outline=outline_color,
                        width=outline_w, tags=(tag,))

            def _on_quad_click(event):
                items = quad_cv.find_overlapping(
                    event.x - 6, event.y - 6, event.x + 6, event.y + 6)
                for item_id in items:
                    for tag in quad_cv.gettags(item_id):
                        if tag.startswith("qm_"):
                            idx = int(tag[3:])
                            if idx < len(fingerprints):
                                _show_preset_detail(fingerprints[idx])
                                return

            quad_cv.bind("<Configure>", draw_quadrant)
            quad_cv.bind("<Button-1>", _on_quad_click)
            _quad_redraw = draw_quadrant

        # Install the real cross-widget selection handler now that
        # draw_chart and _quad_redraw exist. Toggles selection on/off when
        # the same item is clicked twice. Restyles legend labels and
        # redraws both charts to apply visual highlight.
        def _set_selected_impl(idx):
            if idx is not None and idx == selected_idx[0]:
                selected_idx[0] = None  # toggle off
            else:
                selected_idx[0] = idx
            for i, lbl in enumerate(legend_labels):
                if i == selected_idx[0]:
                    try:
                        lbl.config(font=("Helvetica", 9, "bold"),
                                    bg="#FFEB3B")
                    except tk.TclError:
                        pass
                else:
                    try:
                        lbl.config(font=("Helvetica", 9),
                                    bg=legend_default_bg)
                    except tk.TclError:
                        pass
            draw_chart()
            _quad_redraw()
        set_selected = _set_selected_impl

        # --- Descriptor table (rebuilt on view change) ---
        table_label = "Descriptors" if is_single else "Descriptor Comparison"
        table_frame = tk.LabelFrame(main, text=table_label, bg=bg, fg=fg,
                                     font=("Helvetica", 10, "bold"))
        table_frame.pack(fill="x", pady=(0, 6))
        table_inner = tk.Frame(table_frame, bg=bg)
        table_inner.pack(fill="x", padx=5, pady=5)

        all_desc_labels = [
            ("H. Complexity", "richness"),
            ("Warmth", "warmth"),
            ("Even H. Bal.", "even_odd"),
            ("Rolloff Shape", "rolloff_shape"),
            ("Evenness", "evenness"),
        ]
        analysis_desc = self.settings.get("toner_settings", {}).get(
            "analysis_descriptors", {"richness": True, "warmth": True})
        desc_labels = [(label, key) for label, key in all_desc_labels
                       if analysis_desc.get(key, key in (
                           "richness", "warmth", "even_odd"))]
        # Pole labels for descriptor direction (low% → high%)
        desc_poles = {
            "richness": ("pure", "complex"),
            "warmth": ("thin", "warm"),
            "even_odd": ("odd", "even"),
            "rolloff_shape": ("smooth", "peaked"),
            "evenness": ("variable", "even"),
        }

        # Rolloff rates and mic types for each preset (for table and mismatch check)
        _rolloff_rates = [fp.get('rolloff_rate') for fp in fingerprints]
        _mic_types = [fp.get('mic_type', '') for fp in fingerprints]

        # --- Analysis text ---
        analysis_frame = tk.LabelFrame(main, text="Analysis", bg=bg, fg=fg,
                                        font=("Helvetica", 10, "bold"))
        analysis_frame.pack(fill="both", expand=True)
        analysis_inner = tk.Frame(analysis_frame, bg=bg)
        analysis_inner.pack(fill="both", expand=True, padx=5, pady=5)
        analysis_text = tk.Text(analysis_inner, height=10, wrap="word",
                                 font=("Helvetica", 9), bg="white",
                                 relief="flat", padx=8, pady=5)
        analysis_scroll = tk.Scrollbar(analysis_inner, command=analysis_text.yview)
        analysis_text.configure(yscrollcommand=analysis_scroll.set)
        analysis_scroll.pack(side="right", fill="y")
        analysis_text.pack(side="left", fill="both", expand=True)

        def rebuild_table():
            for w in table_inner.winfo_children():
                w.destroy()

            data = get_data_for_view()

            # Compact mode for many-preset comparisons: smaller cells/font
            # so 6+ columns still fit inside the dialog width.
            compact = len(fingerprints) > 3
            data_width = 11 if compact else 14
            data_font = ("Helvetica", 8) if compact else ("Helvetica", 9)
            head_width = 11 if compact else 14
            head_font = ("Helvetica", 8, "bold") if compact else ("Helvetica", 9, "bold")
            head_trunc = 11 if compact else 15
            label_width = 18 if compact else 20
            label_font = ("Helvetica", 8) if compact else ("Helvetica", 9)
            pct_sep = " " if compact else "  "

            def _is_real(v):
                """Reject None and NaN; keep real numbers."""
                if v is None:
                    return False
                if isinstance(v, float) and (v != v):  # NaN
                    return False
                return True

            header = tk.Frame(table_inner, bg=bg)
            header.pack(fill="x")
            tk.Label(header, text="", width=label_width, bg=bg, fg=fg,
                     font=head_font, anchor="w").pack(side="left")
            for fp in fingerprints:
                tk.Label(header, text=fp['_name'][:head_trunc],
                         width=head_width, bg=bg, fg=fg,
                         font=head_font, anchor="center").pack(side="left")

            for label, key in desc_labels:
                row = tk.Frame(table_inner, bg=bg)
                row.pack(fill="x")
                poles = desc_poles.get(key)
                row_label = f"{label} ({poles[0]}\u2194{poles[1]})" if poles else label
                tk.Label(row, text=row_label, width=label_width, bg=bg, fg=fg,
                         font=label_font, anchor="w").pack(side="left")

                values = [d.get('descriptors', {}).get(key, 0) for d in data]
                real_vals = [v for v in values if _is_real(v) and v > 0]
                max_val = max(real_vals) if real_vals else 0
                min_val = min(real_vals) if real_vals else 0

                for val in values:
                    if (_is_real(val) and val > 0 and len(real_vals) > 1
                            and max_val != min_val):
                        if val == max_val:
                            val_fg = "#006600"
                        elif val == min_val:
                            val_fg = "#880000"
                        else:
                            val_fg = fg
                    else:
                        val_fg = fg
                    if not _is_real(val) or val <= 0:
                        text = "\u2014"
                    else:
                        text = f"{val:.0%}"
                    # Add percentile when population data available
                    if (population_stats and population_stats['count'] >= 3
                            and _is_real(val) and val > 0
                            and view_mode.get() == "average"):
                        sorted_vals = population_stats[
                            'descriptor_values'].get(key, [])
                        pct = percentile_rank(val, sorted_vals)
                        if pct is not None:
                            text += f"{pct_sep}P{pct}"
                    tk.Label(row, text=text, width=data_width, bg=bg, fg=val_fg,
                             font=data_font, anchor="center").pack(side="left")

            # Population context note
            if (population_stats and population_stats['count'] >= 3
                    and view_mode.get() == "average"):
                pop_note = tk.Frame(table_inner, bg=bg)
                pop_note.pack(fill="x")
                tk.Label(pop_note,
                         text=f"P = percentile among "
                              f"{population_stats['count']} "
                              f"{population_stats['sax_type'].lower()} "
                              f"presets",
                         bg=bg, fg="#999999",
                         font=("Helvetica", 7)).pack(anchor="e")

            # Rolloff rate and mic type rows (only in horn average view)
            if view_mode.get() == "average":
                row = tk.Frame(table_inner, bg=bg)
                row.pack(fill="x")
                tk.Label(row, text="Rec. Quality", width=label_width, bg=bg,
                         fg=fg, font=label_font, anchor="w").pack(side="left")
                for rate, mt in zip(_rolloff_rates, _mic_types):
                    if not _is_real(rate):
                        text = "\u2014"
                        val_fg = fg
                    else:
                        # Compact mode drops the "/H" to save horizontal space
                        unit = "dB" if compact else "dB/H"
                        text = f"{rate:.1f} {unit}"
                        # Mic-class-aware threshold so ribbons and dynamics
                        # don't get flagged as "bad" for normal HF rolloff.
                        threshold = get_rolloff_threshold(mt)
                        val_fg = "#880000" if rate > threshold else fg
                        if population_stats and population_stats['count'] >= 3:
                            pct = percentile_rank(
                                rate, population_stats['rolloff_values'])
                            if pct is not None:
                                text += f"{pct_sep}P{pct}"
                    tk.Label(row, text=text, width=data_width, bg=bg, fg=val_fg,
                             font=data_font, anchor="center").pack(side="left")

                row2 = tk.Frame(table_inner, bg=bg)
                row2.pack(fill="x")
                tk.Label(row2, text="Mic Type", width=label_width, bg=bg,
                         fg=fg, font=label_font, anchor="w").pack(side="left")
                for mt in _mic_types:
                    text = mt.capitalize() if mt else "\u2014"
                    tk.Label(row2, text=text, width=data_width, bg=bg, fg=fg,
                             font=data_font, anchor="center").pack(side="left")

        def rebuild_analysis():
            data = get_data_for_view()
            analysis_text.configure(state="normal")
            analysis_text.delete("1.0", tk.END)

            mode = view_mode.get()
            prefix = ""
            if mode == "per_note":
                prefix = f"For {note_var.get()}: "

            # Helper: session/capture context for a preset
            def _preset_ctx(fp):
                p = fp.get('_preset', {})
                n_sess = len([s for s in p.get('sessions', [])
                              if s.get('captures')]) if p else 0
                return f"{fp['capture_count']} caps, {n_sess} sess"

            lines = []

            if is_single:
                # Single preset/session view
                fp = fingerprints[0]
                d = data[0].get('descriptors', {})
                lines.append(f"{fp['_name']} ({_preset_ctx(fp)})")

                # Data quality summary
                note_count = fp.get('note_count', 0)
                per_note = fp.get('per_note', {})
                if per_note:
                    midi_vals = [(n, note_name_to_midi(n))
                                 for n in per_note.keys()]
                    midi_vals = [(n, m) for n, m in midi_vals
                                 if m is not None]
                    if midi_vals:
                        midi_vals.sort(key=lambda x: x[1])
                        lo, hi = midi_vals[0][0], midi_vals[-1][0]
                        lines.append(
                            f"{note_count} notes ({lo}\u2013{hi}), "
                            f"{fp.get('capture_count', 0)} captures")
                    else:
                        lines.append(
                            f"{note_count} notes, "
                            f"{fp.get('capture_count', 0)} captures")
                if note_count < MIN_PRESET_NOTES:
                    lines.append(
                        f"\u26a0 Sparse coverage ({note_count} notes, "
                        f"{MIN_PRESET_NOTES}+ recommended)")

                # Out-of-range note detection
                sax_type = fp.get('_preset', {}).get('horn_type', '')
                oor = find_out_of_range_notes(
                    per_note.keys(), sax_type) if per_note else []
                if oor:
                    oor_sorted = sorted(oor, key=lambda n:
                        note_name_to_midi(n) or 0)
                    lines.append(
                        f"\u26a0 Out-of-range notes: "
                        f"{', '.join(oor_sorted)} "
                        f"(possible artifacts)")

                lines.append("")
                desc_parts = []
                for label, key in desc_labels:
                    val = d.get(key, 0)
                    if val > 0:
                        poles = desc_poles.get(key)
                        if poles:
                            pole = poles[1] if val >= 0.5 else poles[0]
                            desc_parts.append(f"{label}: {val:.0%} ({pole})")
                        else:
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

                # Population context
                if (population_stats and population_stats['count'] >= 3
                        and mode == "average"):
                    pop_parts = []
                    for label, key in desc_labels:
                        val = d.get(key, 0)
                        if val <= 0:
                            continue
                        sorted_vals = population_stats[
                            'descriptor_values'].get(key, [])
                        pct = percentile_rank(val, sorted_vals)
                        if pct is None:
                            continue
                        if pct <= 15:
                            tag = "low"
                        elif pct <= 35:
                            tag = "below avg"
                        elif pct <= 65:
                            tag = "mid-range"
                        elif pct <= 85:
                            tag = "above avg"
                        else:
                            tag = "high"
                        pop_parts.append(f"{label.lower()} P{pct} ({tag})")
                    if pop_parts:
                        st = population_stats['sax_type'].lower()
                        lines.append("")
                        lines.append(
                            f"Among {population_stats['count']} "
                            f"{st} presets: " + ", ".join(pop_parts))

            else:
                # Context header for multi-preset
                ctx_parts = [f"{fp['_name']} ({_preset_ctx(fp)})"
                             for fp in fingerprints]
                lines.append(f"Comparing: {', '.join(ctx_parts)}")
                lines.append("")
                if len(fingerprints) >= 3:
                    lines.append(
                        "Tip: switch the chart to Bars to see harmonic "
                        "amplitudes side by side, and use the Character Map "
                        "below to see where each preset sits on the warmth "
                        "\u00d7 brightness axes.")
                    lines.append("")

            if len(fingerprints) == 2:
                n1 = fingerprints[0]['_name']
                n2 = fingerprints[1]['_name']
                da = data[0].get('descriptors', {})
                db_d = data[1].get('descriptors', {})
                h_a = data[0].get('harmonics_db', [])
                h_b = data[1].get('harmonics_db', [])

                if not da and not db_d:
                    lines.append(f"{prefix}No data for this note in either preset.")
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
                                    f"{prefix}Shifts concentrated in upper harmonics. "
                                    "Based on limited data, this pattern may suggest "
                                    "neck or mouthpiece differences, though other "
                                    "factors could contribute. Low harmonics are "
                                    "relatively similar.")
                            else:
                                lines.append(
                                    f"{prefix}Broadband shifts across H2\u2013H12. "
                                    "This pattern is often associated with mouthpiece "
                                    "or player differences, but the causes of harmonic "
                                    "variation are not fully understood. In our test "
                                    "data, mouthpiece changes routinely shift "
                                    "descriptors by 10\u201325% on the same horn.")
                        elif low_shift > 2.0 and low_shift > mid_shift:
                            lines.append(
                                f"{prefix}Shifts concentrated in low harmonics "
                                "(H2\u2013H4). Research suggests this range is "
                                "influenced more by the bore than by the mouthpiece, "
                                "but we are still learning what drives these differences.")
                        # Pointer to the Character Map for spatial intuition
                        if abs(biggest[1]) > 2.0:
                            lines.append(
                                f"{prefix}See the Character Map for where these "
                                "two presets sit on the warmth \u00d7 brightness "
                                "axes \u2014 they're independent, so the same shift "
                                "can come from very different timbral changes.")

                    # Player context
                    p1 = fingerprints[0].get('_preset', {}).get('player', '')
                    p2 = fingerprints[1].get('_preset', {}).get('player', '')
                    if p1 and p2:
                        if p1.lower() == p2.lower():
                            lines.append(
                                f"\nSame player ({p1}) \u2014 differences likely "
                                "reflect horn, neck, mouthpiece, or reed rather "
                                "than embouchure, though day-to-day variation in "
                                "a player's sound is always a factor.")
                        else:
                            lines.append(
                                f"\nDifferent players ({p1} vs {p2}) \u2014 "
                                "player and mouthpiece effects can be as large as "
                                "horn differences. Low harmonics (H1\u2013H4) tend to "
                                "be more stable across players, but treat all "
                                "cross-player comparisons as suggestive, not definitive.")
            elif len(fingerprints) > 2:
                any_spread = False
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
                        any_spread = True
                if any_spread:
                    lines.append("")
                    lines.append(
                        f"{prefix}Remember: warmth and brightness are "
                        "independent axes. Two presets with the same warmth "
                        "can have very different brightness, and vice versa. "
                        "The Character Map shows the 2D spread; the Bars chart "
                        "shows the raw harmonic amplitudes that drive it.")

            # Check for sandbox, mic type, and rolloff mismatches
            if mode == "average":
                # Sandbox warning
                sandbox_names = [fp['_name'] for fp in fingerprints
                                 if fp.get('_preset', {}).get('sandbox')]
                if sandbox_names and len(sandbox_names) < len(fingerprints):
                    sb_list = ", ".join(sandbox_names)
                    lines.append("")
                    lines.append(
                        f"\u26a0 Sandbox preset(s): {sb_list}. "
                        "Sandbox data may use non-standard mic or "
                        "instrument setups \u2014 compare with caution.")
                # Mic type mismatch
                known_types = [mt for mt in _mic_types if mt]
                if len(set(known_types)) > 1:
                    type_list = ", ".join(
                        f"{fp['_name']}: {mt.capitalize()}"
                        for fp, mt in zip(fingerprints, _mic_types) if mt)
                    lines.append("")
                    lines.append(
                        f"\u26a0 Mic types differ ({type_list}). "
                        "Differences in complexity, warmth, and rolloff rate may "
                        "partly reflect the mic's frequency response rather than "
                        "the horn. Even/Odd Ratio is the most stable descriptor "
                        "across mic differences in our test data.")

                # Rolloff mismatch (still useful even with same mic type)
                valid_rates = [r for r in _rolloff_rates if r is not None]
                if len(valid_rates) >= 2:
                    rate_spread = max(valid_rates) - min(valid_rates)
                    if rate_spread > 1.0:
                        lines.append("")
                        lines.append(
                            "\u26a0 Recording quality differs significantly "
                            f"(rolloff spread: {rate_spread:.1f} dB/H). "
                            "Harmonic complexity comparison may reflect mic/room "
                            "differences rather than horn differences.")

            analysis_text.insert("1.0", "\n".join(lines) if lines else
                                 f"{prefix}Presets are similar \u2014 no major differences.")
            analysis_text.configure(state="disabled")

        def refresh_all():
            draw_chart()
            _quad_redraw()
            rebuild_table()
            rebuild_analysis()

        # Initial build
        rebuild_table()
        rebuild_analysis()

        btn_row = tk.Frame(main, bg=bg)
        btn_row.pack(pady=(5, 0))

        if is_single:
            def overlay_from_analysis():
                fp = fingerprints[0]
                self._toner_set_overlay(fp, fp.get('_name', 'Analysis'))
                dlg.destroy()
            tk.Button(btn_row, text="Overlay on Spectrum",
                      command=overlay_from_analysis).pack(
                          side="left", padx=(0, 5))
        def back_to_selection():
            dlg.destroy()
            self._toner_open_analyze_dialog()
        tk.Button(btn_row, text="\u2190 Back",
                  command=back_to_selection).pack(side="left", padx=(0, 5))

        # Maximize toggle — backup for users on Windows where the title bar's
        # maximize button can sometimes be missing or unreliable. Toggles
        # between 'zoomed' and 'normal' window states.
        def toggle_maximize():
            try:
                current = dlg.state()
            except tk.TclError:
                current = "normal"
            if current == "zoomed":
                dlg.state("normal")
            else:
                try:
                    dlg.state("zoomed")
                except tk.TclError:
                    # macOS Aqua doesn't support 'zoomed'; fall back to a
                    # large geometry that fills most of the screen.
                    sw = dlg.winfo_screenwidth()
                    sh = dlg.winfo_screenheight()
                    dlg.geometry(f"{sw - 40}x{sh - 80}+20+20")
        tk.Button(btn_row, text="Maximize",
                  command=toggle_maximize).pack(side="left", padx=(0, 5))

        tk.Button(btn_row, text="Close",
                  command=dlg.destroy).pack(side="left")

    # ------------------------------------------------------------------
    # GROUP REPORT
    # ------------------------------------------------------------------

    def _toner_show_group_report(self, preset_list):
        """Show an aggregated report across multiple profiles.

        Args:
            preset_list: list of (name, profile_data) tuples
        """
        grp = compute_group_fingerprint(preset_list)
        if grp['preset_count'] == 0:
            messagebox.showinfo("No Data",
                "Selected presets have no captures.")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Group Average")
        dlg.geometry("720x650")
        # Same reasoning as the Analyze dialog: skip transient so the user
        # gets the maximize button and an independent taskbar entry, and
        # allow free resize / fullscreen. Content scrolls vertically.
        dlg.resizable(True, True)
        dlg.minsize(640, 500)

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"

        # --- Scrollable main content (same pattern as Analyze) ---
        outer = tk.Frame(dlg, bg=bg)
        outer.pack(fill="both", expand=True)

        scroll_vbar = tk.Scrollbar(outer, orient="vertical")
        scroll_vbar.pack(side="right", fill="y")

        scroll_canvas = tk.Canvas(outer, bg=bg, highlightthickness=0,
                                   yscrollcommand=scroll_vbar.set)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll_vbar.config(command=scroll_canvas.yview)

        main = tk.Frame(scroll_canvas, bg=bg)
        scroll_canvas.create_window(
            (0, 0), window=main, anchor="nw", tags="scroll_inner")

        def _update_scrollregion(event=None):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        main.bind("<Configure>", _update_scrollregion)

        def _resize_inner(event):
            scroll_canvas.itemconfig("scroll_inner", width=event.width)
        scroll_canvas.bind("<Configure>", _resize_inner)

        def _on_mousewheel(event):
            scroll_canvas.yview_scroll(
                int(-1 * (event.delta / 120)), "units")
        dlg.bind("<MouseWheel>", _on_mousewheel)
        dlg.bind("<Button-4>",
                 lambda e: scroll_canvas.yview_scroll(-1, "units"))
        dlg.bind("<Button-5>",
                 lambda e: scroll_canvas.yview_scroll(1, "units"))

        main_pad = tk.Frame(main, bg=bg)
        main_pad.pack(fill="both", expand=True, padx=10, pady=10)
        main = main_pad

        # --- Header ---
        tk.Label(main, text=f"Group Average: {grp['preset_count']} presets",
                 bg=bg, fg=fg,
                 font=("Helvetica", 13, "bold")).pack(anchor="w")
        names = [n for n, _ in preset_list]
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
        for label, key in [("H. Complexity", "richness"),
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
        chart_frame = tk.LabelFrame(main, text="Harmonic Data",
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
                         for _, fp in grp['per_preset']),
                        default=0)
            max_h = max(max_h, len(grp.get('harmonics_db', [])))
            if max_h < 2:
                return

            for hi in range(max_h):
                x = margin_l + cw * hi / (max_h - 1)
                chart_cv.create_text(x, h - 5, text=f"H{hi+1}",
                                      fill="#888888", font=("Helvetica", 7))

            # Individual profiles (thin lines)
            for idx, (name, fp) in enumerate(grp['per_preset']):
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
            for idx, (name, _) in enumerate(grp['per_preset']):
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

        # --- Per-preset breakdown table ---
        tbl_frame = tk.LabelFrame(main, text="Per-Preset Breakdown",
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
        for text, w in [("Preset", 20), ("Sess", 5), ("Caps", 5),
                        ("Notes", 5), ("Cmpx", 5), ("Warm", 5)]:
            tk.Label(thdr, text=text, width=w, bg=bg, fg=fg,
                     font=("Helvetica", 8, "bold")).pack(side="left")

        for name, fp in grp['per_preset']:
            trow = tk.Frame(tbl_inner, bg=bg)
            trow.pack(fill="x")
            # Find session count from preset_list
            n_sess = 0
            for pn, pd in preset_list:
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
        ctx = (f"Group average across {grp['preset_count']} presets "
               f"({grp['total_captures']} total captures). "
               f"\u00b1 values reflect variation across presets, "
               f"not measurement noise within a preset.")
        tk.Label(main, text=ctx, bg=bg, fg="#888888",
                 font=("Helvetica", 7), wraplength=680,
                 justify="left").pack(anchor="w", pady=(0, 4))

        btn_row = tk.Frame(main, bg=bg)
        btn_row.pack(pady=(5, 0))

        def overlay_pick():
            menu = tk.Menu(dlg, tearoff=0)
            # Group average option
            menu.add_command(
                label=f"Group Average ({grp['preset_count']} presets)",
                command=lambda: [
                    self._toner_set_overlay(
                        grp, f"Group avg ({grp['preset_count']} presets)"),
                    dlg.destroy()])
            menu.add_separator()
            # Individual profiles
            for pname, fp in grp['per_preset']:
                menu.add_command(
                    label=pname,
                    command=lambda f=fp, n=pname: [
                        self._toner_set_overlay(f, n), dlg.destroy()])
            menu.tk_popup(btn_row.winfo_rootx(),
                          btn_row.winfo_rooty() - menu.index("end") * 20)

        tk.Button(btn_row, text="Overlay on Spectrum...",
                  command=overlay_pick).pack(side="left", padx=(0, 5))
        def back_to_selection():
            dlg.destroy()
            self._toner_open_analyze_dialog()
        tk.Button(btn_row, text="\u2190 Back",
                  command=back_to_selection).pack(side="left", padx=(0, 5))

        # Maximize toggle — backup for users on Windows where the title bar's
        # maximize button can sometimes be missing or unreliable. Toggles
        # between 'zoomed' and 'normal' window states.
        def toggle_maximize():
            try:
                current = dlg.state()
            except tk.TclError:
                current = "normal"
            if current == "zoomed":
                dlg.state("normal")
            else:
                try:
                    dlg.state("zoomed")
                except tk.TclError:
                    # macOS Aqua doesn't support 'zoomed'; fall back to a
                    # large geometry that fills most of the screen.
                    sw = dlg.winfo_screenwidth()
                    sh = dlg.winfo_screenheight()
                    dlg.geometry(f"{sw - 40}x{sh - 80}+20+20")
        tk.Button(btn_row, text="Maximize",
                  command=toggle_maximize).pack(side="left", padx=(0, 5))

        tk.Button(btn_row, text="Close",
                  command=dlg.destroy).pack(side="left")

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

            # Process capture if active. Captures still record raw harmonics
            # so the Analyze tool can compute deltas after the fact.
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
        """Import an audio file (WAV) — full guided flow with preset setup."""
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
        tk.Label(frame, text="Select a preset for this recording, or create one.\n"
                 "Add notes about the source (who, when, where recorded).",
                 bg=bg, fg=fg, font=("Helvetica", 9),
                 justify="left").pack(pady=(0, 10))

        # Profile selector
        all_presets = []
        for lib_name, lib_presets in self._toner_presets.items():
            if not isinstance(lib_presets, dict):
                continue
            for preset_name in lib_presets:
                all_presets.append((lib_name, preset_name))

        preset_frame = tk.Frame(frame, bg=bg)
        preset_frame.pack(fill="x", pady=(0, 5))

        use_existing = tk.BooleanVar(value=bool(all_presets))

        if all_presets:
            tk.Radiobutton(preset_frame, text="Existing preset:",
                           variable=use_existing, value=True,
                           bg=bg, fg=fg, font=("Helvetica", 10)).pack(
                               anchor="w")
            preset_var = tk.StringVar(
                value=f"[{all_presets[0][0]}] {all_presets[0][1]}" if all_presets else "")
            preset_list = [f"[{lib}] {name}" for lib, name in all_presets]
            preset_combo = ttk.Combobox(preset_frame, textvariable=preset_var,
                                      values=preset_list, state="readonly",
                                      width=40)
            preset_combo.pack(anchor="w", padx=(20, 0), pady=(0, 5))

        tk.Radiobutton(preset_frame, text="Create new preset...",
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

            if use_existing.get() and all_presets:
                # Use selected preset
                sel_text = preset_var.get()
                selected = None
                for lib, name in all_presets:
                    if f"[{lib}] {name}" == sel_text:
                        selected = (lib, name)
                        break
                if not selected:
                    messagebox.showinfo("Select Preset",
                        "Select a preset from the list.", parent=dlg)
                    return

                self._toner_active_library = selected[0]
                self._toner_active_preset = selected[1]

                # Sync sax type
                preset = self._toner_presets[selected[0]][selected[1]]
                sax_type = preset.get('horn_type', '')
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
                    'mic_type': preset.get('mic_type', ''),
                    'mic_model': preset.get('mic_model', ''),
                }

                dlg.destroy()
                self._toner_do_file_import(filepath, source_notes)
            else:
                # Need to create new preset first
                dlg.destroy()
                # Store filepath for after preset creation
                self._toner_pending_file_import = (filepath, source_notes)
                self._toner_new_preset_flow_then_import()

        btn_row = tk.Frame(frame, bg=bg)
        btn_row.pack(fill="x", pady=(5, 0))
        tk.Button(btn_row, text="Import && Analyze",
                  command=do_import).pack(side="left", padx=(0, 5))
        tk.Button(btn_row, text="Cancel",
                  command=dlg.destroy).pack(side="left")

    def _toner_new_preset_flow_then_import(self):
        """Create a new preset, then import the pending file."""
        # Reuse the existing new preset flow but override the callback
        dlg = tk.Toplevel(self.root)
        dlg.title("New Tone Preset")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Create Tone Preset for Audio Import", bg=bg, fg=fg,
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
        existing_libs = [k for k in self._toner_presets.keys()
                        if isinstance(self._toner_presets[k], dict)]
        if not existing_libs:
            existing_libs = [DEFAULT_LIBRARY]
        lib_var = tk.StringVar(value=existing_libs[0])
        ttk.Combobox(lib_row, textvariable=lib_var,
                      values=existing_libs, width=20).pack(
            side="left", fill="x", expand=True)

        add_field("Preset Name:", "name")
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
                    "Please enter a preset name.", parent=dlg)
                return

            if lib not in self._toner_presets:
                self._toner_presets[lib] = {}
            if name in self._toner_presets[lib]:
                messagebox.showwarning("Duplicate",
                    f"'{name}' already exists in '{lib}'.", parent=dlg)
                return

            self._toner_presets[lib][name] = {
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
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
            self._toner_active_library = lib
            self._toner_active_preset = name
            self._toner_active_session = None  # Clear old session
            self._toner_update_preset_label()

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
            f"Preset now has {len(total_notes)} unique notes total.")

    def _toner_export_presets(self):
        """Export selected tone profiles to a JSON file."""
        from tkinter import filedialog

        # Build list of all profiles
        all_presets = []
        for lib_name, lib_presets in self._toner_presets.items():
            if not isinstance(lib_presets, dict):
                continue
            for preset_name, preset_data in lib_presets.items():
                all_presets.append((lib_name, preset_name, preset_data))

        if not all_presets:
            messagebox.showinfo("Nothing to Export", "No tone presets to export.")
            return

        # Selection dialog
        dlg = tk.Toplevel(self.root)
        dlg.title("Export Tone Presets")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        bg = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"
        fg = "black"
        frame = tk.Frame(dlg, bg=bg, padx=20, pady=15)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Select presets to export:", bg=bg, fg=fg,
                 font=("Helvetica", 10)).pack(pady=(0, 5))

        # Checkboxes
        list_frame = tk.Frame(frame, bg=bg)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        check_vars = []
        for lib_name, preset_name, preset_data in all_presets:
            sessions = preset_data.get('sessions', [])
            caps = sum(len(s.get('captures', [])) for s in sessions)
            var = tk.BooleanVar(value=True)
            tk.Checkbutton(list_frame,
                text=f"[{lib_name}] {preset_name} ({caps} captures)",
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
            for (lib, name, data), var in zip(all_presets, check_vars):
                if var.get():
                    if lib not in export:
                        export[lib] = {}
                    export[lib][name] = data
                    count += 1

            if not export:
                messagebox.showinfo("Nothing Selected",
                    "Select at least one preset to export.", parent=dlg)
                return

            dlg.destroy()

            filepath = filedialog.asksaveasfilename(
                title="Export Tone Presets",
                defaultextension=".json",
                filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
                initialfile="toner_data_export.json"
            )
            if not filepath:
                return

            try:
                import json
                with open(filepath, 'w') as f:
                    json.dump(export, f, indent=2)
                messagebox.showinfo("Export Successful",
                    f"Exported {count} presets to:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Could not export:\n{e}")

        btn_frame = tk.Frame(frame, bg=bg)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="Export", command=do_export).pack(
            side="left", padx=(0, 5))
        tk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(
            side="left")

    def _toner_import_presets(self):
        """Import tone profiles from a JSON file, merging into existing libraries."""
        from tkinter import filedialog
        import json

        filepath = filedialog.askopenfilename(
            title="Import Tone Presets",
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
            messagebox.showerror("Invalid Format", "File is not a valid tone presets export.")
            return

        # Check if flat (old format) or nested (library format)
        is_flat = any(isinstance(v, dict) and 'sessions' in v
                      for v in imported.values())

        if is_flat:
            # Wrap in a library
            imported = {"Imported": imported}

        # Merge
        count = 0
        for lib_name, lib_presets in imported.items():
            if not isinstance(lib_presets, dict):
                continue
            if lib_name not in self._toner_presets:
                self._toner_presets[lib_name] = {}
            for preset_name, preset_data in lib_presets.items():
                if preset_name not in self._toner_presets[lib_name]:
                    self._toner_presets[lib_name][preset_name] = preset_data
                    count += 1
                else:
                    # Merge sessions into existing preset
                    existing = self._toner_presets[lib_name][preset_name]
                    existing_dates = {s.get('date') for s in existing.get('sessions', [])}
                    for session in preset_data.get('sessions', []):
                        if session.get('date') not in existing_dates:
                            existing.setdefault('sessions', []).append(session)
                            count += 1

        save_tone_presets(self._toner_presets, TONER_DATA_FILE)
        messagebox.showinfo("Import Complete",
            f"Imported {count} new presets/sessions.")

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
        if self._toner_active_session and self._toner_active_preset:
            save_tone_presets(self._toner_presets, TONER_DATA_FILE)
