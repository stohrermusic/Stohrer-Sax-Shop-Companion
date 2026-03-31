"""
Tone analyzer engine for Stohrer Sax Shop Companion.

Handles audio capture, FFT, fundamental pitch detection (peak-picking with harmonic series verification),
harmonic extraction, and tone descriptor computation. Pure math/audio —
no tkinter dependency.

Saxophone-specific: analyzes both even and odd harmonics (conical bore),
up to the 20th harmonic, in the range relevant to saxophone (Bb2–F#6
fundamentals, harmonics up to ~8kHz).

Requires: numpy, sounddevice (imported with try/except for graceful fallback)
"""

import math
import time
import json
import os
from collections import deque, namedtuple

try:
    import numpy as np
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except (ImportError, OSError):
    AUDIO_AVAILABLE = False
    np = None
    sd = None

from audio_utils import AudioRingBuffer  # noqa: E402 — shared with tuner_engine


# ============================================
# CONSTANTS
# ============================================

SAMPLE_RATE = 44100
BUFFER_SECONDS = 0.4          # 400ms for better low-frequency resolution
BUFFER_SIZE = int(SAMPLE_RATE * BUFFER_SECONDS)
FFT_SIZE = 16384              # ~370ms, ~2.69 Hz bin resolution

MAX_HARMONICS = 20            # Analyze up to 20th harmonic
SPECTRUM_MAX_HZ = 8000        # Display range for spectrum

PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

# Minimum fundamental frequency. Bari sax written Bb3 = concert Db2 (~69 Hz),
# so 65 Hz covers all standard saxophone types including baritone.
MIN_FUNDAMENTAL_HZ = 65.0
# Maximum fundamental frequency (altissimo range ~ 1500 Hz)
MAX_FUNDAMENTAL_HZ = 2000.0

# --- Fundamental detection tuning parameters ---
# Noise floor multiplier for sub-harmonic peak threshold (low divisors ≤3)
SUBHARM_NOISE_MULT_STRICT = 3.0
# Noise floor multiplier for sub-harmonic peak threshold (high divisors 4-5)
SUBHARM_NOISE_MULT_RELAXED = 2.0
# Prominence ratio for sub-harmonic peak vs neighbors (low divisors ≤3)
SUBHARM_PROMINENCE_STRICT = 1.5
# Prominence ratio for sub-harmonic peak vs neighbors (high divisors 4-5)
SUBHARM_PROMINENCE_RELAXED = 1.3
# Noise floor multiplier for harmonic series verification peaks
HARMONIC_VERIFY_NOISE_MULT = 3.0
# Required harmonic matches for low divisors (2-3)
HARMONIC_MATCHES_LOW = 2
# Required harmonic matches for high divisors (4-5)
HARMONIC_MATCHES_HIGH = 3
# Octave jump hysteresis: new peak must be this much stronger to override previous
OCTAVE_HYSTERESIS_FACTOR = 1.5

# --- Signal level and sensitivity ---
# RMS display scaling factor (maps raw RMS to 0-1 signal level indicator)
RMS_DISPLAY_SCALE = 10.0
# Base minimum RMS threshold before sensitivity scaling
RMS_MIN_BASE = 0.002
# Sensitivity range factor: maps sensitivity 0-100 to threshold scale 1.0-5.0
TONER_SENSITIVITY_FACTOR = 0.04

# --- Harmonic extraction ---
# Noise floor cutoff: harmonics weaker than this (dB re: fundamental) are discarded
HARMONIC_NOISE_FLOOR_DB = -60.0
# Cents deviation clamp for harmonic position
HARMONIC_CENTS_CLAMP = 100.0

# --- Descriptor scaling constants ---
# Richness significance threshold (dB re: fundamental) for "significant" harmonics
RICHNESS_SIG_THRESHOLD_DB = -35.0
# Richness: raw spectral flatness*coverage range mapped to gauge
RICHNESS_RAW_MIN = 0.50      # Below this = gauge 0%
RICHNESS_RAW_RANGE = 0.45    # 0.95 - 0.50
# Multi-frame averaging: smooth harmonic data across N frames before computing
# descriptors.  Reduces per-frame noise (~16% stdev within same note) without
# adding latency — consecutive FFT frames overlap heavily (~90% shared audio).
DESCRIPTOR_AVG_FRAMES = 3
# Warmth: H2 (octave harmonic) strength relative to fundamental.
# Strong H2 = round, warm quality. BA tenor reads -1.0 dB (very warm),
# Selmer Supreme reads -10.9 dB (thin). Maps -12 to 0 dB → 0-100%.
WARMTH_DB_FLOOR = -12.0      # H2 at or below this = gauge 0%
WARMTH_DB_RANGE = 12.0       # -12 to 0 dB maps to 0-100%

# --- Stream health ---
# Consecutive stale audio reads before triggering stream restart (~1s at 30fps)
TONER_STALE_RESTART_THRESHOLD = 30

# --- Spectral quality check ---
# Fraction of frames with weak low-freq energy to flag poor mic quality
SPECTRAL_QUALITY_THRESHOLD = 0.6
# Low-freq energy must be below this fraction of mid-freq energy to count as weak
LOW_FREQ_WEAKNESS_RATIO = 0.05

# --- Recording quality (harmonic rolloff) ---
# Average dB drop per harmonic from H2 to H12.  Good close-mic setups read
# 1.0–2.0 dB/harmonic; a laptop mic across the room reads 3.0+.
ROLLOFF_WARN_THRESHOLD = 2.5   # dB/harmonic — warn above this
ROLLOFF_MIN_CAPTURES = 5       # need this many before checking

# --- Delta gauge scaling ---
# Full gauge deflection ranges for comparison-only dials (in dB).
SPECTRAL_TILT_RANGE = 15.0    # ±15 dB = full deflection for H7-H12 avg delta
MID_HARMONIC_RANGE = 15.0     # ±15 dB = full deflection for H3-H6 avg delta

# Calibration capture: written chromatic scale Bb3 to F6
# These are WRITTEN pitches — the transposition to concert pitch
# happens using SAX_TRANSPOSITIONS when computing expected frequencies.
CALIBRATION_NOTES = [
    'A#3', 'B3',
    'C4', 'C#4', 'D4', 'D#4', 'E4', 'F4', 'F#4', 'G4', 'G#4', 'A4', 'A#4', 'B4',
    'C5', 'C#5', 'D5', 'D#5', 'E5', 'F5', 'F#5', 'G5', 'G#5', 'A5', 'A#5', 'B5',
    'C6', 'C#6', 'D6', 'D#6', 'E6', 'F6',
]
CALIBRATION_DURATION_S = 5.0  # Seconds per note


# ============================================
# ANALYSIS RESULT
# ============================================

HarmonicInfo = namedtuple('HarmonicInfo', [
    'harmonic_number',   # 1=fundamental, 2=2nd harmonic, etc.
    'expected_freq',     # Hz, ideal position
    'actual_freq',       # Hz, measured (parabolic interp)
    'magnitude_db',      # dB relative to fundamental
    'cents_deviation',   # cents off from ideal harmonic ratio
])


class TonerResult:
    """Result of one tone analysis frame."""
    __slots__ = [
        'fundamental_freq',      # Hz (0 if no pitch detected)
        'fundamental_note',      # str like "C4", "Bb3"
        'fundamental_cents',     # cents offset from nearest semitone
        'harmonics',             # list of HarmonicInfo
        'spectrum_db',           # numpy array, dB magnitude for spectrum display
        'spectrum_freqs',        # numpy array, frequency axis
        'harmonic_bars',         # list of (freq, magnitude_db) for bar-only mode
        'descriptors',           # dict: richness, warmth
        'signal_level',          # 0.0-1.0, RMS input level
        'spectral_centroid',     # Hz, amplitude-weighted center of harmonic series
    ]

    def __init__(self):
        self.fundamental_freq = 0.0
        self.fundamental_note = ""
        self.fundamental_cents = 0.0
        self.harmonics = []
        self.spectrum_db = None
        self.spectrum_freqs = None
        self.harmonic_bars = []
        self.descriptors = {
            'richness': 0.0,
            'warmth': 0.0,
            'low_harmonic_data': False,
        }
        self.signal_level = 0.0
        self.spectral_centroid = 0.0


# ============================================
# TONER ENGINE
# ============================================

# Keywords that identify built-in/laptop mics (case-insensitive)
_BUILTIN_MIC_KEYWORDS = [
    'built-in', 'internal', 'microphone array', 'realtek',
    'integrated', 'laptop', 'webcam',
]


def check_mic_quality():
    """Check if the default input device looks like a built-in mic.

    Returns:
        (is_builtin, device_name) — is_builtin is True if the default
        input device name matches known built-in mic patterns.
    """
    if not AUDIO_AVAILABLE:
        return False, "unknown"
    try:
        default_input = sd.default.device[0]
        if default_input is None or default_input < 0:
            return False, "unknown"
        info = sd.query_devices(default_input)
        name = info.get('name', '')
        name_lower = name.lower()
        is_builtin = any(kw in name_lower for kw in _BUILTIN_MIC_KEYWORDS)
        return is_builtin, name
    except Exception:
        return False, "unknown"


class TonerEngine:
    """Audio capture and tone analysis engine."""

    def __init__(self):
        self._stream = None
        self._ring_buffer = None
        self._reference_pitch = 440.0
        self._sensitivity = 50
        self._running = False
        self._window = None
        self._last_fundamental = 0.0  # For temporal smoothing
        self._harmonic_history = deque(maxlen=DESCRIPTOR_AVG_FRAMES)
        self._last_device = None  # For auto-restart
        self._stale_count = 0  # Consecutive stale reads
        self.last_error = None     # Set when stream restart fails
        self._break_freq = DEFAULT_BREAK_FREQ  # Benade spectral break frequency
        self._mic_quality_warned = False  # Only warn once per session
        self._spectral_check_frames = 0  # Count frames for spectral quality check
        self._low_energy_frames = 0  # Frames where low freqs are suspiciously weak

    def check_spectral_quality(self):
        """Check if the mic shows poor low-frequency response.

        Returns True if the mic appears to have low-freq rolloff
        (likely a built-in/laptop mic). Only meaningful after ~30
        frames of real audio have been analyzed.
        """
        if self._spectral_check_frames < 20:
            return False  # Not enough data yet
        # If >60% of frames show weak low-frequency energy, flag it
        return self._low_energy_frames > self._spectral_check_frames * SPECTRAL_QUALITY_THRESHOLD

    def set_break_frequency(self, hz):
        """Set the Benade break frequency for sax type."""
        self._break_freq = float(hz)

    def set_sax_type(self, sax_type):
        """Set break frequency from saxophone type name."""
        self._break_freq = BREAK_FREQUENCIES.get(sax_type, DEFAULT_BREAK_FREQ)

    @property
    def reference_pitch(self):
        return self._reference_pitch

    def set_reference_pitch(self, hz):
        self._reference_pitch = float(hz)

    def set_sensitivity(self, value):
        self._sensitivity = max(0, min(100, int(value)))

    def start(self, device=None):
        """Start audio capture. Returns (success, error_message)."""
        if not AUDIO_AVAILABLE:
            return False, "Audio libraries not available.\nInstall numpy and sounddevice:\n  pip install numpy sounddevice"

        if self._running:
            self.stop()

        self._ring_buffer = AudioRingBuffer(BUFFER_SIZE)
        self._window = np.hanning(FFT_SIZE).astype(np.float32)
        self._last_device = device
        self._stale_count = 0
        self.last_error = None

        try:
            self._stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=1,
                dtype='float32',
                blocksize=1024,
                device=device,
                callback=self._audio_callback,
            )
            self._stream.start()
            self._running = True
            return True, None
        except Exception as e:
            self._running = False
            return False, str(e)

    def stop(self):
        self._running = False
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        self._ring_buffer = None

    @property
    def is_running(self):
        return self._running

    def _audio_callback(self, indata, frames, time_info, status):
        if self._ring_buffer is not None:
            self._ring_buffer.write(indata[:, 0])

    def analyze(self):
        """Analyze current audio buffer. Returns TonerResult.

        Monitors stream health and auto-restarts if the audio callback
        appears to have died (no new data for >1 second).
        """
        result = TonerResult()

        buf = self._ring_buffer
        if not self._running or buf is None:
            return result

        # Check stream health: if buffer is stale, the callback may have died
        if buf.is_stale():
            self._stale_count += 1
            if self._stale_count > TONER_STALE_RESTART_THRESHOLD:
                self._stale_count = 0
                self._restart_stream()
                return result
        else:
            self._stale_count = 0

        audio = buf.read()
        if audio is None:
            return result

        result = self.analyze_buffer(audio)

        # --- Multi-frame averaging for descriptor stability ---
        # Smooth harmonic data across recent frames before recomputing
        # descriptors.  result.harmonics stays per-frame for spectrum
        # display and captures; descriptors use the averaged data.
        # analyze_buffer() remains stateless for tests.
        if result.harmonics and result.fundamental_freq > 0:
            # Clear history on note change (>1 semitone)
            if (self._last_fundamental > 0 and
                    abs(1200.0 * math.log2(
                        result.fundamental_freq / self._last_fundamental)) > 100):
                self._harmonic_history.clear()
            frame_data = {}
            for h in result.harmonics:
                frame_data[h.harmonic_number] = (
                    h.magnitude_db, h.cents_deviation,
                    h.expected_freq, h.actual_freq)
            self._harmonic_history.append(frame_data)
            if len(self._harmonic_history) > 1:
                avg_harmonics = self._averaged_harmonics(result.fundamental_freq)
                result.descriptors = self._compute_descriptors(
                    avg_harmonics, result.fundamental_freq)

        return result

    def _restart_stream(self):
        """Restart the audio stream (recover from dead callback)."""
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass

        self._ring_buffer = AudioRingBuffer(BUFFER_SIZE)
        self._window = np.hanning(FFT_SIZE).astype(np.float32)

        try:
            self._stream = sd.InputStream(
                samplerate=SAMPLE_RATE,
                channels=1,
                dtype='float32',
                blocksize=1024,
                device=self._last_device,
                callback=self._audio_callback,
            )
            self._stream.start()
            self.last_error = None
        except Exception as e:
            self._running = False
            self.last_error = f"Audio stream lost: {e}"

    def analyze_buffer(self, audio):
        """Analyze a raw audio buffer. Used by analyze() and tests."""
        result = TonerResult()

        if len(audio) < FFT_SIZE:
            return result

        # RMS signal level
        rms = float(np.sqrt(np.mean(audio[-FFT_SIZE:] ** 2)))
        result.signal_level = min(1.0, rms * RMS_DISPLAY_SCALE)

        # Sensitivity threshold
        sensitivity_scale = 1.0 + (100 - self._sensitivity) * TONER_SENSITIVITY_FACTOR
        min_rms = RMS_MIN_BASE * sensitivity_scale
        if rms < min_rms:
            return result

        frame = audio[-FFT_SIZE:]

        if self._window is None or len(self._window) != FFT_SIZE:
            self._window = np.hanning(FFT_SIZE).astype(np.float32)

        windowed = frame * self._window
        spectrum = np.fft.rfft(windowed)
        mags = np.abs(spectrum)

        bin_freq = SAMPLE_RATE / FFT_SIZE  # ~2.69 Hz

        # --- Build spectrum for display ---
        max_bin = min(len(mags), int(SPECTRUM_MAX_HZ / bin_freq) + 1)
        spec_mags = mags[:max_bin]
        ref_mag = np.max(spec_mags) if np.max(spec_mags) > 0 else 1.0
        spectrum_db = 20.0 * np.log10(spec_mags / ref_mag + 1e-10)
        spectrum_db = np.clip(spectrum_db, -80.0, 0.0)
        result.spectrum_db = spectrum_db
        result.spectrum_freqs = np.arange(max_bin) * bin_freq

        # --- Detect fundamental ---
        f0 = self._detect_fundamental(mags, bin_freq)
        if f0 <= 0:
            return result

        result.fundamental_freq = f0

        # --- Spectral quality check (first 30 frames with audio) ---
        # If low frequencies (<300 Hz) are consistently much weaker than
        # mid frequencies (500-2000 Hz), the mic probably has low-freq rolloff.
        if not self._mic_quality_warned and self._spectral_check_frames < 30:
            self._spectral_check_frames += 1
            low_bin = int(300 / bin_freq)
            mid_lo = int(500 / bin_freq)
            mid_hi = int(2000 / bin_freq)
            if low_bin > 0 and mid_hi < len(mags):
                low_energy = float(np.mean(mags[1:low_bin]))
                mid_energy = float(np.mean(mags[mid_lo:mid_hi]))
                if mid_energy > 0 and low_energy < mid_energy * LOW_FREQ_WEAKNESS_RATIO:
                    self._low_energy_frames += 1

        # Map to note name and cents
        note_name, cents = self._freq_to_note(f0)
        result.fundamental_note = note_name
        result.fundamental_cents = cents

        # --- Extract harmonics ---
        harmonics = self._extract_harmonics(mags, bin_freq, f0)
        result.harmonics = harmonics

        # Build bar data
        result.harmonic_bars = []
        for h in harmonics:
            result.harmonic_bars.append((h.expected_freq, h.magnitude_db))

        # --- Spectral centroid of harmonic series ---
        # Amplitude-weighted center frequency: sum(f * A) / sum(A)
        if harmonics:
            linear_amps = [10.0 ** (h.magnitude_db / 20.0) for h in harmonics]
            freq_sum = sum(h.actual_freq * a for h, a in zip(harmonics, linear_amps))
            amp_sum = sum(linear_amps)
            result.spectral_centroid = freq_sum / amp_sum if amp_sum > 0 else 0.0

        # --- Compute descriptors ---
        result.descriptors = self._compute_descriptors(harmonics, f0)

        return result

    def _detect_fundamental(self, mags, bin_freq):
        """Detect fundamental frequency using peak-picking with harmonic
        series verification.

        Strategy:
        1. Find the strongest spectral peak in the valid range
        2. Check if sub-harmonics (f/2, f/3, f/4, f/5) could be the real
           fundamental by verifying they have their OWN harmonic series
        3. A sub-harmonic is only accepted if multiple of its harmonics
           also have peaks — not just the sub-harmonic alone
        4. Pick the lowest verified sub-harmonic (deepest fundamental)
        5. Apply temporal hysteresis for stability
        """
        min_bin = max(1, int(MIN_FUNDAMENTAL_HZ / bin_freq))
        max_bin = min(len(mags) - 2, int(MAX_FUNDAMENTAL_HZ / bin_freq))

        if max_bin <= min_bin:
            return 0.0

        # Find the strongest peak in the fundamental range
        search_range = mags[min_bin:max_bin + 1]
        strongest_bin = min_bin + int(np.argmax(search_range))
        strongest_mag = float(mags[strongest_bin])

        if strongest_mag <= 0:
            return 0.0

        # Noise floor for peak detection
        noise_floor = float(np.median(mags[min_bin:max_bin])) if max_bin > min_bin else 0.0

        # Check sub-harmonics: could the strongest peak be harmonic 2-5
        # of a lower fundamental? This catches sax low register where
        # harmonics 3-5 are often stronger than the fundamental.
        # Check all divisors and pick the lowest verified sub-harmonic.
        best_candidate = None
        candidate_bin = strongest_bin
        for divisor in [2, 3, 4, 5]:
            sub_bin = int(round(strongest_bin / divisor))
            if sub_bin < min_bin:
                continue

            # Find peak near sub-harmonic position (wider window for
            # higher divisors where rounding error grows)
            spread = 2 + (divisor // 3)  # 2 for /2,/3; 3 for /4,/5
            lo = max(1, sub_bin - spread)
            hi = min(len(mags) - 2, sub_bin + spread)
            local_peak = lo + int(np.argmax(mags[lo:hi + 1]))
            local_mag = float(mags[local_peak])

            # Sub-harmonic must be above noise floor. For higher divisors
            # the fundamental can be very weak, so relax slightly.
            noise_thresh = noise_floor * (SUBHARM_NOISE_MULT_STRICT if divisor <= 3 else SUBHARM_NOISE_MULT_RELAXED)
            if local_mag < noise_thresh:
                continue
            left_mag = float(mags[max(0, local_peak - 4)])
            right_mag = float(mags[min(len(mags) - 1, local_peak + 4)])
            prominence = SUBHARM_PROMINENCE_STRICT if divisor <= 3 else SUBHARM_PROMINENCE_RELAXED
            if not (local_mag > left_mag * prominence
                    and local_mag > right_mag * prominence):
                continue

            # Verify harmonic series: check that multiples of the
            # sub-harmonic also have peaks. For higher divisors, check
            # more multiples and require more matches.
            harmonics_found = 0
            check_mults = [2, 3, 4] if divisor <= 3 else [2, 3, 4, 5, 6]
            for mult in check_mults:
                h_bin = int(round(local_peak * mult))
                if h_bin >= len(mags) - 1:
                    continue
                h_lo = max(1, h_bin - 3)
                h_hi = min(len(mags) - 2, h_bin + 3)
                h_peak_mag = float(np.max(mags[h_lo:h_hi + 1]))
                if h_peak_mag > noise_floor * HARMONIC_VERIFY_NOISE_MULT:
                    harmonics_found += 1

            required = HARMONIC_MATCHES_LOW if divisor <= 3 else HARMONIC_MATCHES_HIGH
            if harmonics_found >= required:
                # Prefer the lowest sub-harmonic (deepest fundamental)
                if best_candidate is None or local_peak < best_candidate:
                    best_candidate = local_peak

        if best_candidate is not None:
            candidate_bin = best_candidate

        # Parabolic interpolation on the candidate
        if 0 < candidate_bin < len(mags) - 1:
            alpha = float(mags[candidate_bin - 1])
            beta = float(mags[candidate_bin])
            gamma = float(mags[candidate_bin + 1])
            denom = alpha - 2 * beta + gamma
            if abs(denom) > 1e-10 and beta > 0:
                p = 0.5 * (alpha - gamma) / denom
                new_freq = (candidate_bin + p) * bin_freq
            else:
                new_freq = candidate_bin * bin_freq
        else:
            new_freq = candidate_bin * bin_freq

        # Temporal hysteresis: if the previous detection was stable and the
        # new one is an octave jump (ratio near 2.0 or 0.5), require the
        # new frequency to be the stronger peak. This prevents flickering.
        prev = self._last_fundamental
        if prev > 0 and new_freq > 0:
            ratio = new_freq / prev
            # Check for octave-ish jumps (within 10% of 2x or 0.5x)
            if 0.45 < ratio < 0.55 or 1.8 < ratio < 2.2:
                # Is the previous frequency's bin still strong?
                prev_bin = int(round(prev / bin_freq))
                if 0 < prev_bin < len(mags) - 1:
                    prev_mag = float(mags[prev_bin])
                    new_mag = float(mags[candidate_bin])
                    # Stick with previous unless new is clearly stronger
                    if new_mag < prev_mag * OCTAVE_HYSTERESIS_FACTOR:
                        new_freq = prev

        self._last_fundamental = new_freq
        return new_freq

    def _freq_to_note(self, freq):
        """Convert frequency to note name and cents deviation."""
        if freq <= 0:
            return "", 0.0

        # Semitones from A4
        semitones_from_a4 = 12.0 * math.log2(freq / self._reference_pitch)
        nearest_semitone = round(semitones_from_a4)
        cents = (semitones_from_a4 - nearest_semitone) * 100.0

        # Map to note name + octave
        midi_note = int(nearest_semitone + 69)  # A4 = MIDI 69
        pc = midi_note % 12  # 0=C, 1=C#, ...
        octave = midi_note // 12 - 1
        note_name = f"{PITCH_CLASSES[pc]}{octave}"

        return note_name, cents

    def _extract_harmonics(self, mags, bin_freq, f0):
        """Extract harmonic information relative to fundamental."""
        harmonics = []
        fundamental_mag = 0.0

        for n in range(1, MAX_HARMONICS + 1):
            expected_freq = f0 * n
            if expected_freq > SPECTRUM_MAX_HZ:
                break

            expected_bin = expected_freq / bin_freq
            center_bin = int(round(expected_bin))

            if center_bin < 1 or center_bin >= len(mags) - 1:
                continue

            # Search +/- 3 bins for the actual peak
            search_lo = max(1, center_bin - 3)
            search_hi = min(len(mags) - 2, center_bin + 3)
            local_peak_bin = search_lo + int(np.argmax(mags[search_lo:search_hi + 1]))
            peak_mag = float(mags[local_peak_bin])

            # Parabolic interpolation for frequency and amplitude
            if local_peak_bin > 0 and local_peak_bin < len(mags) - 1:
                alpha = float(mags[local_peak_bin - 1])
                beta = float(mags[local_peak_bin])
                gamma = float(mags[local_peak_bin + 1])
                denom = alpha - 2 * beta + gamma
                if abs(denom) > 1e-10 and beta > 0:
                    p = 0.5 * (alpha - gamma) / denom
                    actual_freq = (local_peak_bin + p) * bin_freq
                    # Amplitude correction: interpolated peak magnitude
                    corrected = beta - 0.25 * (alpha - gamma) * p
                    if corrected > 0:
                        peak_mag = corrected
                else:
                    actual_freq = local_peak_bin * bin_freq
            else:
                actual_freq = local_peak_bin * bin_freq

            # dB relative to fundamental (set after first pass)
            if n == 1:
                fundamental_mag = peak_mag
                mag_db = 0.0
            else:
                if fundamental_mag > 0:
                    mag_db = 20.0 * math.log10(peak_mag / fundamental_mag + 1e-10)
                else:
                    mag_db = -80.0

            # Skip harmonics buried in the noise floor
            if n > 1 and mag_db < HARMONIC_NOISE_FLOOR_DB:
                continue

            # Cents deviation from ideal harmonic position
            if actual_freq > 0 and expected_freq > 0:
                cents_dev = 1200.0 * math.log2(actual_freq / expected_freq)
                cents_dev = max(-HARMONIC_CENTS_CLAMP, min(HARMONIC_CENTS_CLAMP, cents_dev))
            else:
                cents_dev = 0.0

            harmonics.append(HarmonicInfo(
                harmonic_number=n,
                expected_freq=expected_freq,
                actual_freq=actual_freq,
                magnitude_db=mag_db,
                cents_deviation=cents_dev,
            ))

        return harmonics

    def _averaged_harmonics(self, f0):
        """Average harmonic data across recent frames for descriptor stability.

        Returns a list of HarmonicInfo with averaged dB and cents values.
        Uses current-frame frequencies (expected/actual) since those don't
        benefit from averaging — only amplitude and deviation do.
        """
        if not self._harmonic_history:
            return []

        # Collect all harmonic numbers seen across frames
        all_nums = set()
        for frame in self._harmonic_history:
            all_nums.update(frame.keys())

        averaged = []
        latest = self._harmonic_history[-1]
        for n in sorted(all_nums):
            db_vals = []
            cents_vals = []
            for frame in self._harmonic_history:
                if n in frame:
                    db_vals.append(frame[n][0])
                    cents_vals.append(frame[n][1])
            if not db_vals:
                continue
            avg_db = sum(db_vals) / len(db_vals)
            avg_cents = sum(cents_vals) / len(cents_vals)
            # Use latest frame's frequency data
            if n in latest:
                exp_freq, act_freq = latest[n][2], latest[n][3]
            else:
                exp_freq, act_freq = f0 * n, f0 * n
            averaged.append(HarmonicInfo(n, exp_freq, act_freq,
                                         avg_db, avg_cents))
        return averaged

    def _compute_descriptors(self, harmonics, f0=440.0):
        """Compute tone quality descriptors from harmonic data.

        Richness: spectral flatness (geometric/arithmetic mean ratio).
        Warmth: H2 (octave harmonic) strength relative to fundamental.
        """
        if not harmonics:
            return {'richness': 0.0, 'warmth': 0.0, 'core_tone': 0.0,
                    'even_odd': 0.0, 'rolloff_shape': 0.0,
                    'low_harmonic_data': False}

        # --- Richness: spectral flatness of significant harmonics ---
        # Rescaled for sax-vs-sax comparison: raw flatness*coverage ranges
        # from ~0.50 (thin/uneven harmonics) to ~0.95 (very even spread).
        # The gauge maps 0.50-0.95 to 0-100%, giving useful range between
        # saxophones rather than pegging near 100% for all of them.
        sig_threshold_db = RICHNESS_SIG_THRESHOLD_DB
        upper_harmonics = [h for h in harmonics if h.harmonic_number > 1]
        significant = [h for h in upper_harmonics
                       if h.magnitude_db > sig_threshold_db]
        max_possible = max(1, len(upper_harmonics))

        if len(significant) >= 2:
            linear_mags = [10.0 ** (h.magnitude_db / 20.0) for h in significant]
            log_sum = sum(math.log(m + 1e-10) for m in linear_mags)
            geo_mean = math.exp(log_sum / len(linear_mags))
            arith_mean = sum(linear_mags) / len(linear_mags)
            flatness = geo_mean / arith_mean if arith_mean > 0 else 0.0
            coverage = len(significant) / max_possible
            raw_richness = flatness * coverage
            # Rescale raw range to 0.0-1.0 (sax-specific range)
            richness = max(0.0, min(1.0, (raw_richness - RICHNESS_RAW_MIN) / RICHNESS_RAW_RANGE))
        elif len(significant) == 1:
            richness = 0.05
        else:
            richness = 0.0

        # --- Low harmonic data flag ---
        h_by_num = {h.harmonic_number: h.magnitude_db for h in harmonics}
        harmonic_available = sum(1 for h in range(2, 7) if h in h_by_num)
        low_harmonic_data = harmonic_available < 3

        # --- Warmth: H2 (octave harmonic) strength ---
        # H2 is the octave harmonic. Strong H2 = round, warm quality.
        # BA tenor H2 = -1.0 dB (warmest), Selmer Supreme = -10.9 dB (thinnest).
        h2_db = h_by_num.get(2, -30.0)
        warmth = max(0.0, min(1.0,
            (h2_db - WARMTH_DB_FLOOR) / WARMTH_DB_RANGE))

        # --- Core Tone: Tristimulus T2 (H2-H4 energy weight) ---
        # Measures the proportion of total energy in the body harmonics.
        # H2-H4 are below the tone hole cutoff frequency, so they're shaped
        # primarily by the bore — the most player-independent metric we have.
        # T2 = (|H2| + |H3| + |H4|) / sum(|all|)
        all_lin = {h.harmonic_number: 10.0 ** (h.magnitude_db / 20.0) for h in harmonics}
        total_lin = sum(all_lin.values())
        if total_lin > 0 and all(n in all_lin for n in (2, 3, 4)):
            core_tone = (all_lin[2] + all_lin[3] + all_lin[4]) / total_lin
        else:
            core_tone = 0.0

        # --- Even/Odd Ratio ---
        # Ratio of even harmonic energy to odd harmonic energy (H2+ only).
        # Even harmonics (H2,H4,H6...) produce round/warm quality;
        # odd harmonics (H3,H5,H7...) produce edgier/hollower quality.
        # Conical bore sax produces both, but the ratio varies by horn.
        # Scaled: 0.8 → 0%, 1.8 → 100%  (observed range from 47 profiles).
        even_sum = sum(all_lin.get(n, 0) for n in range(2, 22, 2))
        odd_sum = sum(all_lin.get(n, 0) for n in range(3, 22, 2))
        if odd_sum > 0:
            raw_eo = even_sum / odd_sum
            even_odd = max(0.0, min(1.0, (raw_eo - 0.8) / 1.0))
        else:
            even_odd = 0.0

        # --- Rolloff Shape (nonlinearity) ---
        # How smoothly harmonics roll off vs having bumps/peaks.
        # Linear regression of H2-Hmax dB values, then average residual.
        # Low = smooth rolloff. High = spectral peaks that stick out.
        # May correspond to "presence" or "projection."
        # Per-note raw range: 1.0-5.0.  Maps 1.0-5.0 → 0-100%.
        upper_db = [(h.harmonic_number, h.magnitude_db)
                    for h in harmonics if h.harmonic_number >= 2]
        if len(upper_db) >= 4:
            x_vals = [n for n, _ in upper_db]
            y_vals = [db for _, db in upper_db]
            x_mean = sum(x_vals) / len(x_vals)
            y_mean = sum(y_vals) / len(y_vals)
            denom = sum((x - x_mean) ** 2 for x in x_vals)
            if denom > 0:
                slope = sum((x - x_mean) * (y - y_mean)
                           for x, y in zip(x_vals, y_vals)) / denom
                predicted = [y_mean + slope * (x - x_mean) for x in x_vals]
                residuals = [abs(y - p) for y, p in zip(y_vals, predicted)]
                raw_nonlin = sum(residuals) / len(residuals)
                rolloff_shape = max(0.0, min(1.0, (raw_nonlin - 1.0) / 4.0))
            else:
                rolloff_shape = 0.0
        else:
            rolloff_shape = 0.0

        return {
            'richness': richness,
            'warmth': warmth,
            'core_tone': core_tone,
            'even_odd': even_odd,
            'rolloff_shape': rolloff_shape,
            'low_harmonic_data': low_harmonic_data,
        }


# ============================================
# TONE PROFILES — capture, storage, comparison
# ============================================

# Minimum unique notes to consider a profile "complete"
MIN_PROFILE_NOTES = 8

# Attack transient skip — applied to ALL capture modes.
# Saxophone attacks are 40-120ms depending on articulation. The attack
# contains non-harmonic broadband energy that doesn't represent the
# horn's sustained tone character. Research confirms the sustained
# portion is the stable harmonic fingerprint we want to measure.
# (Saldanha & Corso 1964, Caetano et al., PMC 6166322)
ATTACK_SKIP_FRAMES = 3    # ~100ms at 30fps — skip these after note onset

# Capture timing
CAPTURE_DELAY_S = 1.0     # Seconds settle for calibration mode

# Free mode: shorter stability requirement, no fixed recording period
FREE_STABLE_FRAMES = 15   # ~0.5s at 30fps to consider a note "stable"
FREE_MIN_FRAMES = 10      # Minimum frames to keep a micro-capture

# File import: analysis hop size
FILE_HOP_SAMPLES = 1024   # Advance this many samples between analyses

SAX_TYPES = ["Bass", "Baritone", "Tenor", "C Melody",
             "Alto", "F Mezzo", "Soprano", "Sopranino"]

# Benade spectral envelope break frequencies (Hz).
# Below f_b, harmonics rise ~3 dB/oct. Above, they fall ~18 dB/oct.
# The break frequency marks the acoustic boundary between "lower" and
# "upper" harmonics for spectral analysis.
# Tenor (618) and Alto (837) measured by Benade/Wolfe (JASA 1988, UNSW).
# Others estimated from bore length scaling.
BREAK_FREQUENCIES = {
    "Sopranino": 1550,
    "Soprano": 1300,
    "F Mezzo": 1000,
    "Alto": 837,
    "C Melody": 725,
    "Tenor": 618,
    "Baritone": 450,
    "Bass": 350,
}
DEFAULT_BREAK_FREQ = 750  # Fallback (between alto and tenor)

# Transposition: sax type → semitones UP from concert to written pitch.
# Must include octave displacement for correct note names.
# Soprano in Bb: concert Bb3 = written C4 → shift = +2
# Alto in Eb: concert Eb3 = written C4 → shift = +9
# Tenor in Bb: concert Bb2 = written C4 → shift = +14 (octave + major 2nd)
# Baritone in Eb: concert Eb2 = written C4 → shift = +21 (octave + major 6th)
SAX_TRANSPOSITIONS = {
    "Sopranino": -3,   # Eb (sounds minor 3rd ABOVE written)
    "Soprano": 2,      # Bb (sounds major 2nd below)
    "F Mezzo": 7,      # F (sounds perfect 5th below)
    "Alto": 9,         # Eb (sounds major 6th below)
    "C Melody": 0,     # C (concert pitch)
    "Tenor": 14,       # Bb (sounds major 9th below)
    "Baritone": 21,    # Eb (sounds major 13th below)
    "Bass": 26,        # Bb (sounds 2 octaves + major 2nd below)
}


def transpose_note(concert_note, sax_type):
    """Transpose a concert pitch note name to written pitch for a given sax type.

    Pure function — no UI dependency. Returns the original note if sax_type
    is unknown or has no transposition.
    """
    if not concert_note:
        return concert_note
    shift = SAX_TRANSPOSITIONS.get(sax_type, 0)
    if shift == 0:
        return concert_note
    return _shift_note(concert_note, shift)


def reverse_transpose_note(written_note, sax_type):
    """Transpose a written pitch note name back to concert pitch.

    Pure function — inverse of transpose_note().
    """
    if not written_note:
        return written_note
    shift = SAX_TRANSPOSITIONS.get(sax_type, 0)
    if shift == 0:
        return written_note
    return _shift_note(written_note, -shift)


def note_to_freq(note_name, reference_pitch=440.0):
    """Convert a note name like 'C#4' to its frequency in Hz."""
    if not note_name:
        return 0.0
    if '#' in note_name:
        pc_name = note_name[:-1]
    else:
        pc_name = note_name[:-1]
    try:
        octave = int(note_name[-1])
        pc_idx = PITCH_CLASSES.index(pc_name)
    except (ValueError, IndexError):
        return 0.0
    midi = (octave + 1) * 12 + pc_idx
    return reference_pitch * 2 ** ((midi - 69) / 12)


def _shift_note(note_name, semitones):
    """Shift a note name by the given number of semitones.

    Internal helper for transpose_note / reverse_transpose_note.
    """
    if not note_name or len(note_name) < 2 or not note_name[-1].isdigit():
        return note_name
    if '#' in note_name:
        pc_name = note_name[:-1]
    else:
        pc_name = note_name[:-1]
    try:
        octave = int(note_name[-1])
        pc_idx = PITCH_CLASSES.index(pc_name)
    except (ValueError, IndexError):
        return note_name
    new_pc = (pc_idx + semitones) % 12
    new_octave = octave + ((pc_idx + semitones) // 12)
    return f"{PITCH_CLASSES[new_pc]}{new_octave}"



def descriptors_from_harmonics(harmonics_db, f0, sax_type="Tenor",
                               harmonic_cents=None):
    """Compute descriptors from raw harmonic dB data using current formulas.

    This is the canonical way to interpret stored harmonic measurements.
    Captures store only the raw data (harmonics_db + fundamental_freq +
    harmonic_cents); descriptors are always computed on the fly so formula
    improvements apply retroactively to all historical data.

    Args:
        harmonics_db: list of dB values relative to fundamental (index 0 = H1 = 0dB)
        f0: fundamental frequency in Hz
        sax_type: saxophone type string (reserved for future use)
        harmonic_cents: optional list of cents deviations from ideal position

    Returns:
        dict with richness, warmth (0.0-1.0) and low_harmonic_data flag
    """
    if not harmonics_db or f0 <= 0:
        return {'richness': 0.0, 'warmth': 0.0, 'core_tone': 0.0,
                'even_odd': 0.0, 'rolloff_shape': 0.0,
                'low_harmonic_data': False}
    # Lightweight engine instance for _compute_descriptors
    engine = TonerEngine.__new__(TonerEngine)
    engine._break_freq = BREAK_FREQUENCIES.get(sax_type, DEFAULT_BREAK_FREQ)
    harmonics = []
    for i, db in enumerate(harmonics_db):
        cents = harmonic_cents[i] if harmonic_cents and i < len(harmonic_cents) else 0.0
        harmonics.append(HarmonicInfo(i + 1, f0 * (i + 1), f0 * (i + 1), db, cents))
    return engine._compute_descriptors(harmonics, f0=f0)


def average_captures(captures):
    """Average a list of per-frame capture dicts into one summary.

    Each capture dict must have 'harmonics_db' (list of dB values).
    May also have 'harmonic_cents' (list of cents deviations).
    Descriptors, if present, are ignored — they are always recomputed
    from harmonics_db by the caller.

    Returns dict with 'harmonics_db', 'fundamental_freq', and
    'harmonic_cents' (if any input had cents data).
    """
    if not captures:
        return None

    # Average harmonic dB values
    max_len = max(len(c.get('harmonics_db', [])) for c in captures)
    avg_db = [0.0] * max_len
    counts = [0] * max_len
    for c in captures:
        for i, db in enumerate(c.get('harmonics_db', [])):
            avg_db[i] += db
            counts[i] += 1
    for i in range(max_len):
        if counts[i] > 0:
            avg_db[i] /= counts[i]

    # Average harmonic cents deviations (if available)
    has_cents = any(c.get('harmonic_cents') for c in captures)
    avg_cents = []
    if has_cents:
        max_cents_len = max((len(c.get('harmonic_cents', [])) for c in captures), default=0)
        avg_cents = [0.0] * max_cents_len
        cents_counts = [0] * max_cents_len
        for c in captures:
            for i, cents in enumerate(c.get('harmonic_cents', [])):
                avg_cents[i] += cents
                cents_counts[i] += 1
        for i in range(max_cents_len):
            if cents_counts[i] > 0:
                avg_cents[i] /= cents_counts[i]

    # Average fundamental frequency if available
    freqs = [c.get('fundamental_freq', c.get('freq', 0)) for c in captures]
    freqs = [f for f in freqs if f > 0]
    avg_freq = sum(freqs) / len(freqs) if freqs else 0.0

    result = {
        'harmonics_db': avg_db,
        'fundamental_freq': avg_freq,
    }
    if avg_cents:
        result['harmonic_cents'] = avg_cents
    return result


def compute_rolloff_rate(harmonics_db):
    """Return the average dB drop per harmonic from H2 to the highest available.

    A low value (1.0–2.0) indicates a good close-mic setup.  A high value
    (3.0+) suggests the mic or room is suppressing upper harmonics.
    Returns None if fewer than 8 harmonics are present.
    """
    if not harmonics_db or len(harmonics_db) < 8:
        return None
    last_idx = min(11, len(harmonics_db) - 1)  # H12 = index 11
    span = last_idx - 1  # number of steps from H2 (index 1) to last
    if span <= 0:
        return None
    return abs(harmonics_db[last_idx] - harmonics_db[1]) / span


def compute_delta_descriptors(live_harmonics_db, baseline_harmonics_db,
                              live_descriptors, baseline_descriptors):
    """Compute delta descriptors between live and baseline readings.

    Returns dict with:
        'richness_delta': signed difference (-1 to +1)
        'warmth_delta': signed difference (-1 to +1)
        'spectral_tilt': upper harmonic shift in dB (avg H7-H12 delta),
                         normalized to -1..+1 using SPECTRAL_TILT_RANGE
        'mid_harmonic': mid harmonic shift in dB (avg H3-H6 delta),
                        normalized to -1..+1 using MID_HARMONIC_RANGE
    Returns None if insufficient data.
    """
    if not live_harmonics_db or not baseline_harmonics_db:
        return None
    if not live_descriptors or not baseline_descriptors:
        return None

    result = {
        'richness_delta': live_descriptors.get('richness', 0) -
                          baseline_descriptors.get('richness', 0),
        'warmth_delta': live_descriptors.get('warmth', 0) -
                        baseline_descriptors.get('warmth', 0),
    }

    # Spectral tilt: average dB delta of H7-H12 (indices 6-11)
    n = min(len(live_harmonics_db), len(baseline_harmonics_db))
    if n >= 12:
        upper_deltas = [live_harmonics_db[i] - baseline_harmonics_db[i]
                        for i in range(6, 12)]
        raw_tilt = sum(upper_deltas) / len(upper_deltas)
        result['spectral_tilt'] = max(-1.0, min(1.0,
            raw_tilt / SPECTRAL_TILT_RANGE))
    elif n >= 8:
        # Fewer harmonics available — use what we have
        upper_deltas = [live_harmonics_db[i] - baseline_harmonics_db[i]
                        for i in range(6, n)]
        raw_tilt = sum(upper_deltas) / len(upper_deltas)
        result['spectral_tilt'] = max(-1.0, min(1.0,
            raw_tilt / SPECTRAL_TILT_RANGE))
    else:
        result['spectral_tilt'] = None

    # Mid-harmonic balance: average dB delta of H3-H6 (indices 2-5)
    if n >= 6:
        mid_deltas = [live_harmonics_db[i] - baseline_harmonics_db[i]
                      for i in range(2, 6)]
        raw_mid = sum(mid_deltas) / len(mid_deltas)
        result['mid_harmonic'] = max(-1.0, min(1.0,
            raw_mid / MID_HARMONIC_RANGE))
    else:
        result['mid_harmonic'] = None

    return result


def compute_fingerprint(sessions, sax_type="Tenor"):
    """Compute an aggregate harmonic fingerprint from all sessions in a profile.

    Descriptors are always recomputed from harmonics_db using current
    formulas — stored descriptors (if present) are ignored.  This means
    formula improvements apply retroactively to all historical captures.

    First averages captures within each note, computes descriptors from
    the averaged harmonics, then averages descriptors across notes with
    equal weight per note. This prevents register skew.

    Returns dict with:
        'harmonics_db': averaged harmonic dB curve across all notes
        'descriptors': computed descriptor values (equal weight per note)
        'note_count': total unique notes
        'capture_count': total captures
        'per_note': dict of note_name -> averaged capture with computed descriptors
    """
    all_captures = []
    per_note = {}

    for session in sessions:
        for cap in session.get('captures', []):
            note = cap.get('note', '')
            entry = {
                'harmonics_db': cap.get('harmonics_db', []),
                'fundamental_freq': cap.get('fundamental_freq', cap.get('freq', 0)),
            }
            all_captures.append(entry)

            if note not in per_note:
                per_note[note] = []
            per_note[note].append(entry)

    # Average per-note first, then compute descriptors from averaged harmonics
    per_note_avg = {}
    for note, caps in per_note.items():
        avg = average_captures(caps)
        if avg:
            avg['descriptors'] = descriptors_from_harmonics(
                avg['harmonics_db'], avg['fundamental_freq'], sax_type,
                harmonic_cents=avg.get('harmonic_cents'))
        per_note_avg[note] = avg

    # Horn-level descriptors: average the per-note descriptors (equal weight per note)
    desc_keys = ['richness', 'warmth', 'core_tone', 'even_odd', 'rolloff_shape']
    if per_note_avg:
        horn_descriptors = {}
        for key in desc_keys:
            values = [pn['descriptors'].get(key, 0.0) for pn in per_note_avg.values()
                      if pn and pn.get('descriptors')]
            horn_descriptors[key] = sum(values) / len(values) if values else 0.0
        # Evenness: how uniform is the tone across the register?
        # Low stdev = even tone, high stdev = character changes by register.
        # Inverted so 100% = perfectly even, 0% = wildly variable.
        # StdDev of richness across notes, capped at 40% (beyond = 0% evenness).
        rich_vals = [pn['descriptors'].get('richness', 0.0)
                     for pn in per_note_avg.values()
                     if pn and pn.get('descriptors')]
        if len(rich_vals) >= 5:
            mean_r = sum(rich_vals) / len(rich_vals)
            stdev_r = (sum((v - mean_r) ** 2 for v in rich_vals) / len(rich_vals)) ** 0.5
            horn_descriptors['evenness'] = max(0.0, min(1.0, 1.0 - stdev_r / 0.40))
        else:
            horn_descriptors['evenness'] = 0.5  # Not enough data
    else:
        horn_descriptors = {k: 0.0 for k in desc_keys}
        horn_descriptors['evenness'] = 0.5

    # Horn-level harmonics: still average all captures
    overall = average_captures(all_captures) if all_captures else None

    overall_harmonics = overall['harmonics_db'] if overall else []

    # Determine mic type from sessions (use most common non-empty value)
    mic_types = [s.get('mic_type', '') for s in sessions if s.get('mic_type')]
    if mic_types:
        from collections import Counter
        mic_type = Counter(mic_types).most_common(1)[0][0]
    else:
        mic_type = ''

    return {
        'harmonics_db': overall_harmonics,
        'descriptors': horn_descriptors,
        'note_count': len(per_note),
        'capture_count': len(all_captures),
        'per_note': per_note_avg,
        'rolloff_rate': compute_rolloff_rate(overall_harmonics),
        'mic_type': mic_type,
    }



def compute_session_fingerprint(session, sax_type="Tenor"):
    """Compute a fingerprint for a single session.

    Same algorithm as compute_fingerprint but for one session only.
    Returns the same dict shape (harmonics_db, descriptors, note_count,
    capture_count, per_note).  Returns None if the session has no captures.
    """
    if not session or not session.get('captures'):
        return None
    return compute_fingerprint([session], sax_type)


def compute_session_variation(sessions, sax_type="Tenor"):
    """Compute descriptor variation across sessions within a profile.

    Returns dict with:
        'session_fingerprints': list of (date, fingerprint) tuples
        'descriptor_stats': {descriptor_name: {mean, stdev, min, max, n}}
        'session_count': int (sessions with captures)
    Returns None if fewer than 2 sessions have captures.
    """
    fps = []
    for s in sessions:
        fp = compute_session_fingerprint(s, sax_type)
        if fp and fp['capture_count'] > 0:
            fps.append((s.get('date', '?'), fp))
    if len(fps) < 2:
        return None

    desc_keys = ['richness', 'warmth', 'core_tone', 'even_odd', 'rolloff_shape']
    stats = {}
    for key in desc_keys:
        values = [fp['descriptors'].get(key, 0.0) for _, fp in fps]
        n = len(values)
        mean = sum(values) / n
        variance = sum((v - mean) ** 2 for v in values) / n
        stdev = math.sqrt(variance)
        stats[key] = {
            'mean': mean, 'stdev': stdev,
            'min': min(values), 'max': max(values), 'n': n,
        }
    return {
        'session_fingerprints': fps,
        'descriptor_stats': stats,
        'session_count': len(fps),
    }


def compute_group_fingerprint(profiles, sax_type=None):
    """Compute an aggregate fingerprint across multiple profiles.

    Args:
        profiles: list of (name, profile_data) tuples
        sax_type: if None, uses each profile's own horn_type

    Averages per-profile fingerprints with equal weight per profile.

    Returns dict with:
        'descriptors': averaged descriptors across profiles
        'harmonics_db': averaged harmonic curve
        'descriptor_stats': {key: {mean, stdev, min, max, n}}
        'profile_count': int
        'total_captures': int
        'per_profile': list of (name, fingerprint)
    """
    per_profile = []
    for name, pdata in profiles:
        st = sax_type or pdata.get('horn_type', 'Tenor')
        fp = compute_fingerprint(pdata.get('sessions', []), st)
        if fp['capture_count'] > 0:
            per_profile.append((name, fp))

    if not per_profile:
        return {
            'descriptors': {k: 0.0 for k in
                           ['richness', 'warmth']},
            'harmonics_db': [],
            'descriptor_stats': {},
            'profile_count': 0,
            'total_captures': 0,
            'per_profile': [],
        }

    desc_keys = ['richness', 'warmth', 'core_tone', 'even_odd', 'rolloff_shape']

    # Average descriptors across profiles (equal weight per profile)
    avg_desc = {}
    stats = {}
    for key in desc_keys:
        values = [fp['descriptors'].get(key, 0.0) for _, fp in per_profile]
        n = len(values)
        mean = sum(values) / n
        variance = sum((v - mean) ** 2 for v in values) / n
        stdev = math.sqrt(variance)
        avg_desc[key] = mean
        stats[key] = {
            'mean': mean, 'stdev': stdev,
            'min': min(values), 'max': max(values), 'n': n,
        }

    # Average harmonic curves
    max_len = max((len(fp['harmonics_db']) for _, fp in per_profile), default=0)
    avg_hdb = [0.0] * max_len
    counts = [0] * max_len
    for _, fp in per_profile:
        for i, db in enumerate(fp.get('harmonics_db', [])):
            avg_hdb[i] += db
            counts[i] += 1
    for i in range(max_len):
        if counts[i] > 0:
            avg_hdb[i] /= counts[i]

    total_caps = sum(fp['capture_count'] for _, fp in per_profile)

    return {
        'descriptors': avg_desc,
        'harmonics_db': avg_hdb,
        'descriptor_stats': stats,
        'profile_count': len(per_profile),
        'total_captures': total_caps,
        'per_profile': per_profile,
    }


def analyze_audio_file(filepath, engine, progress_cb=None):
    """Analyze an audio file offline. Returns list of capture dicts.

    Loads the file, slides a window through it, detects stable note
    segments, and extracts averaged captures — same data as live capture.

    Supports WAV files (via stdlib wave module). For other formats,
    the file must first be converted to WAV.

    Args:
        filepath: Path to a WAV audio file
        engine: A TonerEngine instance (configured with sax type, ref pitch, etc.)
        progress_cb: Optional callback(current, total) called periodically

    Returns:
        list of capture dicts (same format as session captures)
    """
    import wave
    import struct

    # Load WAV file
    with wave.open(filepath, 'rb') as wf:
        n_channels = wf.getnchannels()
        sampwidth = wf.getsampwidth()
        framerate = wf.getframerate()
        n_frames = wf.getnframes()
        raw = wf.readframes(n_frames)

    # Convert to float32 mono (16-bit, 24-bit, 32-bit int, 32-bit float)
    n_samples = n_frames * n_channels
    if sampwidth == 2:
        samples = np.array(struct.unpack(f'<{n_samples}h', raw),
                          dtype=np.float32) / 32768.0
    elif sampwidth == 3:
        # 24-bit audio: pad each 3-byte sample to 4 bytes (sign-extend)
        padded = bytearray(4 * n_samples)
        for i in range(n_samples):
            padded[i * 4 + 1] = raw[i * 3]
            padded[i * 4 + 2] = raw[i * 3 + 1]
            padded[i * 4 + 3] = raw[i * 3 + 2]
        samples = np.array(struct.unpack(f'<{n_samples}i', bytes(padded)),
                          dtype=np.float32) / 2147483648.0
    elif sampwidth == 4:
        # Try 32-bit float first (common DAW export), fall back to 32-bit int
        samples = np.frombuffer(raw, dtype=np.float32).copy()
        if np.any(np.abs(samples) > 2.0):
            # Values > 2.0 means it's 32-bit int, not float
            samples = np.array(struct.unpack(f'<{n_samples}i', raw),
                              dtype=np.float32) / 2147483648.0
    else:
        return []

    # Mix to mono if stereo
    if n_channels > 1:
        samples = samples.reshape(-1, n_channels).mean(axis=1)

    # Resample if needed (simple decimation/interpolation)
    if framerate != SAMPLE_RATE:
        duration = len(samples) / framerate
        target_len = int(duration * SAMPLE_RATE)
        indices = np.linspace(0, len(samples) - 1, target_len)
        samples = np.interp(indices, np.arange(len(samples)), samples).astype(np.float32)

    # Slide through the file, analyzing every hop
    results = []  # list of (note_name, result)
    hop = FILE_HOP_SAMPLES
    total_windows = max(1, (len(samples) - FFT_SIZE) // hop)
    window_i = 0
    for start in range(0, len(samples) - FFT_SIZE, hop):
        chunk = samples[start:start + FFT_SIZE]
        if len(chunk) < FFT_SIZE:
            break

        r = engine.analyze_buffer(chunk)
        if r.fundamental_freq > 0 and r.harmonics:
            results.append((r.fundamental_note, r))

        window_i += 1
        if progress_cb and window_i % 200 == 0:
            progress_cb(window_i, total_windows)

    if not results:
        return []

    # Detect stable segments and build captures
    captures = []
    segment_start = 0
    current_note = results[0][0]

    def _process_segment(seg, seg_note):
        """Process a completed segment into a capture."""
        if len(seg) > ATTACK_SKIP_FRAMES:
            seg = seg[ATTACK_SKIP_FRAMES:]  # Skip attack
        if len(seg) < FREE_MIN_FRAMES:
            return None
        # Build a capture from this segment
        frames = []
        for _, r in seg:
            frames.append({
                'note': r.fundamental_note,
                'freq': r.fundamental_freq,
                'harmonics_db': [h.magnitude_db for h in r.harmonics],
                'harmonic_cents': [h.cents_deviation for h in r.harmonics],
            })

        avg = average_captures(frames)
        if avg:
            avg_freq = sum(f['freq'] for f in frames) / len(frames)
            capture = {
                'note': seg_note,
                'fundamental_freq': round(avg_freq, 2),
                'harmonics_db': [round(db, 2) for db in avg['harmonics_db']],
                'timestamp': '',  # Filled by caller
                'n_frames': len(frames),
                'method': 'file',
            }
            if avg.get('harmonic_cents'):
                capture['harmonic_cents'] = [round(c, 2) for c in avg['harmonic_cents']]
            # Spectral centroid from harmonics
            h_db = avg['harmonics_db']
            if h_db and avg_freq > 0:
                lin = [10.0 ** (db / 20.0) for db in h_db]
                freqs = [avg_freq * (i + 1) for i in range(len(h_db))]
                amp_sum = sum(lin)
                if amp_sum > 0:
                    capture['spectral_centroid'] = round(
                        sum(f * a for f, a in zip(freqs, lin)) / amp_sum, 1)
            return capture
        return None

    for i in range(1, len(results)):
        note = results[i][0]
        if note != current_note:
            # Segment ended
            segment = results[segment_start:i]
            capture = _process_segment(segment, current_note)
            if capture:
                captures.append(capture)

            current_note = note
            segment_start = i

    # Process the final segment
    segment = results[segment_start:]
    capture = _process_segment(segment, current_note)
    if capture:
        captures.append(capture)

    return captures


DEFAULT_LIBRARY = "My Profiles"


def load_tone_profiles(filepath):
    """Load tone profiles from JSON file.

    Returns nested dict: {library_name: {profile_name: profile_data}}.
    Migrates flat format (no libraries) if found.
    """
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r') as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {DEFAULT_LIBRARY: {}}

            # Check if this is the old flat format (profile names at top level
            # with 'sessions' key inside — libraries wouldn't have that)
            if data and any(isinstance(v, dict) and 'sessions' in v
                           for v in data.values()):
                # Migrate flat → nested
                migrated = {DEFAULT_LIBRARY: data}
                save_tone_profiles(migrated, filepath)
                return migrated

            return data
        except (json.JSONDecodeError, TypeError):
            pass
    return {DEFAULT_LIBRARY: {}}


def save_tone_profiles(profiles, filepath):
    """Save tone profiles to JSON file."""
    try:
        with open(filepath, 'w') as f:
            json.dump(profiles, f, indent=2)
        return True
    except Exception:
        return False


def flatten_profiles(profiles):
    """Return a flat dict {profile_name: profile_data} from nested library structure.

    If duplicate names exist across libraries, they're prefixed with library name.
    """
    flat = {}
    for lib_name, lib_profiles in profiles.items():
        if not isinstance(lib_profiles, dict):
            continue
        for prof_name, prof_data in lib_profiles.items():
            key = prof_name
            if key in flat:
                key = f"[{lib_name}] {prof_name}"
            flat[key] = prof_data
    return flat
