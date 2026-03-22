"""
Tone analyzer engine for Stohrer Sax Shop Companion.

Handles audio capture, FFT, fundamental pitch detection (peak-picking with harmonic series verification),
harmonic extraction, and tone descriptor computation. Pure math/audio —
no tkinter dependency.

Saxophone-specific: analyzes both even and odd harmonics (conical bore),
up to the 12th harmonic, in the range relevant to saxophone (Bb2–F#6
fundamentals, harmonics up to ~8kHz).

Requires: numpy, sounddevice (imported with try/except for graceful fallback)
"""

import math
import threading
import time
import json
import os
from collections import namedtuple

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

MAX_HARMONICS = 12            # Analyze up to 12th harmonic
SPECTRUM_MAX_HZ = 8000        # Display range for spectrum

PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

# Minimum fundamental frequency. Bari sax written Bb3 = concert Db2 (~69 Hz),
# so 65 Hz covers all standard saxophone types including baritone.
MIN_FUNDAMENTAL_HZ = 65.0
# Maximum fundamental frequency (altissimo range ~ 1500 Hz)
MAX_FUNDAMENTAL_HZ = 2000.0

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
        'descriptors',           # dict: resonance, richness, brightness, darkness, fullness
        'signal_level',          # 0.0-1.0, RMS input level
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
            'resonance': 0.5,
            'richness': 0.0,
            'brightness': 0.0,
            'darkness': 0.0,
            'fullness': 0.0,
        }
        self.signal_level = 0.0


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
        return self._low_energy_frames > self._spectral_check_frames * 0.6

    def set_break_frequency(self, hz):
        """Set the Benade break frequency for brightness/darkness calculation."""
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

        if not self._running or self._ring_buffer is None:
            return result

        # Check stream health: if buffer is stale, the callback may have died
        if self._ring_buffer.is_stale():
            self._stale_count += 1
            # After ~1 second of stale reads (30 frames at 30fps), restart
            if self._stale_count > 30:
                self._stale_count = 0
                self._restart_stream()
                return result
        else:
            self._stale_count = 0

        audio = self._ring_buffer.read()
        if audio is None:
            return result

        return self.analyze_buffer(audio)

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
        result.signal_level = min(1.0, rms * 10.0)  # Scale for display

        # Sensitivity threshold
        sensitivity_scale = 1.0 + (100 - self._sensitivity) * 0.04
        min_rms = 0.002 * sensitivity_scale
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
                if mid_energy > 0 and low_energy < mid_energy * 0.05:
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
            noise_thresh = noise_floor * (3.0 if divisor <= 3 else 2.0)
            if local_mag < noise_thresh:
                continue
            left_mag = float(mags[max(0, local_peak - 4)])
            right_mag = float(mags[min(len(mags) - 1, local_peak + 4)])
            prominence = 1.5 if divisor <= 3 else 1.3
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
                if h_peak_mag > noise_floor * 3.0:
                    harmonics_found += 1

            required = 2 if divisor <= 3 else 3
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
                    if new_mag < prev_mag * 1.5:
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

            # Parabolic interpolation
            if local_peak_bin > 0 and local_peak_bin < len(mags) - 1:
                alpha = float(mags[local_peak_bin - 1])
                beta = float(mags[local_peak_bin])
                gamma = float(mags[local_peak_bin + 1])
                denom = alpha - 2 * beta + gamma
                if abs(denom) > 1e-10 and beta > 0:
                    p = 0.5 * (alpha - gamma) / denom
                    actual_freq = (local_peak_bin + p) * bin_freq
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
            if n > 1 and mag_db < -60.0:
                continue

            # Cents deviation from ideal harmonic position
            if actual_freq > 0 and expected_freq > 0:
                cents_dev = 1200.0 * math.log2(actual_freq / expected_freq)
                cents_dev = max(-100.0, min(100.0, cents_dev))
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

    def _compute_descriptors(self, harmonics, f0=440.0):
        """Compute tone quality descriptors from harmonic data.

        Uses the Benade break frequency to divide "lower" from "upper"
        harmonics. Below f_b, the sax body naturally amplifies harmonics;
        above f_b, they roll off. This makes brightness/darkness
        physically meaningful across soprano, alto, tenor, and bari.
        """
        if not harmonics:
            return {'resonance': 0.5, 'richness': 0.0, 'brightness': 0.0,
                    'darkness': 0.0, 'fullness': 0.0}

        f_b = self._break_freq

        # --- Resonance: how well harmonics align to ideal positions ---
        # Scaled for saxophone reality: even a mediocre horn is within
        # a few cents. The gauge maps 85-100% of the raw resonance score
        # to the full 0-100% range. Most horns read upper half (as they
        # should), great horns peg right, genuinely bad bores read low.
        # Raw 85% (7.5 cents mean dev) = gauge 0%, raw 100% = gauge 100%.
        deviations = [abs(h.cents_deviation) for h in harmonics if h.harmonic_number > 1]
        if deviations:
            mean_dev = sum(deviations) / len(deviations)
            raw = max(0.0, min(1.0, 1.0 - mean_dev / 50.0))
            # Rescale 0.85-1.0 to 0.0-1.0
            resonance = max(0.0, min(1.0, (raw - 0.85) / 0.15))
        else:
            resonance = 0.5

        # --- Richness: spectral flatness of significant harmonics ---
        # Rescaled for sax-vs-sax comparison: raw flatness*coverage ranges
        # from ~0.50 (thin/uneven harmonics) to ~0.95 (very even spread).
        # The gauge maps 0.50-0.95 to 0-100%, giving useful range between
        # saxophones rather than pegging near 100% for all of them.
        sig_threshold_db = -35.0
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
            # Rescale 0.50-0.95 to 0.0-1.0 (sax-specific range)
            richness = max(0.0, min(1.0, (raw_richness - 0.50) / 0.45))
        elif len(significant) == 1:
            richness = 0.05
        else:
            richness = 0.0

        # --- Brightness & Darkness using Benade break frequency ---
        # Harmonics with frequency above f_b contribute to brightness.
        # Harmonics with frequency at or below f_b contribute to darkness.
        # Brightness/darkness: ratio of harmonic energy above/below break.
        # The fundamental always counts toward darkness (it IS the body of
        # the sound, always at or below the break for normal sax range).
        # This ensures the gauges have useful range even on high notes
        # where all upper harmonics are above the break.
        upper_energy = 0.0
        lower_energy = 0.0
        total_energy = 0.0
        for h in harmonics:
            freq = f0 * h.harmonic_number
            linear_mag = 10.0 ** (h.magnitude_db / 20.0)
            total_energy += linear_mag
            if h.harmonic_number == 1:
                lower_energy += linear_mag  # Fundamental always counts as dark
            elif freq > f_b:
                upper_energy += linear_mag
            else:
                lower_energy += linear_mag

        brightness = 0.0
        if total_energy > 0:
            brightness = upper_energy / total_energy

        darkness = 0.0
        if total_energy > 0:
            darkness = lower_energy / total_energy

        # --- Fullness: balance of bright and dark energy ---
        # A "full" tone has strong harmonics on BOTH sides of the break
        # frequency — neither bright nor dark dominates. A thin or harsh
        # tone is lopsided toward one side. Peaks at 50/50 bright/dark,
        # drops toward zero as either side takes over.
        # Uses 1 - |b - d| rather than min(b,d)*k for a more gradual curve.
        # Energy weight ensures a thin tone with coincidentally balanced
        # bright/dark doesn't read as full.
        # Squared to penalize imbalance harder — a 70/30 split reads much
        # lower than 55/45. Calibrated: BA tenor "full" ≈ 50%, SBA "not full" ≈ 38%.
        balance = (1.0 - abs(brightness - darkness)) ** 2.0
        fund_linear = 10.0 ** (harmonics[0].magnitude_db / 20.0) if harmonics else 0.0
        non_fund_energy = total_energy - fund_linear
        # How much harmonic energy exists relative to fundamental
        energy_factor = min(1.0, non_fund_energy / max(fund_linear, 1e-10) / 2.0)
        fullness = balance * (0.6 + 0.4 * energy_factor)

        return {
            'resonance': resonance,
            'richness': richness,
            'brightness': brightness,
            'darkness': darkness,
            'fullness': fullness,
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

SAX_TYPES = ["Sopranino", "Soprano", "F Mezzo", "Alto",
             "C Melody", "Tenor", "Baritone", "Bass"]

# Benade spectral envelope break frequencies (Hz).
# Below f_b, harmonics rise ~3 dB/oct. Above, they fall ~18 dB/oct.
# The break frequency marks the acoustic boundary between "lower" and
# "upper" harmonics for brightness/darkness calculations.
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


def migrate_profile_to_concert(profile):
    """Convert a profile's stored note names from written to concert pitch.

    Uses the profile's horn_type to determine the reverse transposition.
    Returns a new deep-copied profile with concert pitch note names.
    Skips profiles already marked with pitch_format='concert'.
    """
    import copy
    profile = copy.deepcopy(profile)

    if profile.get('pitch_format') == 'concert':
        return profile

    sax_type = profile.get('horn_type', '')
    if not sax_type:
        # Can't migrate without knowing the horn type
        profile['pitch_format'] = 'concert'
        return profile

    for session in profile.get('sessions', []):
        if not isinstance(session, dict):
            continue
        for cap in session.get('captures', []):
            if not isinstance(cap, dict):
                continue
            note = cap.get('note', '')
            if note:
                cap['note'] = reverse_transpose_note(note, sax_type)
            # detected_as stays as-is (it's metadata about detector accuracy,
            # compared against the written-pitch calibration prompt)

    profile['pitch_format'] = 'concert'
    return profile


def average_captures(captures):
    """Average a list of per-frame capture dicts into one summary.

    Each capture dict has:
        'harmonics_db': list of dB values (index 0=fundamental, 1=2nd, ...)
        'descriptors': dict of descriptor values

    Returns averaged dict with same structure.
    """
    if not captures:
        return None

    n = len(captures)

    # Average harmonic dB values
    max_len = max(len(c['harmonics_db']) for c in captures)
    avg_db = [0.0] * max_len
    counts = [0] * max_len
    for c in captures:
        for i, db in enumerate(c['harmonics_db']):
            avg_db[i] += db
            counts[i] += 1
    for i in range(max_len):
        if counts[i] > 0:
            avg_db[i] /= counts[i]

    # Average descriptors
    desc_keys = ['resonance', 'richness', 'brightness', 'darkness', 'fullness']
    avg_desc = {}
    for key in desc_keys:
        total = sum(c['descriptors'].get(key, 0.0) for c in captures)
        avg_desc[key] = total / n

    return {
        'harmonics_db': avg_db,
        'descriptors': avg_desc,
    }


def compute_fingerprint(sessions):
    """Compute an aggregate harmonic fingerprint from all sessions in a profile.

    First averages captures within each note, then averages across notes
    with equal weight per note. This prevents low notes (which tend toward
    darkness) from diluting high-note brightness or vice versa — each
    note's descriptors are computed correctly for its own frequency first,
    then the horn-level summary gives each note equal say.

    Returns dict with:
        'harmonics_db': averaged harmonic dB curve across all notes
        'descriptors': averaged descriptor values (equal weight per note)
        'note_count': total unique notes
        'capture_count': total captures
        'per_note': dict of note_name -> averaged capture for that note
    """
    all_captures = []
    per_note = {}

    for session in sessions:
        for cap in session.get('captures', []):
            note = cap.get('note', '')
            entry = {
                'harmonics_db': cap.get('harmonics_db', []),
                'descriptors': cap.get('descriptors', {}),
            }
            all_captures.append(entry)

            if note not in per_note:
                per_note[note] = []
            per_note[note].append(entry)

    # Average per-note first
    per_note_avg = {}
    for note, caps in per_note.items():
        per_note_avg[note] = average_captures(caps)

    # Horn-level descriptors: average the per-note descriptors (equal weight per note)
    # This is more meaningful than averaging all raw captures, because each note's
    # brightness/darkness is computed at its own frequency relative to the break.
    desc_keys = ['resonance', 'richness', 'brightness', 'darkness', 'fullness']
    if per_note_avg:
        horn_descriptors = {}
        for key in desc_keys:
            values = [pn['descriptors'].get(key, 0.0) for pn in per_note_avg.values()
                      if pn and pn.get('descriptors')]
            horn_descriptors[key] = sum(values) / len(values) if values else 0.0
    else:
        horn_descriptors = {k: 0.0 for k in desc_keys}

    # Horn-level harmonics: still average all captures (harmonic dB is relative
    # to each note's own fundamental, so averaging is reasonable)
    overall = average_captures(all_captures) if all_captures else None

    return {
        'harmonics_db': overall['harmonics_db'] if overall else [],
        'descriptors': horn_descriptors,
        'note_count': len(per_note),
        'capture_count': len(all_captures),
        'per_note': per_note_avg,
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

    # Convert to float32 mono
    if sampwidth == 2:
        samples = np.array(struct.unpack(f'<{n_frames * n_channels}h', raw),
                          dtype=np.float32) / 32768.0
    elif sampwidth == 4:
        samples = np.array(struct.unpack(f'<{n_frames * n_channels}i', raw),
                          dtype=np.float32) / 2147483648.0
    elif sampwidth == 1:
        samples = np.array(struct.unpack(f'{n_frames * n_channels}B', raw),
                          dtype=np.float32) / 128.0 - 1.0
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
                'descriptors': dict(r.descriptors),
            })

        avg = average_captures(frames)
        if avg:
            avg_freq = sum(f['freq'] for f in frames) / len(frames)
            return {
                'note': seg_note,
                'fundamental_freq': round(avg_freq, 2),
                'harmonics_db': [round(db, 2) for db in avg['harmonics_db']],
                'descriptors': {k: round(v, 3) for k, v in avg['descriptors'].items()},
                'timestamp': '',  # Filled by caller
                'n_frames': len(frames),
                'method': 'file',
            }
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
