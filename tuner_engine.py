"""
Strobe tuner audio engine for Stohrer Sax Shop Companion.

Handles audio capture, FFT pitch analysis, phase tracking, and reference tone
generation. Pure math/audio — no tkinter dependency.

12 chromatic pitch classes, each with concentric rings showing different
octaves. Phase tracking drives the stroboscopic rotation effect.

Requires: numpy, sounddevice (imported with try/except for graceful fallback)
"""

import math
import time

try:
    import numpy as np
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except (ImportError, OSError):
    AUDIO_AVAILABLE = False
    np = None
    sd = None

from audio_utils import AudioRingBuffer  # noqa: E402 — shared with toner_engine


# ============================================
# CONSTANTS
# ============================================

SAMPLE_RATE = 44100
BUFFER_SECONDS = 0.2  # 200ms ring buffer
BUFFER_SIZE = int(SAMPLE_RATE * BUFFER_SECONDS)
FFT_SIZE = 4096  # ~93ms at 44100Hz, ~10.77Hz bin resolution

PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

# Strobe disc physics: drift rate is computed from disc RPM.
# Each pitch class disc spins at a different speed. The A disc spins at
# 27.5 rev/sec (440 Hz / 16 segments).
# Drift = ln(2)/1200 * disc_rps * 360 degrees/sec/cent
# disc_rps = reference_freq_for_pitch_class / DISC_BASE_SEGMENTS
# Using the 16-segment ring (octave 4) as reference:
DISC_BASE_SEGMENTS = 16

# Octave range for analysis
MIN_OCTAVE = 1
MAX_OCTAVE = 7

# --- Analysis tuning parameters ---
# Multiplier applied to median spectrum magnitude for adaptive noise floor
NOISE_FLOOR_MULTIPLIER = 3.0
# Sensitivity range factor: maps sensitivity 0-100 to threshold scale 1.0-7.0
SENSITIVITY_RANGE_FACTOR = 0.06
# Minimum frequency (Hz) to consider in FFT analysis
MIN_FREQUENCY_HZ = 25
# Cents error clamp: maximum ±cents before clamping (1 semitone)
CENTS_CLAMP = 100.0
# Fraction of max magnitude below which a wheel is considered inactive
ACTIVE_THRESHOLD = 0.05
# Consecutive stale audio reads before triggering stream restart (~1s at 60fps)
STALE_RESTART_THRESHOLD = 60
# Per-ring temporal smoothing — exponential decay factor per frame.
# Lower = slower fade (more inertia). At 60fps, 0.3 gives ~100ms decay.
RING_SMOOTH_ATTACK = 0.5    # How fast a ring brightens (rise)
RING_SMOOTH_DECAY = 0.15    # How fast a ring fades (fall) — slower = more fuzzy lingering


# ============================================
# ANALYSIS RESULT
# ============================================

NUM_RINGS = MAX_OCTAVE - MIN_OCTAVE + 1  # 7 rings = 7 octaves

class TunerResult:
    """Result of one analysis frame."""
    __slots__ = ['magnitudes', 'phase_offsets', 'cents_errors', 'active',
                 'ring_magnitudes', 'ring_phase_offsets']

    def __init__(self):
        self.magnitudes = [0.0] * 12       # Energy per pitch class (0-1 normalized)
        self.phase_offsets = [0.0] * 12     # Accumulated rotation angle in degrees
        self.cents_errors = [0.0] * 12      # Current cents error from reference
        self.active = [False] * 12          # Whether wheel is "lit"
        # Per-ring (per-octave) magnitudes for each pitch class.
        # ring_magnitudes[pc][ring_idx] = energy at that specific octave.
        # Drives per-ring brightness — the played octave's ring appears
        # sharp/bright while other rings are dim/fuzzy.
        self.ring_magnitudes = [[0.0] * NUM_RINGS for _ in range(12)]
        # Per-ring phase offsets — each ring tracks its own octave's frequency
        # independently, showing inharmonicity across octaves — each ring
        # responds to its own frequency component.
        self.ring_phase_offsets = [[0.0] * NUM_RINGS for _ in range(12)]


# ============================================
# TUNER ENGINE
# ============================================

class TunerEngine:
    """Audio capture and strobe tuner analysis engine."""

    def __init__(self):
        self._stream = None
        self._ring_buffer = None
        self._reference_pitch = 440.0
        self._sensitivity = 50
        self._freq_table = None   # freq_table[pc][oct_idx] = frequency in Hz
        self._phase_offsets = [0.0] * 12
        self._drift_rates = [0.0] * 12  # Per-pitch-class drift rates (deg/sec/cent)
        self._last_time = None
        self._running = False
        self._window = None
        self._last_device = None   # For auto-restart
        self._stale_count = 0      # Consecutive stale reads
        self.last_error = None     # Set when stream restart fails
        # Per-ring smoothed magnitudes — temporal decay like physical disc inertia
        self._smoothed_ring_mags = [[0.0] * NUM_RINGS for _ in range(12)]
        # Per-ring independent phase accumulators — each ring tracks its own
        # octave's frequency, showing real inharmonicity
        self._ring_phase_offsets = [[0.0] * NUM_RINGS for _ in range(12)]
        self._build_freq_table()

    def _build_freq_table(self):
        """Build reference frequency table for all pitch classes and octaves.

        freq_table[pc][oct_idx] gives the frequency for pitch class pc
        at octave (MIN_OCTAVE + oct_idx). pc 0 = C, pc 9 = A.

        Also computes per-pitch-class drift rates matching strobe disc physics.
        Each disc spins at disc_rps = freq_octave4 / DISC_BASE_SEGMENTS.
        Drift = ln(2)/1200 * disc_rps * 360 degrees/sec/cent.
        """
        self._freq_table = []
        for pc in range(12):
            octave_freqs = []
            for octave in range(MIN_OCTAVE, MAX_OCTAVE + 1):
                # Semitones from A4: (pc - 9) + (octave - 4) * 12
                semitones = (pc - 9) + (octave - 4) * 12
                freq = self._reference_pitch * (2.0 ** (semitones / 12.0))
                octave_freqs.append(freq)
            self._freq_table.append(octave_freqs)

            # Drift rate for this pitch class: matches physical strobe disc speed
            # freq_oct4 = reference_pitch * 2^((pc-9)/12)
            freq_oct4 = self._reference_pitch * (2.0 ** ((pc - 9) / 12.0))
            disc_rps = freq_oct4 / DISC_BASE_SEGMENTS
            self._drift_rates[pc] = math.log(2) / 1200.0 * disc_rps * 360.0

    @property
    def reference_pitch(self):
        return self._reference_pitch

    def set_reference_pitch(self, hz):
        """Set reference pitch (e.g. 440.0) and rebuild frequency table."""
        self._reference_pitch = float(hz)
        self._build_freq_table()

    def set_sensitivity(self, value):
        """Set sensitivity 0-100. Higher = responds to quieter signals."""
        self._sensitivity = max(0, min(100, int(value)))

    def start(self, device=None):
        """Start audio capture. Returns (success, error_message)."""
        if not AUDIO_AVAILABLE:
            return False, "Audio libraries not available.\nInstall numpy and sounddevice:\n  pip install numpy sounddevice"

        if self._running:
            self.stop()

        self._ring_buffer = AudioRingBuffer(BUFFER_SIZE)
        self._window = np.hanning(FFT_SIZE).astype(np.float32)
        self._phase_offsets = [0.0] * 12
        self._last_time = time.perf_counter()
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
        """Stop audio capture."""
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
        """Sounddevice input callback (audio thread)."""
        if self._ring_buffer is not None:
            self._ring_buffer.write(indata[:, 0])

    def analyze(self):
        """Analyze current audio buffer. Returns TunerResult.

        Monitors stream health and auto-restarts if the audio callback
        appears to have died (no new data for ~1 second).

        For each pitch class:
        - Sums FFT magnitude across all octaves → magnitudes[pc]
        - Finds dominant frequency via parabolic interpolation
        - Computes cents error from reference → drives phase accumulation
        - Phase offset drives the stroboscopic rotation of that wheel
        """
        result = TunerResult()

        buf = self._ring_buffer
        if not self._running or buf is None:
            return result

        # Check stream health
        if buf.is_stale():
            self._stale_count += 1
            if self._stale_count > STALE_RESTART_THRESHOLD:
                self._stale_count = 0
                self._restart_stream()
                return result
        else:
            self._stale_count = 0

        audio = buf.read()
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
        """Analyze a raw audio buffer. Used by analyze() and tests.

        Args:
            audio: numpy float32 array of audio samples
        Returns:
            TunerResult
        """
        result = TunerResult()

        if len(audio) < FFT_SIZE:
            return result

        now = time.perf_counter()
        dt = now - self._last_time if self._last_time else 1 / 60.0
        self._last_time = now
        dt = min(dt, 0.1)  # Clamp to avoid huge phase jumps

        # Take the most recent FFT_SIZE samples
        frame = audio[-FFT_SIZE:]

        # Ensure we have a window
        if self._window is None or len(self._window) != FFT_SIZE:
            self._window = np.hanning(FFT_SIZE).astype(np.float32)

        # Apply Hanning window and compute FFT
        windowed = frame * self._window
        spectrum = np.fft.rfft(windowed)
        mags = np.abs(spectrum)

        bin_freq = SAMPLE_RATE / FFT_SIZE  # ~10.77 Hz

        # Adaptive noise floor threshold
        noise_floor = np.median(mags[10:]) * NOISE_FLOOR_MULTIPLIER if len(mags) > 10 else 0.0
        sensitivity_scale = 1.0 + (100 - self._sensitivity) * SENSITIVITY_RANGE_FACTOR  # 1.0 to 5.0
        threshold = noise_floor * sensitivity_scale

        max_mag = 0.0

        for pc in range(12):
            total_mag = 0.0
            best_mag = 0.0
            best_octave = -1

            for oct_idx, freq in enumerate(self._freq_table[pc]):
                if freq < MIN_FREQUENCY_HZ or freq > SAMPLE_RATE / 2:
                    continue

                bin_idx = int(round(freq / bin_freq))
                if bin_idx < 1 or bin_idx >= len(mags) - 1:
                    continue

                # Peak magnitude at this bin and immediate neighbors
                local_mags = [mags[bin_idx - 1], mags[bin_idx], mags[bin_idx + 1]]
                peak_offset = int(np.argmax(local_mags)) - 1  # -1, 0, or +1
                peak_bin = bin_idx + peak_offset
                mag = local_mags[peak_offset + 1]

                if mag > threshold:
                    total_mag += mag
                    result.ring_magnitudes[pc][oct_idx] = float(mag)
                    if mag > best_mag:
                        best_mag = mag
                        best_octave = oct_idx

                    # Per-ring phase tracking — each ring independently
                    # measures its own octave's frequency via parabolic
                    # interpolation — each ring responds to its own
                    # frequency component independently.
                    if peak_bin > 0 and peak_bin < len(mags) - 1:
                        alpha = float(mags[peak_bin - 1])
                        beta = float(mags[peak_bin])
                        gamma = float(mags[peak_bin + 1])
                        denom = alpha - 2 * beta + gamma
                        if abs(denom) > 1e-10 and beta > 0:
                            p = 0.5 * (alpha - gamma) / denom
                            ring_freq = (peak_bin + p) * bin_freq
                        else:
                            ring_freq = peak_bin * bin_freq

                        if ring_freq > 0 and freq > 0:
                            ring_cents = 1200.0 * math.log2(ring_freq / freq)
                            ring_cents = max(-CENTS_CLAMP, min(CENTS_CLAMP, ring_cents))
                            self._ring_phase_offsets[pc][oct_idx] += (
                                ring_cents * self._drift_rates[pc] * dt
                            )
                            self._ring_phase_offsets[pc][oct_idx] %= 360.0

            result.magnitudes[pc] = total_mag
            if total_mag > max_mag:
                max_mag = total_mag

            # Overall cents error and phase from strongest octave (for VU meter)
            if best_octave >= 0:
                result.cents_errors[pc] = 0.0
                best_freq = self._freq_table[pc][best_octave]
                best_bin_idx = int(round(best_freq / bin_freq))
                if 0 < best_bin_idx < len(mags) - 1:
                    local_mags = [mags[best_bin_idx - 1], mags[best_bin_idx], mags[best_bin_idx + 1]]
                    po = int(np.argmax(local_mags)) - 1
                    pb = best_bin_idx + po
                    if 0 < pb < len(mags) - 1:
                        a2 = float(mags[pb - 1])
                        b2 = float(mags[pb])
                        g2 = float(mags[pb + 1])
                        d2 = a2 - 2 * b2 + g2
                        if abs(d2) > 1e-10 and b2 > 0:
                            p2 = 0.5 * (a2 - g2) / d2
                            af = (pb + p2) * bin_freq
                        else:
                            af = pb * bin_freq
                        if af > 0 and best_freq > 0:
                            cents = 1200.0 * math.log2(af / best_freq)
                            cents = max(-CENTS_CLAMP, min(CENTS_CLAMP, cents))
                            result.cents_errors[pc] = cents
                            self._phase_offsets[pc] += cents * self._drift_rates[pc] * dt
                            self._phase_offsets[pc] %= 360.0

            result.phase_offsets[pc] = self._phase_offsets[pc]
            for r in range(NUM_RINGS):
                if result.ring_magnitudes[pc][r] > 0:
                    # This ring has signal — use its independently tracked phase
                    result.ring_phase_offsets[pc][r] = self._ring_phase_offsets[pc][r]
                else:
                    # No signal at this octave — inherit overall wheel phase
                    # so the dim background spins coherently like a real disc.
                    # Also sync the accumulator so phase is smooth when signal
                    # first appears on this ring.
                    self._ring_phase_offsets[pc][r] = self._phase_offsets[pc]
                    result.ring_phase_offsets[pc][r] = self._phase_offsets[pc]

        # Normalize magnitudes to 0-1, but only if the strongest signal
        # is meaningfully above the noise floor (not just noise peaks)
        if max_mag > threshold * 1.5:
            for pc in range(12):
                result.magnitudes[pc] /= max_mag
                for r in range(NUM_RINGS):
                    result.ring_magnitudes[pc][r] /= max_mag
        else:
            # No real signal — zero everything so wheels stay dark
            for pc in range(12):
                result.magnitudes[pc] = 0.0
                for r in range(NUM_RINGS):
                    result.ring_magnitudes[pc][r] = 0.0

        # Temporal smoothing on per-ring magnitudes — simulates physical disc
        # inertia.  Strobe rings fade in/out gradually because the
        # illumination has persistence and the human eye integrates.
        # Fast attack (ring lights up quickly), slow decay (fades out gradually,
        # giving the "fuzzy" look on secondary octave rings).
        for pc in range(12):
            for r in range(NUM_RINGS):
                target = result.ring_magnitudes[pc][r]
                prev = self._smoothed_ring_mags[pc][r]
                if target > prev:
                    alpha = RING_SMOOTH_ATTACK
                else:
                    alpha = RING_SMOOTH_DECAY
                smoothed = prev + alpha * (target - prev)
                self._smoothed_ring_mags[pc][r] = smoothed
                result.ring_magnitudes[pc][r] = smoothed

        # Determine active wheels
        for pc in range(12):
            result.active[pc] = result.magnitudes[pc] > ACTIVE_THRESHOLD

        return result

    def reset_phases(self):
        """Reset all phase offsets to zero."""
        self._phase_offsets = [0.0] * 12
        self._ring_phase_offsets = [[0.0] * NUM_RINGS for _ in range(12)]


# ============================================
# REFERENCE TONE PLAYER
# ============================================

class ReferencePlayer:
    """Plays reference tones via sounddevice output stream."""

    def __init__(self):
        self._stream = None
        self._playing = False
        self._frequency = 440.0
        self._waveform = "pure"
        self._sample_idx = 0

    def play(self, frequency, waveform="pure"):
        """Start playing a reference tone.

        Args:
            frequency: Tone frequency in Hz
            waveform: "pure" (sine) or "rich" (harmonics)
        Returns:
            True if started successfully
        """
        if not AUDIO_AVAILABLE:
            return False

        self.stop()

        self._frequency = frequency
        self._waveform = waveform
        self._sample_idx = 0

        try:
            self._stream = sd.OutputStream(
                samplerate=SAMPLE_RATE,
                channels=1,
                dtype='float32',
                blocksize=1024,
                callback=self._output_callback,
            )
            self._stream.start()
            self._playing = True
            return True
        except Exception:
            return False

    def stop(self):
        """Stop playing."""
        self._playing = False
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None

    @property
    def is_playing(self):
        return self._playing

    def _output_callback(self, outdata, frames, time_info, status):
        """Sounddevice output callback (audio thread)."""
        t = (self._sample_idx + np.arange(frames, dtype=np.float64)) / SAMPLE_RATE
        freq = self._frequency

        if self._waveform == "pure":
            signal = 0.3 * np.sin(2 * np.pi * freq * t)
        else:
            # Rich tone: fundamental + harmonics (2nd at -6dB, 3rd at -12dB, 4th at -18dB)
            signal = 0.20 * np.sin(2 * np.pi * freq * t)
            signal += 0.10 * np.sin(2 * np.pi * 2 * freq * t)
            signal += 0.05 * np.sin(2 * np.pi * 3 * freq * t)
            signal += 0.025 * np.sin(2 * np.pi * 4 * freq * t)

        self._sample_idx += frames
        outdata[:, 0] = signal.astype(np.float32)
