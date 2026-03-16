"""
Strobe tuner audio engine for Stohrer Sax Shop Companion.

Handles audio capture, FFT pitch analysis, phase tracking, and reference tone
generation. Pure math/audio — no tkinter dependency.

Modeled after the Peterson Stroboconn 6T-5: 12 chromatic pitch classes, each
with concentric rings showing different octaves. Phase tracking drives the
stroboscopic rotation effect.

Requires: numpy, sounddevice (imported with try/except for graceful fallback)
"""

import math
import threading
import time

try:
    import numpy as np
    import sounddevice as sd
    AUDIO_AVAILABLE = True
except ImportError:
    AUDIO_AVAILABLE = False
    np = None
    sd = None


# ============================================
# CONSTANTS
# ============================================

SAMPLE_RATE = 44100
BUFFER_SECONDS = 0.2  # 200ms ring buffer
BUFFER_SIZE = int(SAMPLE_RATE * BUFFER_SECONDS)
FFT_SIZE = 4096  # ~93ms at 44100Hz, ~10.77Hz bin resolution

PITCH_CLASSES = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']

# Phase drift rate: degrees per second per cent of error
# Tuned so 10 cents ≈ 36°/s (one revolution every 10 seconds)
DRIFT_RATE = 3.6

# Octave range for analysis
MIN_OCTAVE = 1
MAX_OCTAVE = 7


# ============================================
# AUDIO RING BUFFER
# ============================================

class AudioRingBuffer:
    """Thread-safe ring buffer for audio samples."""

    def __init__(self, size):
        self.buffer = np.zeros(size, dtype=np.float32)
        self.write_pos = 0
        self.lock = threading.Lock()
        self.has_data = False

    def write(self, data):
        """Write audio data. Called from audio callback thread."""
        n = len(data)
        with self.lock:
            if n >= len(self.buffer):
                self.buffer[:] = data[-len(self.buffer):]
                self.write_pos = 0
            else:
                end = self.write_pos + n
                if end <= len(self.buffer):
                    self.buffer[self.write_pos:end] = data
                else:
                    first = len(self.buffer) - self.write_pos
                    self.buffer[self.write_pos:] = data[:first]
                    self.buffer[:n - first] = data[first:]
                self.write_pos = (self.write_pos + n) % len(self.buffer)
            self.has_data = True

    def read(self):
        """Read the full buffer in chronological order. Returns None if no data."""
        with self.lock:
            if not self.has_data:
                return None
            return np.roll(self.buffer, -self.write_pos).copy()

    def clear(self):
        """Zero out the buffer."""
        with self.lock:
            self.buffer[:] = 0
            self.write_pos = 0
            self.has_data = False


# ============================================
# ANALYSIS RESULT
# ============================================

class TunerResult:
    """Result of one analysis frame."""
    __slots__ = ['magnitudes', 'phase_offsets', 'cents_errors', 'active']

    def __init__(self):
        self.magnitudes = [0.0] * 12       # Energy per pitch class (0-1 normalized)
        self.phase_offsets = [0.0] * 12     # Accumulated rotation angle in degrees
        self.cents_errors = [0.0] * 12      # Current cents error from reference
        self.active = [False] * 12          # Whether wheel is "lit"


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
        self._last_time = None
        self._running = False
        self._window = None
        self._build_freq_table()

    def _build_freq_table(self):
        """Build reference frequency table for all pitch classes and octaves.

        freq_table[pc][oct_idx] gives the frequency for pitch class pc
        at octave (MIN_OCTAVE + oct_idx). pc 0 = C, pc 9 = A.
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

        For each pitch class:
        - Sums FFT magnitude across all octaves → magnitudes[pc]
        - Finds dominant frequency via parabolic interpolation
        - Computes cents error from reference → drives phase accumulation
        - Phase offset drives the stroboscopic rotation of that wheel
        """
        result = TunerResult()

        if not self._running or self._ring_buffer is None:
            return result

        audio = self._ring_buffer.read()
        if audio is None:
            return result

        return self.analyze_buffer(audio)

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
        noise_floor = np.median(mags[10:]) * 3.0 if len(mags) > 10 else 0.0
        sensitivity_scale = 1.0 + (100 - self._sensitivity) * 0.04  # 1.0 to 5.0
        threshold = noise_floor * sensitivity_scale

        max_mag = 0.0

        for pc in range(12):
            total_mag = 0.0
            best_mag = 0.0
            best_octave = -1
            best_bin = -1

            for oct_idx, freq in enumerate(self._freq_table[pc]):
                if freq < 25 or freq > SAMPLE_RATE / 2:
                    continue

                bin_idx = int(round(freq / bin_freq))
                if bin_idx < 1 or bin_idx >= len(mags) - 1:
                    continue

                # Peak magnitude at this bin and immediate neighbors
                mag = max(mags[bin_idx - 1], mags[bin_idx], mags[bin_idx + 1])

                if mag > threshold:
                    total_mag += mag
                    if mag > best_mag:
                        best_mag = mag
                        best_octave = oct_idx
                        best_bin = bin_idx

            result.magnitudes[pc] = total_mag
            if total_mag > max_mag:
                max_mag = total_mag

            # Phase tracking — only for pitch classes with sufficient energy
            if best_bin > 0 and best_mag > threshold:
                # Parabolic interpolation for sub-bin frequency accuracy
                alpha = float(mags[best_bin - 1])
                beta = float(mags[best_bin])
                gamma = float(mags[best_bin + 1])

                denom = alpha - 2 * beta + gamma
                if abs(denom) > 1e-10 and beta > 0:
                    p = 0.5 * (alpha - gamma) / denom
                    actual_freq = (best_bin + p) * bin_freq
                else:
                    actual_freq = best_bin * bin_freq

                ref_freq = self._freq_table[pc][best_octave]

                if actual_freq > 0 and ref_freq > 0:
                    cents = 1200.0 * math.log2(actual_freq / ref_freq)
                    cents = max(-100.0, min(100.0, cents))  # Clamp to ±1 semitone
                    result.cents_errors[pc] = cents

                    # Accumulate phase offset (the strobe rotation)
                    self._phase_offsets[pc] += cents * DRIFT_RATE * dt
                    self._phase_offsets[pc] %= 360.0

            result.phase_offsets[pc] = self._phase_offsets[pc]

        # Normalize magnitudes to 0-1
        if max_mag > 0:
            for pc in range(12):
                result.magnitudes[pc] /= max_mag

        # Determine active wheels (>5% of max magnitude)
        for pc in range(12):
            result.active[pc] = result.magnitudes[pc] > 0.05

        return result

    def reset_phases(self):
        """Reset all phase offsets to zero."""
        self._phase_offsets = [0.0] * 12


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
