"""
Shared audio utilities for Stohrer Sax Shop Companion.

Contains the AudioRingBuffer class and the Hann peak-frequency estimator used
by both tuner_engine.py and toner_engine.py. Pure math/threading — no tkinter or
sounddevice dependency.

Requires: numpy, threading (stdlib)
"""

import threading

try:
    import numpy as np
except ImportError:
    np = None


class AudioRingBuffer:
    """Thread-safe ring buffer for audio samples."""

    def __init__(self, size):
        self.buffer = np.zeros(size, dtype=np.float32)
        self.write_pos = 0
        self.lock = threading.Lock()
        self.has_data = False
        self.write_count = 0         # Increments on each write
        self.last_read_count = 0     # write_count at last read

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
            self.write_count += 1

    def read(self):
        """Read the full buffer in chronological order. Returns None if no data."""
        with self.lock:
            if not self.has_data:
                return None
            self.last_read_count = self.write_count
            return np.roll(self.buffer, -self.write_pos).copy()

    def is_stale(self):
        """True if no new data has been written since last read."""
        with self.lock:
            return self.write_count == self.last_read_count

    def clear(self):
        """Zero out the buffer."""
        with self.lock:
            self.buffer[:] = 0
            self.write_pos = 0
            self.has_data = False
            self.write_count = 0
            self.last_read_count = 0


def hann_peak_freq(mags, k, bin_freq):
    """Frequency (Hz) of a Hann-windowed spectral peak at bin k.

    A parabola through three linear magnitudes is the wrong shape for a
    Hann main lobe and lands up to ~0.05 bin off: ~8 cents at 110 Hz on
    the tuner's 10.77 Hz bins, ~2 cents on the toner's 2.69 Hz bins. For a
    Hann window the ratio of the larger neighbour to the peak fixes the
    offset exactly: |X[k+1]| / |X[k]| = (1 + d) / (2 - d), so
    d = (2a - 1) / (a + 1). Measured 2026-10-02 on tones with harmonics
    and -55 dBFS noise: worst 0.0012 bin, the same as a 4x zero-padded
    FFT with log-parabolic interpolation, at no extra FFT.

    The formula needs k to be the lobe's maximum, but callers pick k from
    a fixed window (the tuner looks at the 3 bins round the reference
    pitch), which misses the top when a high note is well off pitch: B5
    26 c flat peaks at bin 90 while the window offers 91-93. So climb to
    the local maximum first, at most 2 bins (the Hann main lobe's
    half-width), so the climb can't leave this peak's lobe.

    Only the frequency comes from here; callers read levels as before.

    Args:
        mags: |rfft| of the Hann-windowed frame
        k: bin at or near the peak, 0 < k < len(mags) - 1
        bin_freq: bin width in Hz
    """
    for _ in range(2):
        if k > 1 and mags[k - 1] > mags[k]:
            k -= 1
        elif k < len(mags) - 2 and mags[k + 1] > mags[k]:
            k += 1
        else:
            break
    peak = float(mags[k])
    if peak <= 0:
        return k * bin_freq
    left, right = float(mags[k - 1]), float(mags[k + 1])
    if right >= left:
        a = right / peak
        d = (2.0 * a - 1.0) / (a + 1.0)
    else:
        a = left / peak
        d = -(2.0 * a - 1.0) / (a + 1.0)
    d = max(-0.5, min(0.5, d))
    return (k + d) * bin_freq
