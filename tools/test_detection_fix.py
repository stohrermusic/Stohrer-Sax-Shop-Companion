#!/usr/bin/env python3
"""Test detection improvements for low register saxophone notes.

Verifies that _detect_fundamental correctly identifies the fundamental
when upper harmonics (3rd, 4th, 5th) are stronger, as is common in
low-register saxophone. Also tests the transposition fix for calibration
detected_as field consistency.
"""

import sys
sys.path.insert(0, '.')
import numpy as np
from toner_engine import TonerEngine, SAMPLE_RATE, FFT_SIZE, PITCH_CLASSES

passes = 0
fails = 0

def test(name, condition):
    global passes, fails
    if condition:
        passes += 1
        print(f"  PASS: {name}")
    else:
        fails += 1
        print(f"  FAIL: {name}")

def section(title):
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")

def note_to_freq(note_name):
    """Convert note name like 'G#2' to frequency."""
    if '#' in note_name:
        pc = note_name[:-1]
        octave = int(note_name[-1])
    else:
        pc = note_name[:-1]
        octave = int(note_name[-1])
    idx = PITCH_CLASSES.index(pc)
    midi = (octave + 1) * 12 + idx
    return 440.0 * 2 ** ((midi - 69) / 12)

def make_sax_tone(fundamental_hz, duration=0.4, harmonic_profile=None):
    """Generate a synthetic sax tone with realistic harmonic content.

    harmonic_profile: list of (harmonic_number, relative_db) tuples.
    If None, uses a generic sax profile.
    """
    t = np.arange(int(SAMPLE_RATE * duration)) / SAMPLE_RATE
    signal = np.zeros_like(t, dtype=np.float32)

    if harmonic_profile is None:
        # Generic sax profile: fundamental not necessarily strongest
        harmonic_profile = [
            (1, 0), (2, -3), (3, -6), (4, -10), (5, -12),
            (6, -15), (7, -18), (8, -20), (9, -22), (10, -25)
        ]

    for h_num, rel_db in harmonic_profile:
        freq = fundamental_hz * h_num
        if freq > SAMPLE_RATE / 2:
            continue
        amp = 10 ** (rel_db / 20)
        signal += amp * np.sin(2 * np.pi * freq * t)

    # Normalize
    peak = np.max(np.abs(signal))
    if peak > 0:
        signal /= peak
    return signal * 0.5

def detect_note(engine, signal):
    """Run detection on a signal and return (note_name, freq)."""
    # Pad signal to FFT_SIZE if needed
    if len(signal) < FFT_SIZE:
        signal = np.pad(signal, (0, FFT_SIZE - len(signal)))
    result = engine.analyze_buffer(signal)
    return result.fundamental_note, result.fundamental_freq

# ============================================================
section("Low register detection with strong upper harmonics")
# ============================================================
# These simulate the real problem: sax low register where H3-H5
# are stronger than H1 (the fundamental)

engine = TonerEngine()
engine._last_fundamental = 0  # Reset hysteresis

# G#2 (concert) = written A#3 on tenor = ~103.8 Hz
# Real data showed H5 strongest (freq ~510 Hz detected as C5)
print("\nG#2 (103.8 Hz) - 5th harmonic strongest (sax low register):")
signal = make_sax_tone(103.826, harmonic_profile=[
    (1, -15), (2, -10), (3, -8), (4, -6), (5, 0),  # H5 strongest
    (6, -5), (7, -8), (8, -10), (9, -12), (10, -15)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("G#2 detected correctly", note == "G#2" and 100 < freq < 108)

# A2 (concert) = written B3 on tenor = ~110 Hz
print("\nA2 (110 Hz) - 5th harmonic strongest:")
signal = make_sax_tone(110.0, harmonic_profile=[
    (1, -12), (2, -8), (3, -5), (4, -3), (5, 0),
    (6, -4), (7, -7), (8, -10), (9, -13), (10, -16)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("A2 detected correctly", note == "A2" and 107 < freq < 113)

# Bb2 (concert) = written C4 on tenor = ~116.5 Hz
print("\nA#2 (116.5 Hz) - 4th harmonic strongest:")
signal = make_sax_tone(116.541, harmonic_profile=[
    (1, -10), (2, -6), (3, -3), (4, 0), (5, -2),
    (6, -5), (7, -8), (8, -11), (9, -14), (10, -17)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("A#2 detected correctly", note == "A#2" and 113 < freq < 120)

# C3 (concert) = written D4 on tenor = ~130.8 Hz
# H2 strongest (octave confusion)
print("\nC3 (130.8 Hz) - 2nd harmonic strongest (octave error case):")
signal = make_sax_tone(130.813, harmonic_profile=[
    (1, -8), (2, 0), (3, -5), (4, -10), (5, -12),
    (6, -15), (7, -18), (8, -20)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("C3 detected correctly", note == "C3" and 127 < freq < 135)

# D3 (concert) = written E4 = ~146.8 Hz
# H4 strongest
print("\nD3 (146.8 Hz) - 4th harmonic strongest:")
signal = make_sax_tone(146.832, harmonic_profile=[
    (1, -12), (2, -5), (3, -3), (4, 0), (5, -4),
    (6, -8), (7, -12), (8, -15)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("D3 detected correctly", note == "D3" and 143 < freq < 150)

# ============================================================
section("Mid/high register (should still work)")
# ============================================================
# These should not be affected by the change — fundamental is dominant

# A4 = 440 Hz
print("\nA4 (440 Hz) - normal harmonic profile:")
signal = make_sax_tone(440.0)
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("A4 detected correctly", note == "A4" and 435 < freq < 445)

# D5 = 587.3 Hz
print("\nD5 (587.3 Hz) - normal harmonic profile:")
signal = make_sax_tone(587.33)
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("D5 detected correctly", note == "D5" and 583 < freq < 592)

# C6 = 1046.5 Hz
print("\nC6 (1046.5 Hz) - high register:")
signal = make_sax_tone(1046.5)
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("C6 detected correctly", note == "C6" and 1040 < freq < 1053)

# ============================================================
section("Edge case: very weak fundamental with strong H5")
# ============================================================
# Fundamental barely above noise, but H5 is 25 dB stronger

print("\nBb2 (116.5 Hz) - fundamental -25 dB below H5:")
signal = make_sax_tone(116.541, harmonic_profile=[
    (1, -25), (2, -15), (3, -10), (4, -5), (5, 0),
    (6, -3), (7, -6), (8, -10), (9, -14), (10, -18)
])
engine._last_fundamental = 0
note, freq = detect_note(engine, signal)
print(f"  Detected: {note} at {freq:.1f} Hz")
test("Bb2 with very weak fundamental", note == "A#2" and 113 < freq < 120)

# ============================================================
section("Transposition helper test")
# ============================================================
# Test that _toner_transpose_note logic works for common cases
# (We can't easily test the full UI method, but we can test the logic)

from toner_engine import SAX_TRANSPOSITIONS, PITCH_CLASSES

def transpose_note(concert_note, sax_type):
    """Standalone version of _toner_transpose_note for testing."""
    shift = SAX_TRANSPOSITIONS.get(sax_type, 0)
    if shift == 0 or not concert_note:
        return concert_note
    if '#' in concert_note:
        pc_name = concert_note[:-1]
        octave = int(concert_note[-1])
    else:
        pc_name = concert_note[:-1]
        octave = int(concert_note[-1])
    pc_idx = PITCH_CLASSES.index(pc_name)
    new_pc = (pc_idx + shift) % 12
    new_octave = octave + ((pc_idx + shift) // 12)
    return f"{PITCH_CLASSES[new_pc]}{new_octave}"

# Tenor: shift = 14 (Bb, sounds major 9th below written)
# Concert Bb2 = written C4 on tenor
written = transpose_note("C4", "Tenor")
print(f"\nConcert C4 -> Tenor written: {written}")
test("Concert C4 = Tenor written D5", written == "D5")

# Concert G#2 -> Tenor written A#3
written = transpose_note("G#2", "Tenor")
print(f"Concert G#2 -> Tenor written: {written}")
test("Concert G#2 = Tenor written A#3", written == "A#3")

# Concert Bb2 -> Tenor written C4
written = transpose_note("A#2", "Tenor")
print(f"Concert A#2 -> Tenor written: {written}")
test("Concert A#2 = Tenor written C4", written == "C4")

# Alto: shift = 9 (Eb instrument)
# Concert C4 = written A4 on alto
written = transpose_note("C4", "Alto")
print(f"Concert C4 -> Alto written: {written}")
test("Concert C4 = Alto written A4", written == "A4")

# ============================================================
section("Regression: pure tones still work")
# ============================================================

for note_name, freq_hz in [("A4", 440.0), ("E4", 329.63), ("B3", 246.94)]:
    t = np.arange(int(SAMPLE_RATE * 0.4)) / SAMPLE_RATE
    signal = (0.5 * np.sin(2 * np.pi * freq_hz * t)).astype(np.float32)
    engine._last_fundamental = 0
    note, freq = detect_note(engine, signal)
    print(f"\nPure {note_name} ({freq_hz} Hz): detected {note} at {freq:.1f} Hz")
    test(f"Pure {note_name} detected correctly", note == note_name)

# ============================================================
# Summary
# ============================================================
print(f"\n{'='*60}")
print(f"  RESULTS: {passes} passed, {fails} failed out of {passes + fails}")
print(f"{'='*60}")
sys.exit(0 if fails == 0 else 1)
