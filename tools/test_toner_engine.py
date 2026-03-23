"""Test script for toner_engine.py — exercises tone analysis with synthetic audio."""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import TonerEngine, TonerResult, SAMPLE_RATE, FFT_SIZE

passed = 0
failed = 0

def test(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


def make_audio(freq, harmonics=None, duration_s=0.5, sr=SAMPLE_RATE):
    """Generate synthetic audio with specified fundamental and harmonics.

    harmonics: list of (harmonic_number, relative_amplitude)
    """
    t = np.arange(int(sr * duration_s), dtype=np.float64) / sr
    signal = np.sin(2 * np.pi * freq * t).astype(np.float32)
    if harmonics:
        for n, amp in harmonics:
            signal += amp * np.sin(2 * np.pi * freq * n * t).astype(np.float32)
    # Normalize
    peak = np.max(np.abs(signal))
    if peak > 0:
        signal = signal / peak * 0.5
    return signal


# ================================================
print("\n=== Test 1: Fundamental Detection (A4 = 440 Hz) ===")
engine = TonerEngine()
engine.set_sensitivity(50)
audio = make_audio(440.0)
result = engine.analyze_buffer(audio)
test("Detects fundamental", result.fundamental_freq > 0)
test("Fundamental near 440 Hz", abs(result.fundamental_freq - 440.0) < 5.0)
test("Note is A4", result.fundamental_note == "A4")
test("Cents near zero", abs(result.fundamental_cents) < 10.0)
test("Has harmonics list", len(result.harmonics) >= 1)
test("Has spectrum data", result.spectrum_db is not None)

# ================================================
print("\n=== Test 2: Fundamental Detection (Bb3 = 233.08 Hz, bari sax range) ===")
audio = make_audio(233.08)
result = engine.analyze_buffer(audio)
test("Detects fundamental", result.fundamental_freq > 0)
test("Fundamental near 233 Hz", abs(result.fundamental_freq - 233.08) < 5.0)
test("Note is A#3 or Bb3", "A#3" in result.fundamental_note or "Bb3" in result.fundamental_note)

# ================================================
print("\n=== Test 3: Rich Tone (many harmonics) ===")
harmonics = [(2, 0.8), (3, 0.6), (4, 0.5), (5, 0.4), (6, 0.3),
             (7, 0.25), (8, 0.2), (9, 0.15), (10, 0.1)]
audio = make_audio(440.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Detects fundamental at 440", abs(result.fundamental_freq - 440.0) < 5.0)
test("Finds multiple harmonics", len(result.harmonics) >= 6)
test("Richness is high", result.descriptors['richness'] > 0.5)

# ================================================
print("\n=== Test 4: Pure Tone (fundamental only) ===")
audio = make_audio(440.0)
result = engine.analyze_buffer(audio)
test("Richness is low", result.descriptors['richness'] < 0.3)

# ================================================
print("\n=== Test 5: Bright Tone (strong presence harmonics H3-H5) ===")
# Brightness is based on H2-H6 presence strength (H3-H5 weighted highest).
# Strong H3-H5 = bright/present, weak = dark/muted.
harmonics = [(2, 0.6), (3, 0.8), (4, 0.7), (5, 0.6), (6, 0.4)]
audio = make_audio(440.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Brightness is high", result.descriptors['brightness'] > 0.3)

# ================================================
print("\n=== Test 6: Dark Tone (weak upper harmonics) ===")
# Only fundamental with very weak H2 — presence harmonics are absent.
harmonics = [(2, 0.1)]
audio = make_audio(200.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Darkness is high", result.descriptors['darkness'] > 0.4)

# ================================================
print("\n=== Test 7: Resonance (harmonics at exact integer multiples) ===")
# Perfect harmonics = exact multiples, should be highly resonant
harmonics = [(2, 0.7), (3, 0.5), (4, 0.4), (5, 0.3)]
audio = make_audio(440.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Resonance is high for perfect harmonics", result.descriptors['resonance'] > 0.7)

# ================================================
print("\n=== Test 8: Harmonic bar data ===")
harmonics = [(2, 0.8), (3, 0.6)]
audio = make_audio(440.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Has harmonic_bars", len(result.harmonic_bars) >= 3)
if result.harmonic_bars:
    test("First bar near 440 Hz", abs(result.harmonic_bars[0][0] - 440.0) < 5.0)
    test("Second bar near 880 Hz", abs(result.harmonic_bars[1][0] - 880.0) < 10.0)

# ================================================
print("\n=== Test 9: Signal level ===")
audio = make_audio(440.0) * 0.001  # Very quiet
result = engine.analyze_buffer(audio)
test("Low signal level for quiet audio", result.signal_level < 0.1)

audio = make_audio(440.0)  # Normal level
result = engine.analyze_buffer(audio)
test("Higher signal level for normal audio", result.signal_level > 0.1)

# ================================================
print("\n=== Test 10: Silence produces empty result ===")
audio = np.zeros(FFT_SIZE * 2, dtype=np.float32)
result = engine.analyze_buffer(audio)
test("No fundamental on silence", result.fundamental_freq == 0.0)
test("No note name on silence", result.fundamental_note == "")

# ================================================
print("\n=== Test 11: Reference pitch affects note mapping ===")
engine.set_reference_pitch(432.0)
audio = make_audio(432.0)
result = engine.analyze_buffer(audio)
test("A4 at 432 Hz when ref is 432", result.fundamental_note == "A4")
test("Cents near zero at ref pitch", abs(result.fundamental_cents) < 10.0)
engine.set_reference_pitch(440.0)  # Reset

# ================================================
print("\n=== Test 12: High frequency fundamental (altissimo) ===")
audio = make_audio(1400.0)
result = engine.analyze_buffer(audio)
test("Detects high fundamental", result.fundamental_freq > 0)
test("Fundamental near 1400 Hz", abs(result.fundamental_freq - 1400.0) < 10.0)

# ================================================
print(f"\n{'='*50}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print(f"{failed} TESTS FAILED")
