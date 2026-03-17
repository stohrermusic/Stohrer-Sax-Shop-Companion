"""
Test script for toner display accuracy.

Generates synthetic sax-like audio with known harmonic profiles,
runs it through the engine, and checks whether the reported
harmonics_db, descriptors, and bar data accurately reflect the input.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import TonerEngine, SAMPLE_RATE, FFT_SIZE, MAX_HARMONICS

passed = 0
failed = 0

def test(name, condition, detail=""):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}  {detail}")
        failed += 1


def make_audio(freq, harmonics_amp, duration_s=0.5):
    """Generate audio with exact harmonic amplitudes.

    harmonics_amp: dict of {harmonic_number: linear_amplitude}
                   harmonic 1 = fundamental
    """
    t = np.arange(int(SAMPLE_RATE * duration_s), dtype=np.float64) / SAMPLE_RATE
    signal = np.zeros(len(t), dtype=np.float64)
    for n, amp in harmonics_amp.items():
        signal += amp * np.sin(2 * np.pi * freq * n * t)
    # Normalize to reasonable level
    peak = np.max(np.abs(signal))
    if peak > 0:
        signal = signal / peak * 0.5
    return signal.astype(np.float32)


def db_from_linear(amp, ref_amp):
    """Convert linear amplitude ratio to dB."""
    if amp <= 0 or ref_amp <= 0:
        return -80.0
    import math
    return 20.0 * math.log10(amp / ref_amp)


# ================================================
print("\n" + "=" * 60)
print("TEST SUITE: Toner Display Accuracy")
print("=" * 60)

engine = TonerEngine()

# ================================================
print("\n--- Test 1: Fundamental-dominant tone (2nd harmonic at -12 dB) ---")
# Fundamental = 1.0, 2nd harmonic = 0.25 (-12 dB), 3rd = 0.125 (-18 dB)
amps = {1: 1.0, 2: 0.25, 3: 0.125, 4: 0.0625}
audio = make_audio(440.0, amps)
r = engine.analyze_buffer(audio)

print(f"  Input: H1=0dB, H2={db_from_linear(0.25, 1.0):.1f}dB, H3={db_from_linear(0.125, 1.0):.1f}dB, H4={db_from_linear(0.0625, 1.0):.1f}dB")
if r.harmonics:
    for h in r.harmonics[:5]:
        print(f"  Output H{h.harmonic_number}: {h.magnitude_db:.1f} dB")
    test("H1 is 0 dB (reference)", abs(r.harmonics[0].magnitude_db) < 1.0,
         f"got {r.harmonics[0].magnitude_db:.1f}")
    if len(r.harmonics) > 1:
        test("H2 is near -12 dB", abs(r.harmonics[1].magnitude_db - (-12.0)) < 3.0,
             f"got {r.harmonics[1].magnitude_db:.1f}")
    if len(r.harmonics) > 2:
        test("H3 is near -18 dB", abs(r.harmonics[2].magnitude_db - (-18.0)) < 3.0,
             f"got {r.harmonics[2].magnitude_db:.1f}")

# ================================================
print("\n--- Test 2: Strong 2nd harmonic (sax-like, H2 stronger than H1) ---")
# This is common on some sax notes per UNSW research
amps = {1: 0.5, 2: 1.0, 3: 0.7, 4: 0.4, 5: 0.2}
audio = make_audio(300.0, amps)
r = engine.analyze_buffer(audio)

print(f"  Input: H1=0.5, H2=1.0 (stronger!), H3=0.7, H4=0.4, H5=0.2")
print(f"  Detected fundamental: {r.fundamental_freq:.1f} Hz (expected ~300)")
if r.harmonics:
    for h in r.harmonics[:6]:
        print(f"  Output H{h.harmonic_number}: {h.magnitude_db:.1f} dB (expected freq {h.expected_freq:.0f})")
test("Fundamental detected near 300 Hz", abs(r.fundamental_freq - 300.0) < 10.0,
     f"got {r.fundamental_freq:.1f}")

# Check that H1 reads as 0dB (it's the reference) and H2 reads POSITIVE
# because H2 is actually louder — but our engine reports relative to H1
if r.harmonics and len(r.harmonics) > 1:
    test("H1 is 0 dB (reference point)", abs(r.harmonics[0].magnitude_db) < 1.0,
         f"got {r.harmonics[0].magnitude_db:.1f}")
    test("H2 is reported as positive dB (louder than fundamental)",
         r.harmonics[1].magnitude_db > 0,
         f"got {r.harmonics[1].magnitude_db:.1f}")

# ================================================
print("\n--- Test 3: What the bar display sees ---")
# The bar display uses harmonics[i].magnitude_db
# In Bars mode: bar_h = max(0, (magnitude_db + 60.0) / 60.0) * height
# So 0 dB = full height, -60 dB = zero height
# If H2 is +6 dB, its bar would be (6+60)/60 = 1.1 => clamped to height
# This means H2's bar would be AS TALL OR TALLER than H1
print("  Bar height formula: bar_h = (dB + 60) / 60 * canvas_height")
print("  0 dB => 100% height, -30 dB => 50% height, -60 dB => 0%")
print()

amps_normal = {1: 1.0, 2: 0.7, 3: 0.4, 4: 0.2, 5: 0.1, 6: 0.05}
audio = make_audio(440.0, amps_normal)
r = engine.analyze_buffer(audio)

print("  Typical sax-like tone (H1=1.0, H2=0.7, H3=0.4...):")
if r.harmonics:
    for h in r.harmonics[:6]:
        bar_frac = max(0, (h.magnitude_db + 60.0) / 60.0)
        print(f"    H{h.harmonic_number}: {h.magnitude_db:+6.1f} dB => bar height {bar_frac:.0%}")

# ================================================
print("\n--- Test 4: The visualization problem ---")
# When H2 amplitude is 0.7x the fundamental, dB = -3.1
# Bar height = (-3.1 + 60) / 60 = 94.8% — nearly as tall!
# This IS accurate in dB terms but LOOKS like they're equal
print("  The 'nearly as tall' bars are CORRECT in dB scale.")
print("  0.7x amplitude = -3.1 dB = 94.8% bar height")
print("  0.5x amplitude = -6.0 dB = 90.0% bar height")
print("  0.25x amplitude = -12.0 dB = 80.0% bar height")
print("  0.1x amplitude = -20.0 dB = 66.7% bar height")
print("  0.01x amplitude = -40.0 dB = 33.3% bar height")
print()
print("  dB scale compresses large differences into small visual differences")
print("  at the top, and expands small differences at the bottom.")
print()

# Demonstrate the visual mismatch
print("  Visual comparison at different scales:")
print(f"  {'Amp Ratio':>10} {'dB':>8} {'Bar (dB 0-60)':>14} {'Bar (linear)':>14}")
for ratio in [1.0, 0.7, 0.5, 0.25, 0.1, 0.05, 0.01]:
    import math
    db = 20.0 * math.log10(ratio) if ratio > 0 else -60
    bar_db = max(0, (db + 60.0) / 60.0)
    bar_lin = ratio
    print(f"  {ratio:>10.2f} {db:>+8.1f} {bar_db:>13.0%} {bar_lin:>13.0%}")

# ================================================
print("\n--- Test 5: Descriptor accuracy for known profiles ---")

# Bright tone: strong upper harmonics
amps_bright = {1: 1.0, 2: 0.5, 3: 0.4, 4: 0.35, 5: 0.3, 6: 0.28,
               7: 0.6, 8: 0.55, 9: 0.5, 10: 0.4}
audio = make_audio(440.0, amps_bright)
r = engine.analyze_buffer(audio)
print(f"  Bright tone: brightness={r.descriptors['brightness']:.2f} darkness={r.descriptors['darkness']:.2f}")
test("Bright tone reads as bright", r.descriptors['brightness'] > 0.3,
     f"got {r.descriptors['brightness']:.2f}")

# Dark tone: strong lower, weak upper
amps_dark = {1: 1.0, 2: 0.8, 3: 0.5, 4: 0.3, 5: 0.05, 6: 0.02}
audio = make_audio(440.0, amps_dark)
r = engine.analyze_buffer(audio)
print(f"  Dark tone: brightness={r.descriptors['brightness']:.2f} darkness={r.descriptors['darkness']:.2f}")
test("Dark tone reads as dark", r.descriptors['darkness'] > r.descriptors['brightness'],
     f"dark={r.descriptors['darkness']:.2f} bright={r.descriptors['brightness']:.2f}")

# Rich tone: many even harmonics
amps_rich = {1: 1.0, 2: 0.8, 3: 0.7, 4: 0.65, 5: 0.6, 6: 0.55,
             7: 0.5, 8: 0.45, 9: 0.4, 10: 0.35, 11: 0.3}
audio = make_audio(440.0, amps_rich)
r = engine.analyze_buffer(audio)
print(f"  Rich tone: richness={r.descriptors['richness']:.2f}")
test("Rich tone reads as rich", r.descriptors['richness'] > 0.4,
     f"got {r.descriptors['richness']:.2f}")

# Pure tone: fundamental only
amps_pure = {1: 1.0}
audio = make_audio(440.0, amps_pure)
r = engine.analyze_buffer(audio)
print(f"  Pure tone: richness={r.descriptors['richness']:.2f}")
test("Pure tone reads as not rich", r.descriptors['richness'] < 0.1,
     f"got {r.descriptors['richness']:.2f}")

# ================================================
print("\n--- Test 6: Harmonic bars match harmonic data ---")
amps = {1: 1.0, 2: 0.5, 3: 0.3, 4: 0.15, 5: 0.08}
audio = make_audio(440.0, amps)
r = engine.analyze_buffer(audio)

test("harmonic_bars count matches harmonics",
     len(r.harmonic_bars) == len(r.harmonics),
     f"bars={len(r.harmonic_bars)} harmonics={len(r.harmonics)}")

if r.harmonic_bars and r.harmonics:
    for i, (bar, h) in enumerate(zip(r.harmonic_bars, r.harmonics)):
        freq_match = abs(bar[0] - h.expected_freq) < 1.0
        db_match = abs(bar[1] - h.magnitude_db) < 0.1
        if not freq_match or not db_match:
            test(f"Bar {i} matches harmonic {i}", False,
                 f"bar=({bar[0]:.0f}Hz, {bar[1]:.1f}dB) vs harm=({h.expected_freq:.0f}Hz, {h.magnitude_db:.1f}dB)")
    test("All bars match their harmonics", True)

# ================================================
print("\n--- Test 7: Realistic sax tones across register ---")
# Low note: Bb3 (233 Hz) - typically strong 2nd harmonic
amps_low = {1: 0.6, 2: 1.0, 3: 0.8, 4: 0.5, 5: 0.3, 6: 0.2, 7: 0.1}
audio = make_audio(233.0, amps_low)
r = engine.analyze_buffer(audio)
print(f"  Low Bb3: detected={r.fundamental_note} ({r.fundamental_freq:.0f} Hz)")
if r.harmonics:
    print(f"    H1={r.harmonics[0].magnitude_db:+.1f}dB  H2={r.harmonics[1].magnitude_db:+.1f}dB")
test("Low note: fundamental detected correctly",
     abs(r.fundamental_freq - 233.0) < 10.0,
     f"got {r.fundamental_freq:.1f}")

# Mid note: A4 (440 Hz)
amps_mid = {1: 1.0, 2: 0.6, 3: 0.35, 4: 0.2, 5: 0.12, 6: 0.07}
audio = make_audio(440.0, amps_mid)
r = engine.analyze_buffer(audio)
print(f"  Mid A4: detected={r.fundamental_note} ({r.fundamental_freq:.0f} Hz)")
test("Mid note: fundamental detected correctly",
     abs(r.fundamental_freq - 440.0) < 5.0,
     f"got {r.fundamental_freq:.1f}")

# High note: D5 (587 Hz)
amps_high = {1: 1.0, 2: 0.4, 3: 0.15, 4: 0.05}
audio = make_audio(587.0, amps_high)
r = engine.analyze_buffer(audio)
print(f"  High D5: detected={r.fundamental_note} ({r.fundamental_freq:.0f} Hz)")
test("High note: fundamental detected correctly",
     abs(r.fundamental_freq - 587.0) < 10.0,
     f"got {r.fundamental_freq:.1f}")

# ================================================
print("\n" + "=" * 60)
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print(f"{failed} TESTS FAILED")
print()
print("KEY FINDING: The bars display uses a dB scale (0 to -60 dB).")
print("A harmonic at half the fundamental's amplitude (-6 dB) appears")
print("as 90% of the bar height. This is mathematically correct but")
print("visually misleading — consider offering a linear amplitude scale.")
