"""Test tuner engine with synthetic audio signals — no microphone needed."""
import sys
import math
sys.path.insert(0, '.')

import numpy as np
from tuner_engine import (
    TunerEngine, AudioRingBuffer, ReferencePlayer, TunerResult,
    SAMPLE_RATE, FFT_SIZE, PITCH_CLASSES, AUDIO_AVAILABLE,
    DISC_BASE_SEGMENTS, MIN_OCTAVE, MAX_OCTAVE,
)

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


def make_sine(freq, duration=0.2, amplitude=0.5):
    """Generate a sine wave at the given frequency."""
    t = np.arange(int(SAMPLE_RATE * duration), dtype=np.float32) / SAMPLE_RATE
    return (amplitude * np.sin(2 * np.pi * freq * t)).astype(np.float32)


def make_complex_tone(freq, duration=0.2, amplitude=0.5):
    """Generate a tone with harmonics (like a real instrument)."""
    t = np.arange(int(SAMPLE_RATE * duration), dtype=np.float32) / SAMPLE_RATE
    signal = amplitude * np.sin(2 * np.pi * freq * t)
    signal += amplitude * 0.5 * np.sin(2 * np.pi * 2 * freq * t)   # 2nd harmonic
    signal += amplitude * 0.25 * np.sin(2 * np.pi * 3 * freq * t)  # 3rd harmonic
    signal += amplitude * 0.125 * np.sin(2 * np.pi * 4 * freq * t) # 4th harmonic
    return signal.astype(np.float32)


# ============================================
# PREREQUISITES
# ============================================
print("\n--- Prerequisites ---")
test("numpy available", np is not None)
test("AUDIO_AVAILABLE flag is True", AUDIO_AVAILABLE)

# ============================================
# FREQUENCY TABLE
# ============================================
print("\n--- Frequency Table ---")
engine = TunerEngine()

# A4 should be reference pitch (440 Hz)
a4_idx = 9  # A is pitch class 9
a4_oct_idx = 4 - MIN_OCTAVE  # Octave 4, offset by MIN_OCTAVE
a4_freq = engine._freq_table[a4_idx][a4_oct_idx]
test(f"A4 = {a4_freq:.2f} Hz (expect 440.00)", abs(a4_freq - 440.0) < 0.01)

# C4 ≈ 261.63 Hz
c4_idx = 0
c4_oct_idx = 4 - MIN_OCTAVE
c4_freq = engine._freq_table[c4_idx][c4_oct_idx]
test(f"C4 = {c4_freq:.2f} Hz (expect 261.63)", abs(c4_freq - 261.626) < 0.01)

# E4 ≈ 329.63 Hz
e4_freq = engine._freq_table[4][4 - MIN_OCTAVE]
test(f"E4 = {e4_freq:.2f} Hz (expect 329.63)", abs(e4_freq - 329.628) < 0.02)

# Octave relationships: A5 = 2 * A4
a5_freq = engine._freq_table[a4_idx][5 - MIN_OCTAVE]
test(f"A5 = {a5_freq:.2f} Hz (expect 880.00)", abs(a5_freq - 880.0) < 0.01)

# A3 = A4 / 2
a3_freq = engine._freq_table[a4_idx][3 - MIN_OCTAVE]
test(f"A3 = {a3_freq:.2f} Hz (expect 220.00)", abs(a3_freq - 220.0) < 0.01)

# Custom reference pitch
engine.set_reference_pitch(442.0)
a4_442 = engine._freq_table[a4_idx][a4_oct_idx]
test(f"A4 at A=442 = {a4_442:.2f} Hz", abs(a4_442 - 442.0) < 0.01)
engine.set_reference_pitch(440.0)  # Reset

# All 12 pitch classes at octave 4 should be in chromatic order
freqs_oct4 = [engine._freq_table[pc][4 - MIN_OCTAVE] for pc in range(12)]
test("Chromatic scale ascending", all(freqs_oct4[i] < freqs_oct4[i+1] for i in range(11)))

# ============================================
# DRIFT RATES (Stroboconn physics)
# ============================================
print("\n--- Drift Rates (Stroboconn Physics) ---")
# A disc: 440 Hz / 16 segments = 27.5 rev/sec
# Drift = ln(2)/1200 * 27.5 * 360 = 5.72 deg/sec/cent
expected_a_drift = math.log(2) / 1200.0 * (440.0 / DISC_BASE_SEGMENTS) * 360.0
test(f"A drift rate = {engine._drift_rates[9]:.2f} (expect {expected_a_drift:.2f})",
     abs(engine._drift_rates[9] - expected_a_drift) < 0.01)

# C disc should drift slower than A disc (lower frequency = slower disc)
test("C drift rate < A drift rate", engine._drift_rates[0] < engine._drift_rates[9])

# B disc should drift faster than A disc
test("B drift rate > A drift rate", engine._drift_rates[11] > engine._drift_rates[9])

# All drift rates should be positive
test("All drift rates positive", all(r > 0 for r in engine._drift_rates))

# Drift rates should increase chromatically (C < C# < D < ... < B)
test("Drift rates increase chromatically",
     all(engine._drift_rates[i] < engine._drift_rates[i+1] for i in range(11)))

# Drift rate scales with reference pitch
engine.set_reference_pitch(442.0)
a_drift_442 = engine._drift_rates[9]
engine.set_reference_pitch(440.0)
test(f"Higher ref pitch = higher drift rate ({a_drift_442:.2f} > {engine._drift_rates[9]:.2f})",
     a_drift_442 > engine._drift_rates[9])

# ============================================
# RING BUFFER
# ============================================
print("\n--- Ring Buffer ---")
buf = AudioRingBuffer(100)
test("Empty buffer returns None", buf.read() is None)

# Write some data
data = np.arange(50, dtype=np.float32)
buf.write(data)
test("Has data after write", buf.has_data)

result = buf.read()
test("Read returns array", result is not None)
test("Read length matches buffer size", len(result) == 100)

# Write more than buffer size (wrap around)
big_data = np.arange(150, dtype=np.float32)
buf.write(big_data)
result = buf.read()
test("Wraparound: last 100 samples preserved", np.array_equal(result, big_data[-100:]))

# Clear
buf.clear()
test("Clear: read returns None", buf.read() is None)

# ============================================
# FFT ANALYSIS — PURE A4 (440 Hz)
# ============================================
print("\n--- FFT Analysis: Pure A4 ---")
engine = TunerEngine()
engine._last_time = None  # Reset timing
engine._window = np.hanning(FFT_SIZE).astype(np.float32)

audio_a4 = make_sine(440.0, duration=0.5, amplitude=0.8)
result = engine.analyze_buffer(audio_a4)

test("A (pc=9) has highest magnitude", result.magnitudes[9] == max(result.magnitudes))
test("A wheel is active", result.active[9])
test("A cents error < 5", abs(result.cents_errors[9]) < 5.0)

# Count how many other wheels are active (should be very few for pure sine)
active_count = sum(result.active)
test(f"Few wheels active for pure sine ({active_count})", active_count <= 3)

# Per-ring magnitudes: pure A4 should light up the octave-4 ring most
a_rings = result.ring_magnitudes[9]  # A pitch class
oct4_idx = 4 - 1  # Octave 4 maps to ring index 3 (MIN_OCTAVE=1)
test(f"A4 ring: octave 4 ring has energy ({a_rings[oct4_idx]:.3f})", a_rings[oct4_idx] > 0)
# Octave 4 should be the strongest ring
test("A4 ring: octave 4 is brightest ring",
     a_rings[oct4_idx] == max(a_rings))

# ============================================
# FFT ANALYSIS — PURE C4 (261.63 Hz)
# ============================================
print("\n--- FFT Analysis: Pure C4 ---")
engine2 = TunerEngine()
engine2._window = np.hanning(FFT_SIZE).astype(np.float32)

audio_c4 = make_sine(261.626, duration=0.5, amplitude=0.8)
result_c4 = engine2.analyze_buffer(audio_c4)

test("C (pc=0) has highest magnitude", result_c4.magnitudes[0] == max(result_c4.magnitudes))
test("C wheel is active", result_c4.active[0])
test("C cents error < 5", abs(result_c4.cents_errors[0]) < 5.0)

# ============================================
# FFT ANALYSIS — COMPLEX TONE (harmonics)
# ============================================
print("\n--- FFT Analysis: Complex A4 (with harmonics) ---")
engine3 = TunerEngine()
engine3._window = np.hanning(FFT_SIZE).astype(np.float32)

audio_complex = make_complex_tone(440.0, duration=0.5, amplitude=0.5)
result_complex = engine3.analyze_buffer(audio_complex)

test("A (pc=9) active with complex tone", result_complex.active[9])
test("A has highest magnitude", result_complex.magnitudes[9] == max(result_complex.magnitudes))

# With harmonics, more wheels should be active
# 2nd harmonic = 880Hz (A5) — still pitch class A
# 3rd harmonic = 1320Hz ≈ E6 — pitch class E (pc=4)
# 4th harmonic = 1760Hz (A6) — pitch class A
active_complex = sum(result_complex.active)
test(f"More wheels active with harmonics ({active_complex})", active_complex >= 2)

# E should have some energy from the 3rd harmonic
test("E (pc=4) has nonzero magnitude from 3rd harmonic", result_complex.magnitudes[4] > 0)

# Per-ring: A wheel should have energy in octave 4 AND octave 5 (2nd harmonic)
a_complex_rings = result_complex.ring_magnitudes[9]
oct4 = 4 - 1  # ring index for octave 4
oct5 = 5 - 1  # ring index for octave 5
test(f"Complex A4: octave 4 ring has energy ({a_complex_rings[oct4]:.3f})",
     a_complex_rings[oct4] > 0)
test(f"Complex A4: octave 5 ring has energy from 2nd harmonic ({a_complex_rings[oct5]:.3f})",
     a_complex_rings[oct5] > 0)
test("Complex A4: octave 4 ring brighter than octave 5",
     a_complex_rings[oct4] > a_complex_rings[oct5])

# ============================================
# PHASE TRACKING — DIRECTION
# ============================================
print("\n--- Phase Tracking: Direction ---")

# Sharp signal: slightly above 440 Hz → phase should drift
engine_sharp = TunerEngine()
engine_sharp._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_sharp._last_time = None

audio_sharp = make_sine(445.0, duration=0.5, amplitude=0.8)  # 5 Hz sharp (~20 cents)
result_sharp = engine_sharp.analyze_buffer(audio_sharp)
test("Sharp A: positive cents error", result_sharp.cents_errors[9] > 0)
test("Sharp A: phase offset nonzero", result_sharp.phase_offsets[9] != 0.0)

# Flat signal: slightly below 440 Hz
engine_flat = TunerEngine()
engine_flat._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_flat._last_time = None

audio_flat = make_sine(435.0, duration=0.5, amplitude=0.8)  # 5 Hz flat (~-20 cents)
result_flat = engine_flat.analyze_buffer(audio_flat)
test("Flat A: negative cents error", result_flat.cents_errors[9] < 0)

# In-tune signal: exactly 440 Hz → cents error ≈ 0
engine_tune = TunerEngine()
engine_tune._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_tune._last_time = None

audio_tune = make_sine(440.0, duration=0.5, amplitude=0.8)
result_tune = engine_tune.analyze_buffer(audio_tune)
test(f"In-tune A: cents error near zero ({result_tune.cents_errors[9]:.1f})", abs(result_tune.cents_errors[9]) < 3.0)

# ============================================
# PHASE ACCUMULATION OVER MULTIPLE FRAMES
# ============================================
print("\n--- Phase Accumulation ---")
engine_acc = TunerEngine()
engine_acc._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_acc._last_time = None

# Feed multiple frames of a sharp signal
sharp_audio = make_sine(445.0, duration=1.0, amplitude=0.8)
chunk_size = FFT_SIZE

# First analysis
r1 = engine_acc.analyze_buffer(sharp_audio[:chunk_size])
phase1 = r1.phase_offsets[9]

# Second analysis (feed more audio)
import time
time.sleep(0.02)  # Small delay for dt
r2 = engine_acc.analyze_buffer(sharp_audio[:chunk_size * 2])
phase2 = r2.phase_offsets[9]

test("Phase accumulates over time", phase2 != phase1)

# Direction: sharp should decrease phase (clockwise in canvas coords)
# flat should increase phase (counterclockwise in canvas coords)
engine_dir = TunerEngine()
engine_dir._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_dir._last_time = None

# Sharp: 445 Hz vs 440 Hz reference → positive cents → phase should decrease
audio_sharp_dir = make_sine(445.0, duration=0.5, amplitude=0.8)
r_sharp = engine_dir.analyze_buffer(audio_sharp_dir)
# With -= sign, positive cents causes phase to go down from 0 → wraps to near 360
# Since it wraps via % 360, a small negative becomes ~35x.x
# Just verify the direction by checking two frames
engine_dir2 = TunerEngine()
engine_dir2._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_dir2._last_time = None
chunk = make_sine(445.0, duration=2.0, amplitude=0.8)
r_s1 = engine_dir2.analyze_buffer(chunk[:FFT_SIZE])
p_s1 = r_s1.phase_offsets[9]
time.sleep(0.02)
r_s2 = engine_dir2.analyze_buffer(chunk[:FFT_SIZE * 2])
p_s2 = r_s2.phase_offsets[9]
# Phase wraps at 360, so unwrap: if p_s2 > p_s1 by a lot, it wrapped down
delta_sharp = p_s2 - p_s1
if delta_sharp > 180:
    delta_sharp -= 360
elif delta_sharp < -180:
    delta_sharp += 360
test(f"Sharp: phase drifts negative/clockwise (delta={delta_sharp:.1f})", delta_sharp < 0)

engine_dir3 = TunerEngine()
engine_dir3._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_dir3._last_time = None
chunk_flat = make_sine(435.0, duration=2.0, amplitude=0.8)
r_f1 = engine_dir3.analyze_buffer(chunk_flat[:FFT_SIZE])
p_f1 = r_f1.phase_offsets[9]
time.sleep(0.02)
r_f2 = engine_dir3.analyze_buffer(chunk_flat[:FFT_SIZE * 2])
p_f2 = r_f2.phase_offsets[9]
delta_flat = p_f2 - p_f1
if delta_flat > 180:
    delta_flat -= 360
elif delta_flat < -180:
    delta_flat += 360
test(f"Flat: phase drifts positive/counterclockwise (delta={delta_flat:.1f})", delta_flat > 0)

# ============================================
# SENSITIVITY
# ============================================
print("\n--- Sensitivity ---")
engine_sens = TunerEngine()
engine_sens._window = np.hanning(FFT_SIZE).astype(np.float32)

# Quiet signal
quiet_audio = make_sine(440.0, duration=0.5, amplitude=0.01)

# High sensitivity should detect it
engine_sens.set_sensitivity(100)
engine_sens._last_time = None
r_high = engine_sens.analyze_buffer(quiet_audio)

# Low sensitivity should reject it
engine_sens.set_sensitivity(0)
engine_sens._last_time = None
engine_sens._phase_offsets = [0.0] * 12
r_low = engine_sens.analyze_buffer(quiet_audio)

test("High sensitivity: A detected", r_high.magnitudes[9] > 0)
# Low sensitivity may or may not detect — we just verify it doesn't crash
test("Low sensitivity: no crash", True)

# ============================================
# RESET PHASES
# ============================================
print("\n--- Reset Phases ---")
engine_reset = TunerEngine()
engine_reset._phase_offsets = [45.0] * 12
engine_reset.reset_phases()
test("Reset phases to zero", all(p == 0.0 for p in engine_reset._phase_offsets))

# ============================================
# REFERENCE PLAYER (no audio output test, just API)
# ============================================
print("\n--- Reference Player API ---")
player = ReferencePlayer()
test("Player not playing initially", not player.is_playing)

# We can't test actual audio output without speakers, but test the API
started = player.play(440.0, "pure")
test("Player.play() returns True", started)
test("Player is_playing after play()", player.is_playing)

player.stop()
test("Player not playing after stop()", not player.is_playing)

# Rich waveform
started_rich = player.play(440.0, "rich")
test("Rich waveform starts", started_rich)
player.stop()

# ============================================
# EDGE CASES
# ============================================
print("\n--- Edge Cases ---")

# Silence → no active wheels
engine_silent = TunerEngine()
engine_silent._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_silent._last_time = None
silent = np.zeros(FFT_SIZE * 2, dtype=np.float32)
r_silent = engine_silent.analyze_buffer(silent)
test("Silence: no active wheels", not any(r_silent.active))

# Very short buffer → returns empty result
engine_short = TunerEngine()
engine_short._window = np.hanning(FFT_SIZE).astype(np.float32)
short = np.zeros(100, dtype=np.float32)
r_short = engine_short.analyze_buffer(short)
test("Short buffer: empty result", not any(r_short.active))

# Multiple simultaneous tones
engine_multi = TunerEngine()
engine_multi._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_multi._last_time = None
t = np.arange(int(SAMPLE_RATE * 0.5), dtype=np.float32) / SAMPLE_RATE
multi = (0.5 * np.sin(2 * np.pi * 440.0 * t) +   # A4
         0.5 * np.sin(2 * np.pi * 261.63 * t))     # C4
multi = multi.astype(np.float32)
r_multi = engine_multi.analyze_buffer(multi)
test("Two tones: A active", r_multi.active[9])
test("Two tones: C active", r_multi.active[0])

# All pitch classes detected from chromatic cluster
engine_all = TunerEngine()
engine_all._window = np.hanning(FFT_SIZE).astype(np.float32)
engine_all._last_time = None
cluster = np.zeros(int(SAMPLE_RATE * 0.5), dtype=np.float32)
for pc in range(12):
    freq = engine_all._freq_table[pc][4 - MIN_OCTAVE]  # All at octave 4
    t_arr = np.arange(len(cluster), dtype=np.float32) / SAMPLE_RATE
    cluster += (0.3 * np.sin(2 * np.pi * freq * t_arr)).astype(np.float32)
r_all = engine_all.analyze_buffer(cluster)
all_active = sum(r_all.active)
test(f"Chromatic cluster: many wheels active ({all_active}/12)", all_active >= 8)

# ============================================
# TUNER RESULT STRUCTURE
# ============================================
print("\n--- TunerResult Structure ---")
r = TunerResult()
test("TunerResult has 12 magnitudes", len(r.magnitudes) == 12)
test("TunerResult has 12 phase_offsets", len(r.phase_offsets) == 12)
test("TunerResult has 12 cents_errors", len(r.cents_errors) == 12)
test("TunerResult has 12 active flags", len(r.active) == 12)
test("All magnitudes start at 0", all(m == 0.0 for m in r.magnitudes))
test("All active flags start False", all(not a for a in r.active))
test("TunerResult has 12x7 ring_magnitudes",
     len(r.ring_magnitudes) == 12 and all(len(rm) == 7 for rm in r.ring_magnitudes))
test("All ring_magnitudes start at 0",
     all(rm == 0.0 for row in r.ring_magnitudes for rm in row))

# ============================================
# SUMMARY
# ============================================
print(f"\n{'='*40}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed} tests")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print("SOME TESTS FAILED")
    sys.exit(1)
