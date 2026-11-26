"""Test tuner engine with synthetic audio signals — no microphone needed."""
import os
import sys
import math
import time
sys.path.insert(0, '.')

import numpy as np
from audio_utils import hann_peak_freq
from tuner_engine import (
    TunerEngine, AudioRingBuffer, ReferencePlayer, TunerResult,
    SAMPLE_RATE, FFT_SIZE, AUDIO_AVAILABLE,
    DISC_BASE_SEGMENTS, MIN_OCTAVE,
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
try:
    import sounddevice as _sd_probe  # noqa: F401
    _sd_present = True
except Exception:  # noqa: BLE001
    _sd_present = False
test("AUDIO_AVAILABLE matches whether sounddevice is installed", AUDIO_AVAILABLE == _sd_present)

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
# DRIFT RATES (strobe disc physics)
# ============================================
print("\n--- Drift Rates (Strobe Disc Physics) ---")
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
test(f"Sharp: phase drifts positive/CW in shader (delta={delta_sharp:.1f})", delta_sharp > 0)

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
test(f"Flat: phase drifts negative/CCW in shader (delta={delta_flat:.1f})", delta_flat < 0)

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

# Opening an output stream needs an output device. CI runners (and some
# headless boxes) have none, and play() correctly returns False there —
# that's the fallback working, not a bug, so skip rather than fail.
#
# And PortAudio can block without bound on a machine with a strange audio
# device: on 2026-11-10 this suite, which takes 1.5 s, hit the runner's
# 900 s timeout on macOS. Every call that touches the audio device runs on
# a daemon thread with a time limit; past it the player checks are
# reported as skipped and the suite goes on (the stuck thread is left
# behind and the process ends with os._exit, which does not wait for it).
AUDIO_BOUND_S = 20.0
_audio_hung = []


def bounded(what, fn):
    """fn() within AUDIO_BOUND_S, or None and a SKIP line."""
    import threading
    box = {}

    def run():
        try:
            box['value'] = fn()
        except Exception as e:  # noqa: BLE001
            box['error'] = e
    t = threading.Thread(target=run, name=f"audio-{what}", daemon=True)
    t.start()
    t.join(AUDIO_BOUND_S)
    if t.is_alive():
        _audio_hung.append(what)
        print(f"  SKIP: {what} did not return in {AUDIO_BOUND_S:.0f} s (audio device hung); "
              "player checks skipped")
        return None
    if 'error' in box:
        raise box['error']
    return box['value']


def _probe_output():
    import sounddevice as _sd
    _sd.query_devices(kind='output')
    return True


try:
    _has_output = bounded("query_devices", _probe_output) is True
except Exception as _e:  # noqa: BLE001 — PortAudioError or no sounddevice
    _has_output = False
    print(f"  SKIP: no audio output device here ({type(_e).__name__}); "
          "player start tests skipped")
if _audio_hung:
    pass
elif _has_output:
    started = bounded("play", lambda: player.play(440.0, "pure"))
    if not _audio_hung:
        test("Player.play() returns True", started)
        test("Player is_playing after play()", player.is_playing)
else:
    started = bounded("play (no device)", lambda: player.play(440.0, "pure"))
    if not _audio_hung:
        test("Player.play() returns False with no output device", not started)

if not _audio_hung:
    bounded("stop", player.stop)
if not _audio_hung:
    test("Player not playing after stop()", not player.is_playing)

# Rich waveform
if _has_output and not _audio_hung:
    started_rich = bounded("play rich", lambda: player.play(440.0, "rich"))
    if not _audio_hung:
        test("Rich waveform starts", started_rich)
if not _audio_hung:
    bounded("stop", player.stop)

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
# PITCH READING ACCURACY (2026-10-02)
# ============================================
# A parabola through three linear magnitudes is the wrong shape for a Hann
# peak: an exactly-in-tune A2 read 7.8 c flat, so the strobe drifted on a
# note that was dead on. Peak frequency now comes from the Hann closed form
# (audio_utils.hann_peak_freq).
print("\n--- Pitch reading accuracy ---")


def make_tone_h2(freq, duration=0.2):
    """Sine plus a -6 dB 2nd harmonic."""
    t = np.arange(int(SAMPLE_RATE * duration)) / SAMPLE_RATE
    s = 0.5 * np.sin(2 * np.pi * freq * t) + 0.25 * np.sin(2 * np.pi * 2 * freq * t)
    return s.astype(np.float32)


# The estimator itself, on a pure Hann lobe at known sub-bin offsets.
_w = np.hanning(FFT_SIZE)
_t = np.arange(FFT_SIZE) / SAMPLE_RATE
_bin = SAMPLE_RATE / FFT_SIZE
worst_bin = 0.0
for d_true in (-0.45, -0.3, -0.1, 0.0, 0.1, 0.25, 0.4, 0.5):
    f = (40 + d_true) * _bin
    mags = np.abs(np.fft.rfft(np.sin(2 * np.pi * f * _t) * _w))
    worst_bin = max(worst_bin, abs(hann_peak_freq(mags, 40, _bin) / _bin - 40 - d_true))
test(f"Hann estimator on a pure Hann lobe within 0.01 bin (worst {worst_bin:.4f} bin)", worst_bin < 0.01)

OFFSETS = [-20, -10, -5, -2, 0, 2, 5, 10, 20]
worst, worst_at = 0.0, ""
for label, f in [("A2", 110.0), ("A3", 220.0), ("A4", 440.0), ("A5", 880.0)]:
    for off in OFFSETS:
        r = TunerEngine().analyze_buffer(make_tone_h2(f * 2 ** (off / 1200)))
        err = abs(r.cents_errors[9] - off)
        if err > worst:
            worst, worst_at = err, f"{label} {off:+d} c"
test(f"A2-A5 at -20..+20 c: cents_errors within 0.15 c (worst {worst:.3f} c, {worst_at})",
     worst < 0.15)

# The strobe itself: over one 0.1 s frame an in-tune note's wheel and ring
# must not turn by more than 0.15 c worth of drift.
worst_deg, worst_at = 0.0, ""
for label, f, octave in [("A2", 110.0, 2), ("A3", 220.0, 3), ("A4", 440.0, 4), ("A5", 880.0, 5)]:
    eng = TunerEngine()
    eng._last_time = time.perf_counter() - 1.0   # dt clamps to 0.1 s
    r = eng.analyze_buffer(make_tone_h2(f))
    for what, ph in [("wheel", r.phase_offsets[9]),
                     ("ring", r.ring_phase_offsets[9][octave - MIN_OCTAVE])]:
        turn = abs((ph + 180.0) % 360.0 - 180.0)
        if turn > worst_deg:
            worst_deg, worst_at = turn, f"{label} {what}"
limit_deg = 0.15 * eng._drift_rates[9] * 0.1
test(f"In-tune A2-A5: strobe turns < {limit_deg:.3f} deg per 0.1 s (worst {worst_deg:.4f}, {worst_at})",
     worst_deg < limit_deg)

# A high note well off pitch peaks outside the 3 bins round the reference
# bin (B5 26 c flat: peak at bin 90, window 91-93). The estimator climbs to
# the lobe's top instead of stopping half a bin short.
f_b5 = 440.0 * 2 ** ((83 - 69) / 12)
r = TunerEngine().analyze_buffer(make_tone_h2(f_b5 * 2 ** (-26.3 / 1200)))
test(f"B5 26.3 c flat reads -26.3 c (got {r.cents_errors[11]:+.3f})",
     abs(r.cents_errors[11] + 26.3) < 0.15)

# Which way it turns. A sharp note RAISES the wheel's phase and its ring's,
# a flat one lowers them: tuner_tab.window_phases and both renderers are
# built on this sign.
turns = {}
for off in (+10, -10):
    eng = TunerEngine()
    eng._last_time = time.perf_counter() - 1.0   # dt clamps to 0.1 s
    r = eng.analyze_buffer(make_tone_h2(440.0 * 2 ** (off / 1200)))
    turns[off] = [(ph + 180.0) % 360.0 - 180.0
                  for ph in (r.phase_offsets[9], r.ring_phase_offsets[9][4 - MIN_OCTAVE])]
test(f"A4 10 c sharp raises the wheel and ring phase ({turns[10][0]:+.2f}, {turns[10][1]:+.2f} deg per 0.1 s)",
     min(turns[10]) > 1.0)
test(f"A4 10 c flat lowers them ({turns[-10][0]:+.2f}, {turns[-10][1]:+.2f} deg per 0.1 s)",
     max(turns[-10]) < -1.0)


# ============================================
# A RING FOLLOWS ONLY ITS OWN PARTIAL (2026-11-25, from JustATuner)
# ============================================
# A ring was tracked on its own whenever anything at its bins cleared the
# threshold, which is the whole spectrum's median: in a real room the low
# rings sit in rumble far above that median, so under an A4 the rings A1, A2
# and A3 (no partial of the note there) were steered by noise and the disc
# did not look solid. A ring now follows its own pitch only for a real peak
# (RING_OWN_* in tuner_engine.py); anything else turns with the wheel. The
# room here is harsher than the one measured in JustATuner (low end 30x the
# median, not 16x).
print("\n--- A ring follows only its own partial ---")
_rng = np.random.default_rng(11)


def _room(seconds):
    n = int(SAMPLE_RATE * seconds)
    hiss = 10 ** (-62 / 20) * _rng.standard_normal(n)
    spec = np.fft.rfft(_rng.standard_normal(n))
    f = np.fft.rfftfreq(n, 1.0 / SAMPLE_RATE)
    spec[f > 350.0] = 0.0                          # rumble: everything under 350 Hz
    rumble = np.fft.irfft(spec, n)
    rumble *= 10 ** (-40 / 20) / np.sqrt(np.mean(rumble ** 2))
    return hiss + rumble


def _partials(freqs_amps, seconds, level):
    t = np.arange(int(SAMPLE_RATE * seconds)) / SAMPLE_RATE
    x = sum(a * np.sin(2 * np.pi * f * t + _rng.uniform(0, 6.28)) for f, a in freqs_amps)
    return level * x / np.max(np.abs(x))


def _ring_readings(x, pc=9):
    """Per ring: (fraction of frames it turned exactly with the wheel, mean
    cents it read, frames it went left, frames it went right)."""
    eng = TunerEngine()
    with_wheel = [0] * 7
    cents = [[] for _ in range(7)]
    prev, frames = None, 0
    hop_s = 1024 / SAMPLE_RATE                               # frames arrive at the hop's own pace
    for i, end in enumerate(range(FFT_SIZE, len(x), 1024)):
        eng._last_time = time.perf_counter() - hop_s
        r = eng.analyze_buffer(x[end - FFT_SIZE:end].astype(np.float32))
        cur = list(r.ring_phase_offsets[pc])
        if prev is not None and i >= 8:                      # past the 0.1 s a real partial waits
            frames += 1
            for k in range(7):
                with_wheel[k] += cur[k] == r.phase_offsets[pc]
                d = (cur[k] - prev[k] + 180.0) % 360.0 - 180.0
                cents[k].append(d / (eng._drift_rates[pc] * hop_s))
        prev = cur
    return [(with_wheel[k] / frames, float(np.mean(cents[k])),
             int(np.sum(np.array(cents[k]) < -0.5)), int(np.sum(np.array(cents[k]) > 0.5))) for k in range(7)]


_sharp = 2 ** (5 / 1200)
_harm = [(440.0 * _sharp * h, 1.0 / h) for h in range(1, 11)]
for _label, _level in (("at a normal level", 0.2), ("played quietly", 0.02)):
    _r = _ring_readings(_room(3.0) + _partials(_harm, 3.0, _level))
    _low = min(_r[k][0] for k in (0, 1, 2))
    test(f"A4 5 c sharp {_label}, in a rumbling room: the rings with no partial (A1 A2 A3) turn with the wheel "
         f"(worst ring: {100 * _low:.1f} % of frames)", _low >= 0.99)
    test(f"A4 5 c sharp {_label}: no ring goes left in any frame "
         f"(left/right per ring: {[(_r[k][2], _r[k][3]) for k in range(7)]})",
         all(_r[k][2] == 0 for k in range(7)))
    test(f"A4 5 c sharp {_label}: every ring reads +5 c (ring means {[round(_r[k][1], 2) for k in range(7)]})",
         all(abs(_r[k][1] - 5.0) < 0.3 for k in range(7)))

# The reason rings are independent at all: a partial that is not an exact
# multiple (a piano's are sharp) drifts on its own ring. Kept.
_inharm = [(440.0, 1.0), (880.0 * 2 ** (4 / 1200), 0.5), (1760.0 * 2 ** (9 / 1200), 0.25)]
_r = _ring_readings(_room(3.0) + _partials(_inharm, 3.0, 0.2))
test(f"A4 in tune with its octave partials 4 c and 9 c sharp: rings read {_r[3][1]:+.2f}, {_r[4][1]:+.2f}, "
     f"{_r[5][1]:+.2f}", abs(_r[3][1]) < 0.3 and abs(_r[4][1] - 4.0) < 0.3 and abs(_r[5][1] - 9.0) < 0.3)
test("...and the rings with no partial stand still with the wheel "
     f"({[round(_r[k][1], 2) for k in (0, 1, 2, 6)]} c)", all(abs(_r[k][1]) < 0.3 for k in (0, 1, 2, 6)))

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
if _audio_hung:
    # A thread is stuck inside PortAudio: a normal exit may wait on it for
    # ever. Say so, flush, and leave without it (the verdict is above).
    print(f"(audio device hung in: {', '.join(_audio_hung)}; exiting without waiting for it)")
    sys.stdout.flush()
    os._exit(1 if failed else 0)
if failed:
    sys.exit(1)
