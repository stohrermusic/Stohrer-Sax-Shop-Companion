"""Test script for toner_engine.py — exercises tone analysis with synthetic audio."""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import TonerEngine, SAMPLE_RATE, FFT_SIZE

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
print("\n=== Test 5: Warm Tone (strong H2 octave harmonic) ===")
harmonics = [(2, 0.9), (3, 0.5), (4, 0.3)]
audio = make_audio(440.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Warmth is high", result.descriptors['warmth'] > 0.3)

# ================================================
print("\n=== Test 6: Thin Tone (weak H2) ===")
harmonics = [(2, 0.05), (3, 0.3)]
audio = make_audio(200.0, harmonics=harmonics)
result = engine.analyze_buffer(audio)
test("Warmth is low", result.descriptors['warmth'] < 0.3)

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

# === Test 13: Warmth with varying H2 strength ===
print("\n=== Test 13: Warmth with varying H2 strength ===")
# Strong H2 = warm, weak H2 = thin
audio_warm = make_audio(440.0, harmonics=[(2, 0.9), (3, 0.5)])
audio_thin = make_audio(440.0, harmonics=[(2, 0.05), (3, 0.5)])
result_warm = engine.analyze_buffer(audio_warm)
result_thin = engine.analyze_buffer(audio_thin)
test("Strong H2 gives higher warmth", result_warm.descriptors['warmth'] > result_thin.descriptors['warmth'])

# === Test 14: low_harmonic_data flag ===
print("\n=== Test 14: low_harmonic_data flag ===")
# With strong harmonics — flag should be False
audio_good = make_audio(440.0, harmonics=[
    (2, 0.8), (3, 0.9), (4, 0.7), (5, 0.5), (6, 0.3),
])
result_good = engine.analyze_buffer(audio_good)
test("low_harmonic_data=False with full harmonics",
     result_good.descriptors.get('low_harmonic_data') is False)
# Pure tone (only fundamental) — fewer than 3 of H2-H6, flag should be True
audio_pure = make_audio(440.0)
result_pure = engine.analyze_buffer(audio_pure)
test("low_harmonic_data=True for pure tone (no H2-H6)",
     result_pure.descriptors.get('low_harmonic_data') is True)
# Just barely enough: H3 and H4 only (2 of H2-H6 = not enough, need 3)
audio_two = make_audio(440.0, harmonics=[(3, 0.5), (4, 0.5)])
result_two = engine.analyze_buffer(audio_two)
test("low_harmonic_data=True with only 2 of H2-H6",
     result_two.descriptors.get('low_harmonic_data') is True)
# Three of H2-H6 present — should be enough
audio_three = make_audio(440.0, harmonics=[(3, 0.8), (4, 0.7), (5, 0.6)])
result_three = engine.analyze_buffer(audio_three)
test("low_harmonic_data=False with 3 of H2-H6",
     result_three.descriptors.get('low_harmonic_data') is False)


# ================================================
print("\n=== Test 15: get_rolloff_threshold (mic + sax type aware) ===")
from toner_engine import get_rolloff_threshold, ROLLOFF_WARN_THRESHOLD

# Mic-only baseline (sax_type omitted) — backward compat
test("ribbon mic returns 3.5",
     get_rolloff_threshold("ribbon") == 3.5)
test("dynamic mic returns 2.8 with no sax type",
     get_rolloff_threshold("dynamic") == 2.8)
test("condenser mic returns base threshold",
     get_rolloff_threshold("condenser") == ROLLOFF_WARN_THRESHOLD)
test("unknown mic returns base threshold",
     get_rolloff_threshold("") == ROLLOFF_WARN_THRESHOLD)

# Alto + dynamic gets the bump (Foster Conn 6M data — 2.99 to 3.49)
test("alto + dynamic returns 3.5 (bumped for Foster Conn 6M)",
     get_rolloff_threshold("dynamic", "Alto") == 3.5)
test("alto + dynamic case-insensitive",
     get_rolloff_threshold("DYNAMIC", "alto") == 3.5)

# Other sax + dynamic: still 2.8 (Foster bari maxed at 2.31)
test("tenor + dynamic still 2.8",
     get_rolloff_threshold("dynamic", "Tenor") == 2.8)
test("baritone + dynamic still 2.8",
     get_rolloff_threshold("dynamic", "Baritone") == 2.8)
test("soprano + dynamic still 2.8",
     get_rolloff_threshold("dynamic", "Soprano") == 2.8)

# Sax type doesn't affect non-dynamic mics (yet)
test("alto + condenser unchanged",
     get_rolloff_threshold("condenser", "Alto") == ROLLOFF_WARN_THRESHOLD)
test("alto + ribbon unchanged",
     get_rolloff_threshold("ribbon", "Alto") == 3.5)

# None sax_type behaves like backward compat
test("None sax_type returns base mic threshold",
     get_rolloff_threshold("dynamic", None) == 2.8)


# ================================================
# Pitch reading accuracy and octave errors (2026-10-02)
# ================================================
# Peak frequencies come from the Hann closed form (audio_utils.hann_peak_freq);
# the old parabola through linear magnitudes read up to ~4 c off. Levels
# (harmonics_db) are still read the old way so stored captures stay comparable.
print("\n=== Test 20: Pitch reading accuracy ===")
from toner_engine import SAX_NOTE_RANGES  # noqa: E402

TRUE_DB = [0, -3, -6, -10, -14, -18]
t = np.arange(int(SAMPLE_RATE * 0.5)) / SAMPLE_RATE
tone = sum(10 ** (db / 20) * np.sin(2 * np.pi * 233.08 * n * t) for n, db in enumerate(TRUE_DB, 1))
tone = (0.3 * tone / np.max(np.abs(tone))).astype(np.float32)
eng = TonerEngine()
eng.set_sensitivity(50)
r = eng.analyze_buffer(tone)
f_err = 1200 * np.log2(r.fundamental_freq / 233.08) if r.fundamental_freq > 0 else 999.0
test(f"233.08 Hz reads within 0.05 c (got {f_err:+.3f} c)", abs(f_err) < 0.05)
lvl = {h.harmonic_number: h.magnitude_db for h in r.harmonics}
db_err = max(abs(lvl.get(n, -99.0) - db) for n, db in enumerate(TRUE_DB, 1))
test(f"233.08 Hz harmonics H1-H6 within 1 dB (worst {db_err:.2f} dB)", db_err < 1.0)
c_err = max(abs(h.cents_deviation) for h in r.harmonics if h.harmonic_number <= len(TRUE_DB))
test(f"233.08 Hz harmonic_cents H1-H6 within 0.1 c (worst {c_err:.3f} c)", c_err < 0.1)


def _detect(mags):
    """Run _detect_fundamental on a hand-built spectrum, fresh engine."""
    e = TonerEngine()
    return e._detect_fundamental(mags, SAMPLE_RATE / FFT_SIZE)


def _peak(mags, b, amp):
    """A Hann-shaped peak centred on bin b."""
    mags[b - 1] += 0.5 * amp
    mags[b] += amp
    mags[b + 1] += 0.5 * amp


print("\n=== Test 21: Octave errors in fundamental detection ===")
BIN = SAMPLE_RATE / FFT_SIZE
# A noise bump at half a note must not become its fundamental. The /2 check
# used to count x2 (the strongest peak itself) and x4 (the note's own H2),
# so the bump passed on evidence the real note supplies.
m = np.full(FFT_SIZE // 2 + 1, 1e-4)
for b, a in [(200, 1.0), (400, 0.5), (600, 0.3)]:
    _peak(m, b, a)
_peak(m, 100, 0.05)
f0 = _detect(m)
test(f"Bump at half a note is not its fundamental (read {f0:.1f} Hz, note {200 * BIN:.1f})",
     abs(f0 - 200 * BIN) < BIN)

# The peak search round a sub-harmonic can land under MIN_FUNDAMENTAL_HZ;
# a fundamental below the floor is not one.
m = np.full(FFT_SIZE // 2 + 1, 1e-4)
_peak(m, 121, 1.0)
for b in (22, 66, 88, 132):         # bin 22 = 59 Hz, under the floor, plus its x3 x4 x6
    _peak(m, b, 0.05)
f0 = _detect(m)
test(f"No fundamental under the 65 Hz floor (read {f0:.1f} Hz, note {121 * BIN:.1f})",
     abs(f0 - 121 * BIN) < BIN)

# The tone that the narrower fix (skip only mult == divisor) read an octave
# HIGH: H2 a little over H1, weak H3, no H4. The /2 check must accept it on H3.
for f in (207.65, 329.63, 523.25):
    tt = np.arange(int(SAMPLE_RATE * 0.5)) / SAMPLE_RATE
    x = (0.7 * np.sin(2 * np.pi * f * tt) + np.sin(2 * np.pi * 2 * f * tt)
         + 0.3 * np.sin(2 * np.pi * 3 * f * tt)).astype(np.float32) * 0.3
    e = TonerEngine()
    e.set_sensitivity(50)
    got = e.analyze_buffer(x).fundamental_freq
    test(f"H2-over-H1 tone at {f} Hz reads {f} (got {got:.1f})", abs(got - f) < 2.0)

# Sweep: every concert note of soprano/alto/tenor/bari, rich and thin tones,
# noise at -55 dBFS, +-25 c detune, one engine per note fed growing buffers.
RECIPES = {
    "bright": [0, -3, -6, -9, -12, -15, -18, -21, -24, -27],
    "weakH1": [-20, 0, -4, -8, -12, -16, -20],   # fundamental 20 dB under H2
    "h2over": [-3, 0, -10],
    "thin2": [0, -2],
}
rng = np.random.default_rng(11)
frames = low = high = 0
worst_c = 0.0
for sax in ["Soprano", "Alto", "Tenor", "Baritone"]:
    lo_m, hi_m = SAX_NOTE_RANGES[sax]
    for midi in range(lo_m, hi_m + 1):
        for rec in RECIPES.values():
            f = 440.0 * 2 ** ((midi - 69 + rng.uniform(-0.25, 0.25)) / 12)
            tt = np.arange(int(SAMPLE_RATE * 0.45)) / SAMPLE_RATE
            x = sum(10 ** (db / 20) * np.sin(2 * np.pi * f * n * tt + rng.uniform(0, 2 * np.pi))
                    for n, db in enumerate(rec, 1) if f * n < SAMPLE_RATE / 2 - 200)
            x = (0.3 * x / np.max(np.abs(x)) + rng.normal(0, 10 ** (-55 / 20), len(tt))).astype(np.float32)
            e = TonerEngine()
            e.set_sensitivity(50)
            for n_buf in range(FFT_SIZE, len(x) + 1, 1024):
                got = e.analyze_buffer(x[:n_buf]).fundamental_freq
                frames += 1
                c = 1200 * np.log2(got / f) if got > 0 else -9999.0
                if c < -50:
                    low += 1
                elif c > 50:
                    high += 1
                else:
                    worst_c = max(worst_c, abs(c))
test(f"Sweep: no frame an octave high ({high} of {frames})", high == 0)
test(f"Sweep: no frame an octave low ({low} of {frames})", low == 0)
test(f"Sweep: every right-octave frame within 0.5 c (worst {worst_c:.3f} c)", worst_c < 0.5)


# ================================================
print(f"\n{'='*50}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print(f"{failed} TESTS FAILED")

sys.exit(0 if failed == 0 else 1)
