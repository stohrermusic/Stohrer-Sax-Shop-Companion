"""The analyzer's pitch reading on generated tones: the octave, from silence, at every note.

Every note of soprano, alto, tenor, baritone and bass is synthesised from silence (a gap, then the
note) with its fundamental at 0, -6, -10 and -20 dB against the harmonic series, written to a WAV and
read back through the file path exactly as a guided run's recording is (`analyze_asked_notes`, each
note from its own stretch). Each must read its own octave within 0.5 cents. Until 2026-11-11 a note
whose second harmonic was the louder read an octave high for its whole length: the octave hysteresis
in `_detect_fundamental` held the first readings (made while the window was still filling, before the
fundamental could verify its own series) for ever, and returned the previous NUMBER rather than a
fresh measurement, so the pitch was frozen too (`docs/decisions/toner.md`, round 14).

Also here: the frozen-frequency case on the live path's shape; vibrato and drift bounds as measured;
noise; the bass floor (its four lowest notes lie under 65 Hz and have their own floor).

Generated tones only: no recording, no name. Runs in about a minute.
"""
import os
import sys
import tempfile

os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-test-pitch-")   # never the real settings
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import math
import wave

import numpy as np

import toner_engine as te

SR, N = te.SAMPLE_RATE, te.FFT_SIZE
NAMES = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
NOTE_S, GAP_S = 1.0, 0.3
FLOOR_DBFS = -60.0          # the quiet room under every generated tone
TMP = tempfile.mkdtemp(prefix="ssc-pitch-wav-")
rng = np.random.default_rng(20261111)

passed = failed = 0


def check(name, fn):
    global passed, failed
    try:
        r = fn()
        ok = r is not False
    except Exception as e:                    # noqa: BLE001  (a check fails on any exception)
        ok, r = False, f"{type(e).__name__}: {e}"
    if ok:
        passed += 1
        print(f"  PASS  {name}")
    else:
        failed += 1
        print(f"  FAIL  {name}: {r}")


def midi_name(m):
    return f"{NAMES[m % 12]}{m // 12 - 1}"


def tone(f0, cents, h1_db=0.0, vib_c=0.0, drift_c=0.0, snr_db=None, secs=NOTE_S):
    """A harmonic-rich tone: 20 partials falling 3 dB each, the fundamental moved by h1_db; vibrato at
    5.5 Hz, a linear drift, white noise at an SNR, and always a floor of noise at FLOOR_DBFS: a tone
    with no floor at all (16-bit quantisation alone, -96 dB) is not what a microphone gives, and on it
    the sub-harmonic check's thresholds, set against the window's noise median, let spectral leakage
    pass for a fundamental two octaves down. Returns the samples and the mean cents."""
    n = int(secs * SR)
    t = np.arange(n) / SR
    c = np.full(n, float(cents))
    if vib_c:
        c = c + vib_c * np.sin(2 * np.pi * 5.5 * t)
    if drift_c:
        c = c + drift_c * (t / secs - 0.5)
    phase = 2 * np.pi * np.cumsum(f0 * 2.0 ** (c / 1200.0)) / SR
    x = np.zeros(n)
    for h in range(1, 21):
        if h * f0 * 1.03 > SR / 2:
            break
        db = -3.0 * (h - 1) + (h1_db if h == 1 else 0.0)
        x += 10 ** (db / 20.0) * np.sin(h * phase + rng.uniform(0, 2 * np.pi))
    x *= 0.25 / np.max(np.abs(x))
    a = int(0.04 * SR)
    x[:a] *= np.linspace(0, 1, a)
    x[-a:] *= np.linspace(1, 0, a)
    if snr_db is not None:
        x = x + rng.normal(0, np.sqrt(np.mean(x ** 2) / 10 ** (snr_db / 10.0)), n)
    x = x + rng.normal(0, 10 ** (FLOOR_DBFS / 20.0), n)
    return x, float(np.mean(c))


def read_scale(horn, secs=NOTE_S, **kw):
    """Every note of the horn's range from silence, through the file path under the asked names.
    Returns [(midi, cents error after removing whole octaves, octaves off)]."""
    lo, hi = te.SAX_NOTE_RANGES[horn]
    parts, asked, truth, pos = [np.zeros(int(GAP_S * SR))], [], [], GAP_S
    for m in range(lo, hi + 1):
        x, true_c = tone(te.midi_to_hz(m), float(rng.uniform(-30, 30)), secs=secs, **kw)
        parts += [x, np.zeros(int(GAP_S * SR))]
        asked.append(dict(note=midi_name(m), written_note=midi_name(m), wav_start_s=pos, wav_end_s=pos + secs))
        truth.append((m, true_c))
        pos += secs + GAP_S
    path = os.path.join(TMP, f"{horn}.wav")
    with wave.open(path, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes((np.clip(np.concatenate(parts), -1, 1) * 32767).astype("<i2").tobytes())
    eng = te.TonerEngine()
    eng.set_sax_type(horn)
    caps = te.analyze_asked_notes(path, eng, asked)
    os.remove(path)
    out = []
    for (m, true_c), cap in zip(truth, caps):
        f = (cap or {}).get("fundamental_freq") or 0.0
        if f <= 0:
            out.append((m, float("nan"), None))
            continue
        e = 1200 * math.log2(f / te.midi_to_hz(m)) - true_c
        k = int(round(e / 1200.0))
        out.append((m, e - 1200 * k, k))
    return out


def worst(readings):
    """(notes an octave or more off, worst |cents| of the rest, the first few bad ones)."""
    off = [(midi_name(m), k) for m, e, k in readings if k is None or k != 0]
    cents = max([abs(e) for m, e, k in readings if k == 0] or [0.0])
    return off, cents


# ------------------------------------------------------------------ the octave, from silence
print("=== every note from silence, the fundamental at 0, -6, -10, -20 dB: its own octave within 0.5 c ===")
for horn in ("Soprano", "Alto", "Tenor", "Baritone"):
    for h1 in (0, -6, -10, -20):
        def gate(horn=horn, h1=h1):
            off, cents = worst(read_scale(horn, h1_db=h1))
            assert not off, f"octave errors at {off[:6]}{'...' if len(off) > 6 else ''} ({len(off)} notes)"
            assert cents < 0.5, f"worst {cents:.3f} c"
            return True
        check(f"{horn}, fundamental {h1:+d} dB: every note its own octave, within 0.5 c", gate)

# ------------------------------------------------------------------ the bass floor
print("\n=== Bass: its own floor (50 Hz) reads the four notes under 65 Hz too ===")
for h1 in (0, -10):
    def gate_bass(h1=h1):
        off, cents = worst(read_scale("Bass", h1_db=h1))
        assert not off, f"octave errors at {off[:6]} ({len(off)} notes)"
        assert cents < 0.5, f"worst {cents:.3f} c"
        return True
    check(f"Bass, fundamental {h1:+d} dB: all 33 notes, G#1 51.9 Hz up, their own octave within 0.5 c", gate_bass)


def t_floor_per_type():
    assert te.min_fundamental_hz_for("Bass") == te.MIN_FUNDAMENTAL_HZ_BASS == 50.0
    for s in te.SAX_TYPES:
        if s != "Bass":
            assert te.min_fundamental_hz_for(s) == te.MIN_FUNDAMENTAL_HZ == 65.0, s
    e = te.TonerEngine()
    assert e._min_fundamental_hz == 65.0
    e.set_sax_type("Bass")
    assert e._min_fundamental_hz == 50.0
    e.set_sax_type("Tenor")
    assert e._min_fundamental_hz == 65.0
    return True
check("the floor is per horn type: 50 Hz for Bass through set_sax_type, 65 Hz for every other", t_floor_per_type)

# ------------------------------------------------------------------ the frozen frequency
print("\n=== the live path's shape: a tone after silence is measured fresh in every window ===")


def t_frozen():
    """A 440 Hz tone with its second harmonic the louder, after a second of silence, fed window by
    window as the live path does: once the whole window holds the tone, every reading is 440 within
    0.5 c. Before the fix the first readings (the window half full) landed on 880 or on a frequency a
    tenth of a hertz off and the hold kept that NUMBER for the rest of the tone."""
    x, _c = tone(440.0, 0.0, h1_db=-6, secs=3.0)
    audio = np.concatenate([rng.normal(0, 10 ** (FLOOR_DBFS / 20.0), SR), x]).astype(np.float32)
    eng = te.TonerEngine()
    eng.set_sax_type("Alto")
    reads = []
    for end in range(N, len(audio), 1024):
        f = eng.analyze_buffer(audio[end - N:end]).fundamental_freq
        reads.append(((end - SR) / SR, f))                    # (seconds into the tone at the window's end, Hz)
    settled = [f for t, f in reads if t >= te.WINDOW_S + 0.1 and f > 0]
    assert settled, "no reading"
    errs = [1200 * math.log2(f / 440.0) for f in settled]
    assert max(abs(e) for e in errs) < 0.5, f"readings off 440: worst {max(errs, key=abs):+.2f} c; first {settled[:4]}"
    return True
check("a 440 Hz tone (H2 louder) after silence: every full-window reading 440 within 0.5 c, none frozen", t_frozen)


def t_hold_is_bounded():
    """The hold itself: a window whose own reading is an octave up and not the clearly stronger peak
    is held for at most OCTAVE_LATCH_WINDOWS windows; the hold is a measurement of the previous
    octave's peak in THIS window."""
    bf = SR / N
    e = te.TonerEngine()
    m = np.full(N // 2 + 1, 1e-4)
    for b, a in [(100, 1.0), (200, 0.9), (300, 0.5)]:          # a note at bin 100 with a strong H2
        m[b - 1] += 0.5 * a
        m[b] += a
        m[b + 1] += 0.5 * a
    f_note = e._detect_fundamental(m, bf)
    assert abs(f_note - 100 * bf) < bf, f_note
    # now the window's own candidate is H2 (bin 200, no peak at 100 any more but a shoulder of 0.7 x)
    m2 = np.full(N // 2 + 1, 1e-4)
    m2[199] += 0.45
    m2[200] += 0.9
    m2[201] += 0.45
    m2[99] += 0.35                  # the previous octave still there, not verified as a series
    m2[100] += 0.7
    m2[101] += 0.35
    reads = [e._detect_fundamental(m2, bf) for _ in range(te.OCTAVE_LATCH_WINDOWS + 2)]
    held = [r for r in reads if abs(r - 100 * bf) < bf]
    assert len(held) == te.OCTAVE_LATCH_WINDOWS - 1, (len(held), reads)
    assert abs(reads[-1] - 200 * bf) < bf, reads[-1]
    return True
check(f"an octave jump that is not the clearly stronger peak is held {getattr(te, 'OCTAVE_LATCH_WINDOWS', 0) - 1} windows, "
      "then the window's own reading wins", t_hold_is_bounded)

# ------------------------------------------------------------------ vibrato, drift, noise
print("\n=== vibrato, drift and noise: bounds as measured ===")
VIB_BOUND = {"Soprano": 10.0, "Alto": 1.5, "Tenor": 1.5, "Baritone": 1.5}
for horn, bound in VIB_BOUND.items():
    def gate_vib(horn=horn, bound=bound):
        off, cents = worst(read_scale(horn, vib_c=15, secs=2.2))
        assert not off, f"octave errors at {off[:6]}"
        assert cents < bound, f"worst {cents:.2f} c"
        return True
    check(f"{horn}, vibrato +/-15 c: its own octave, within {bound:g} c "
          f"{'(the peak of a vibrato wider than a bin reads at a turning point)' if horn == 'Soprano' else ''}", gate_vib)

for horn in ("Alto", "Tenor"):
    def gate_drift(horn=horn):
        off, cents = worst(read_scale(horn, drift_c=12, secs=2.2))
        assert not off, f"octave errors at {off[:6]}"
        assert cents < 0.5, f"worst {cents:.2f} c"
        return True
    check(f"{horn}, drift 12 c over the note: its own octave, the mean within 0.5 c", gate_drift)

for horn in ("Alto", "Baritone"):
    def gate_noise(horn=horn):
        off, cents = worst(read_scale(horn, h1_db=-10, snr_db=20))
        assert not off, f"octave errors at {off[:6]}"
        assert cents < 0.5, f"worst {cents:.2f} c"
        return True
    check(f"{horn}, fundamental -10 dB in noise at SNR 20: its own octave within 0.5 c", gate_noise)

print(f"\n{passed} passed, {failed} failed")
sys.exit(1 if failed else 0)
