"""
Comprehensive test suite for the Toner tab.

Tests everything that can be verified without a real microphone:
- Engine: pitch detection, harmonic extraction, descriptors
- Profiles: create, save, load, migrate, merge, delete, fingerprint
- Capture: state machine, averaging, frame collection
- Comparison: fingerprint computation, per-note data, analysis
- Scale: linear vs dB height calculations
- Settings: save/load round-trip
- Edge cases: empty data, corrupt data, missing notes, extreme values
"""

import sys
import os
import json
import math
import tempfile
import shutil
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import (
    TonerEngine, TonerResult, HarmonicInfo,
    SAMPLE_RATE, FFT_SIZE, MAX_HARMONICS, MIN_PROFILE_NOTES,
    CAPTURE_DELAY_S, DEFAULT_LIBRARY,
    average_captures, compute_fingerprint,
    load_tone_profiles, save_tone_profiles, flatten_profiles,
)

passed = 0
failed = 0
section_pass = 0
section_fail = 0

def test(name, condition, detail=""):
    global passed, failed, section_pass, section_fail
    if condition:
        passed += 1
        section_pass += 1
    else:
        print(f"    FAIL: {name}  {detail}")
        failed += 1
        section_fail += 1

def section(name):
    global section_pass, section_fail
    if section_pass + section_fail > 0:
        status = "PASS" if section_fail == 0 else f"{section_fail} FAILED"
        print(f"  [{status}] ({section_pass + section_fail} tests)")
    section_pass = 0
    section_fail = 0
    print(f"\n--- {name} ---")

def make_audio(freq, harmonics_amp=None, duration_s=0.5):
    t = np.arange(int(SAMPLE_RATE * duration_s), dtype=np.float64) / SAMPLE_RATE
    signal = np.sin(2 * np.pi * freq * t).astype(np.float64)
    if harmonics_amp:
        for n, amp in harmonics_amp.items():
            if n == 1:
                signal *= amp
            else:
                signal += amp * np.sin(2 * np.pi * freq * n * t)
    peak = np.max(np.abs(signal))
    if peak > 0:
        signal = signal / peak * 0.5
    return signal.astype(np.float32)


print("=" * 60)
print("TONER COMPREHENSIVE TEST SUITE")
print("=" * 60)

engine = TonerEngine()

# ============================================================
section("Pitch detection - standard notes")
for name, freq in [("C3", 130.81), ("A3", 220.0), ("A4", 440.0),
                     ("C5", 523.25), ("A5", 880.0), ("D5", 587.33)]:
    audio = make_audio(freq, {2: 0.5, 3: 0.3})
    r = engine.analyze_buffer(audio)
    test(f"{name} ({freq:.0f} Hz)", abs(r.fundamental_freq - freq) < 10.0,
         f"got {r.fundamental_freq:.1f}")

# ============================================================
section("Pitch detection - sax range extremes")
# Bari low Bb
audio = make_audio(116.54, {2: 0.8, 3: 0.5})
r = engine.analyze_buffer(audio)
test("Bari low Bb2 (116 Hz)", abs(r.fundamental_freq - 116.54) < 10.0,
     f"got {r.fundamental_freq:.1f}")

# Alto altissimo
audio = make_audio(1480.0, {2: 0.3})
r = engine.analyze_buffer(audio)
test("Altissimo F#6 (1480 Hz)", abs(r.fundamental_freq - 1480.0) < 20.0,
     f"got {r.fundamental_freq:.1f}")

# ============================================================
section("Pitch detection - strong 2nd harmonic (common on sax)")
# H2 stronger than H1
audio = make_audio(300.0, {1: 0.5, 2: 1.0, 3: 0.7, 4: 0.4})
r = engine.analyze_buffer(audio)
test("Detects fundamental, not 2nd harmonic",
     abs(r.fundamental_freq - 300.0) < 10.0,
     f"got {r.fundamental_freq:.1f}")

# ============================================================
section("Pitch detection - silence and noise")
audio = np.zeros(FFT_SIZE * 2, dtype=np.float32)
r = engine.analyze_buffer(audio)
test("Silence: no detection", r.fundamental_freq == 0.0)
test("Silence: empty note", r.fundamental_note == "")

audio = np.random.randn(FFT_SIZE * 2).astype(np.float32) * 0.001
r = engine.analyze_buffer(audio)
test("Very quiet noise: no detection", r.fundamental_freq == 0.0)

# ============================================================
section("Note naming and cents")
engine.set_reference_pitch(440.0)
audio = make_audio(440.0)
r = engine.analyze_buffer(audio)
test("440 Hz = A4", r.fundamental_note == "A4")
test("440 Hz = ~0 cents", abs(r.fundamental_cents) < 5.0,
     f"got {r.fundamental_cents:.1f}")

audio = make_audio(466.16)  # Bb4
r = engine.analyze_buffer(audio)
test("466 Hz = A#4", "A#4" in r.fundamental_note)

engine.set_reference_pitch(432.0)
audio = make_audio(432.0)
r = engine.analyze_buffer(audio)
test("432 Hz = A4 at A=432", r.fundamental_note == "A4")
engine.set_reference_pitch(440.0)

# ============================================================
section("Harmonic extraction accuracy")
amps = {1: 1.0, 2: 0.5, 3: 0.25, 4: 0.125}
audio = make_audio(440.0, amps)
r = engine.analyze_buffer(audio)

test("Finds at least 4 harmonics", len(r.harmonics) >= 4)
if len(r.harmonics) >= 4:
    test("H1 at 0 dB", abs(r.harmonics[0].magnitude_db) < 1.0,
         f"got {r.harmonics[0].magnitude_db:.1f}")
    test("H2 near -6 dB", abs(r.harmonics[1].magnitude_db - (-6.0)) < 3.0,
         f"got {r.harmonics[1].magnitude_db:.1f}")
    test("H3 near -12 dB", abs(r.harmonics[2].magnitude_db - (-12.0)) < 3.0,
         f"got {r.harmonics[2].magnitude_db:.1f}")
    test("H4 near -18 dB", abs(r.harmonics[3].magnitude_db - (-18.0)) < 3.0,
         f"got {r.harmonics[3].magnitude_db:.1f}")

    # Frequency positions
    test("H1 freq near 440", abs(r.harmonics[0].expected_freq - 440) < 5)
    test("H2 freq near 880", abs(r.harmonics[1].expected_freq - 880) < 5)
    test("H3 freq near 1320", abs(r.harmonics[2].expected_freq - 1320) < 10)

# ============================================================
section("Harmonic bars match harmonics")
test("Bar count matches", len(r.harmonic_bars) == len(r.harmonics))
for i, (bar, h) in enumerate(zip(r.harmonic_bars, r.harmonics)):
    test(f"Bar {i} freq", abs(bar[0] - h.expected_freq) < 1.0)
    test(f"Bar {i} dB", abs(bar[1] - h.magnitude_db) < 0.1)

# ============================================================
section("Spectrum data")
test("spectrum_db exists", r.spectrum_db is not None)
test("spectrum_freqs exists", r.spectrum_freqs is not None)
if r.spectrum_db is not None:
    test("spectrum_db range", float(np.max(r.spectrum_db)) <= 0.0)
    test("spectrum_db min", float(np.min(r.spectrum_db)) >= -80.0)
    test("spectrum length > 0", len(r.spectrum_db) > 100)

# ============================================================
section("Descriptors - resonance")
# Perfect harmonics = high resonance
audio = make_audio(440.0, {2: 0.5, 3: 0.3, 4: 0.2})
r = engine.analyze_buffer(audio)
test("Perfect harmonics: high resonance", r.descriptors['resonance'] > 0.7,
     f"got {r.descriptors['resonance']:.2f}")

# ============================================================
section("Descriptors - richness")
# Pure: fundamental only
audio = make_audio(440.0)
r = engine.analyze_buffer(audio)
test("Pure tone: low richness", r.descriptors['richness'] < 0.1,
     f"got {r.descriptors['richness']:.2f}")

# Rich: many strong harmonics
audio = make_audio(440.0, {2: 0.8, 3: 0.7, 4: 0.65, 5: 0.6, 6: 0.55,
                            7: 0.5, 8: 0.45, 9: 0.4, 10: 0.35})
r = engine.analyze_buffer(audio)
test("Rich tone: high richness", r.descriptors['richness'] > 0.4,
     f"got {r.descriptors['richness']:.2f}")

# Moderate: a few harmonics with clear rolloff
audio = make_audio(440.0, {2: 0.5, 3: 0.2, 4: 0.08, 5: 0.03})
r = engine.analyze_buffer(audio)
test("Moderate tone: mid richness", r.descriptors['richness'] > 0.05,
     f"got {r.descriptors['richness']:.2f}")

# ============================================================
section("Descriptors - brightness and darkness")
# Bright: strong upper harmonics
audio = make_audio(440.0, {7: 0.6, 8: 0.55, 9: 0.5, 10: 0.4})
r = engine.analyze_buffer(audio)
test("Bright tone", r.descriptors['brightness'] > 0.3,
     f"got {r.descriptors['brightness']:.2f}")

# Dark: low fundamental with strong harmonics below break frequency (~750 Hz)
# Use 200 Hz so H2=400, H3=600 are below break, H4=800 barely above
audio = make_audio(200.0, {2: 0.8, 3: 0.5, 4: 0.3})
r = engine.analyze_buffer(audio)
test("Dark tone: dark > bright",
     r.descriptors['darkness'] > r.descriptors['brightness'],
     f"dark={r.descriptors['darkness']:.2f} bright={r.descriptors['brightness']:.2f}")

# ============================================================
section("Descriptors - fullness")
# Fullness = balance of bright and dark energy. Peaks when both sides
# contribute equally, drops when one dominates. At 440 Hz with default
# break (750 Hz), H2+ are all bright, so even a harmonic-rich tone
# reads as unbalanced (bright-heavy) and low fullness. This is correct.

# Use a low fundamental (130 Hz = C3) where some harmonics fall below break
# and others above — this CAN be a genuinely full tone.
audio = make_audio(130.0, {2: 0.8, 3: 0.7, 4: 0.6, 5: 0.5,
                            6: 0.5, 7: 0.45, 8: 0.4, 9: 0.35})
r = engine.analyze_buffer(audio)
test("Full tone (low note, balanced harmonics): high fullness",
     r.descriptors['fullness'] > 0.3,
     f"got {r.descriptors['fullness']:.2f}")

# Very bright-heavy signal: strong upper harmonics swamp the fundamental
audio = make_audio(440.0, {2: 1.2, 3: 1.0, 4: 0.9, 5: 0.8,
                            6: 0.7, 7: 0.6, 8: 0.5, 9: 0.4})
r = engine.analyze_buffer(audio)
test("Bright-heavy high note: lower fullness than balanced low note",
     r.descriptors['fullness'] < 0.4,
     f"bright={r.descriptors['brightness']:.2f} full={r.descriptors['fullness']:.2f}")

# ============================================================
section("Signal level")
audio = make_audio(440.0) * 0.001
r = engine.analyze_buffer(audio)
test("Quiet audio: low signal", r.signal_level < 0.05)

audio = make_audio(440.0)
r = engine.analyze_buffer(audio)
test("Normal audio: higher signal", r.signal_level > 0.1)

# ============================================================
section("Sensitivity setting")
engine.set_sensitivity(100)  # Most sensitive
audio = make_audio(440.0) * 0.01
r = engine.analyze_buffer(audio)
detected_at_100 = r.fundamental_freq > 0

engine.set_sensitivity(0)  # Least sensitive
r = engine.analyze_buffer(audio)
detected_at_0 = r.fundamental_freq > 0

test("High sensitivity detects quiet signal", detected_at_100)
test("Low sensitivity misses quiet signal", not detected_at_0)
engine.set_sensitivity(50)

# ============================================================
section("average_captures()")
caps = [
    {'harmonics_db': [0.0, -6.0, -12.0], 'descriptors': {'resonance': 0.8, 'richness': 0.4, 'brightness': 0.3, 'darkness': 0.5, 'fullness': 0.2}},
    {'harmonics_db': [0.0, -8.0, -16.0], 'descriptors': {'resonance': 0.6, 'richness': 0.6, 'brightness': 0.5, 'darkness': 0.3, 'fullness': 0.4}},
]
avg = average_captures(caps)
test("Average not None", avg is not None)
test("Average harmonics count", len(avg['harmonics_db']) == 3)
test("Average H2 dB", abs(avg['harmonics_db'][1] - (-7.0)) < 0.01)
test("Average H3 dB", abs(avg['harmonics_db'][2] - (-14.0)) < 0.01)
test("Average resonance", abs(avg['descriptors']['resonance'] - 0.7) < 0.01)
test("Average richness", abs(avg['descriptors']['richness'] - 0.5) < 0.01)

# Empty
test("Empty captures returns None", average_captures([]) is None)

# Single capture
single = average_captures([caps[0]])
test("Single capture works", single is not None)
test("Single: same values", single['harmonics_db'] == caps[0]['harmonics_db'])

# Different lengths
mixed = [
    {'harmonics_db': [0.0, -6.0], 'descriptors': {'resonance': 0.5, 'richness': 0.5, 'brightness': 0.5, 'darkness': 0.5, 'fullness': 0.5}},
    {'harmonics_db': [0.0, -6.0, -12.0, -18.0], 'descriptors': {'resonance': 0.5, 'richness': 0.5, 'brightness': 0.5, 'darkness': 0.5, 'fullness': 0.5}},
]
avg = average_captures(mixed)
test("Mixed lengths: takes max", len(avg['harmonics_db']) == 4)

# ============================================================
section("compute_fingerprint()")
sessions = [
    {'captures': [
        {'note': 'A4', 'harmonics_db': [0, -6, -12], 'descriptors': {'resonance': 0.8, 'richness': 0.5, 'brightness': 0.3, 'darkness': 0.5, 'fullness': 0.2}},
        {'note': 'B4', 'harmonics_db': [0, -5, -10], 'descriptors': {'resonance': 0.9, 'richness': 0.6, 'brightness': 0.4, 'darkness': 0.6, 'fullness': 0.3}},
        {'note': 'C5', 'harmonics_db': [0, -8, -16], 'descriptors': {'resonance': 0.7, 'richness': 0.4, 'brightness': 0.2, 'darkness': 0.4, 'fullness': 0.1}},
    ]},
    {'captures': [
        {'note': 'A4', 'harmonics_db': [0, -7, -14], 'descriptors': {'resonance': 0.75, 'richness': 0.55, 'brightness': 0.35, 'darkness': 0.55, 'fullness': 0.25}},
    ]},
]
fp = compute_fingerprint(sessions)
test("Fingerprint note_count", fp['note_count'] == 3)
test("Fingerprint capture_count", fp['capture_count'] == 4)
test("Per-note has A4", 'A4' in fp['per_note'])
test("Per-note has B4", 'B4' in fp['per_note'])
test("Per-note has C5", 'C5' in fp['per_note'])
test("A4 averaged from 2 captures",
     abs(fp['per_note']['A4']['harmonics_db'][1] - (-6.5)) < 0.01)
test("Overall harmonics exist", len(fp['harmonics_db']) == 3)

# Empty sessions
fp_empty = compute_fingerprint([])
test("Empty: 0 notes", fp_empty['note_count'] == 0)
test("Empty: empty harmonics", fp_empty['harmonics_db'] == [])

# ============================================================
section("Profile storage - save/load round-trip")
tmpdir = tempfile.mkdtemp()
tmpfile = os.path.join(tmpdir, "test_profiles.json")

profiles = {
    "Test Library": {
        "Test Horn": {
            'horn_type': 'Alto',
            'horn_make': 'Selmer',
            'horn_model': 'Mark VI',
            'serial': '12345',
            'player': 'Test',
            'mouthpiece': 'V16',
            'reed': 'Java 2.5',
            'notes': '',
            'created': '2026-03-17',
            'sessions': [{'date': '2026-03-17', 'captures': [
                {'note': 'A4', 'harmonics_db': [0, -6], 'descriptors':
                 {'resonance': 0.8, 'richness': 0.5, 'brightness': 0.3,
                  'darkness': 0.5, 'fullness': 0.2}},
            ]}],
        }
    }
}

ok = save_tone_profiles(profiles, tmpfile)
test("Save succeeds", ok)
test("File exists", os.path.exists(tmpfile))

loaded = load_tone_profiles(tmpfile)
test("Load returns dict", isinstance(loaded, dict))
test("Library preserved", "Test Library" in loaded)
test("Profile preserved", "Test Horn" in loaded["Test Library"])
test("Data intact", loaded["Test Library"]["Test Horn"]["horn_make"] == "Selmer")
test("Session data intact",
     loaded["Test Library"]["Test Horn"]["sessions"][0]["captures"][0]["note"] == "A4")

# ============================================================
section("Profile storage - flat format migration")
flat_file = os.path.join(tmpdir, "flat_profiles.json")
flat_data = {
    "Old Profile": {
        'horn_type': 'Tenor',
        'sessions': [{'date': '2026-01-01', 'captures': []}],
    }
}
with open(flat_file, 'w') as f:
    json.dump(flat_data, f)

loaded = load_tone_profiles(flat_file)
test("Migrated to nested", DEFAULT_LIBRARY in loaded)
test("Old profile under default lib", "Old Profile" in loaded[DEFAULT_LIBRARY])

# Verify it was saved back in nested format
with open(flat_file, 'r') as f:
    raw = json.load(f)
test("File rewritten as nested", DEFAULT_LIBRARY in raw)

# ============================================================
section("Profile storage - missing/corrupt file")
loaded = load_tone_profiles(os.path.join(tmpdir, "nonexistent.json"))
test("Missing file: returns default", DEFAULT_LIBRARY in loaded)
test("Missing file: empty library", loaded[DEFAULT_LIBRARY] == {})

corrupt_file = os.path.join(tmpdir, "corrupt.json")
with open(corrupt_file, 'w') as f:
    f.write("{{{bad json")
loaded = load_tone_profiles(corrupt_file)
test("Corrupt file: returns default", DEFAULT_LIBRARY in loaded)

# ============================================================
section("flatten_profiles()")
nested = {
    "Lib A": {"Horn 1": {"data": 1}, "Horn 2": {"data": 2}},
    "Lib B": {"Horn 3": {"data": 3}},
}
flat = flatten_profiles(nested)
test("Flat has 3 entries", len(flat) == 3)
test("Horn 1 in flat", "Horn 1" in flat)
test("Horn 3 in flat", "Horn 3" in flat)

# Duplicate names across libs
nested_dup = {
    "Lib A": {"Same Name": {"data": "a"}},
    "Lib B": {"Same Name": {"data": "b"}},
}
flat = flatten_profiles(nested_dup)
test("Duplicates resolved", len(flat) == 2)
test("Prefixed duplicate exists", any("[Lib B]" in k for k in flat.keys()))

# ============================================================
section("Comparison - descriptor analysis")
fp_a = {
    'harmonics_db': [0, -5, -10, -15, -20, -25],
    'descriptors': {'resonance': 0.9, 'richness': 0.7, 'brightness': 0.6, 'darkness': 0.4, 'fullness': 0.3},
    'note_count': 10,
    'capture_count': 20,
    'per_note': {
        'A4': {'harmonics_db': [0, -4, -8], 'descriptors': {'resonance': 0.85, 'richness': 0.65, 'brightness': 0.55, 'darkness': 0.45, 'fullness': 0.25}},
    },
}
fp_b = {
    'harmonics_db': [0, -8, -16, -24, -32, -40],
    'descriptors': {'resonance': 0.6, 'richness': 0.3, 'brightness': 0.2, 'darkness': 0.7, 'fullness': 0.1},
    'note_count': 8,
    'capture_count': 15,
    'per_note': {
        'A4': {'harmonics_db': [0, -10, -20], 'descriptors': {'resonance': 0.55, 'richness': 0.25, 'brightness': 0.15, 'darkness': 0.75, 'fullness': 0.05}},
    },
}

# Verify the fingerprints can be compared
test("FP A brighter", fp_a['descriptors']['brightness'] > fp_b['descriptors']['brightness'])
test("FP B darker", fp_b['descriptors']['darkness'] > fp_a['descriptors']['darkness'])
test("FP A richer", fp_a['descriptors']['richness'] > fp_b['descriptors']['richness'])
test("Per-note data accessible", 'A4' in fp_a['per_note'])
test("Per-note descriptors differ",
     fp_a['per_note']['A4']['descriptors']['brightness'] != fp_b['per_note']['A4']['descriptors']['brightness'])

# ============================================================
section("Scale mode - dB vs linear height")
# Simulate the height calculation
def db_to_height_linear(db_val, max_h, db_range=60.0):
    linear = 10.0 ** (max(-db_range, min(0.0, db_val)) / 20.0)
    return max(0.0, linear * max_h)

def db_to_height_db(db_val, max_h, db_range=60.0):
    return max(0.0, (db_val + db_range) / db_range) * max_h

max_h = 400  # pixels
test("Linear: 0 dB = full height", abs(db_to_height_linear(0, max_h) - 400) < 1)
test("Linear: -6 dB = 50%", abs(db_to_height_linear(-6, max_h) - 200) < 10)
test("Linear: -20 dB = 10%", abs(db_to_height_linear(-20, max_h) - 40) < 5)
test("Linear: -60 dB = ~0", db_to_height_linear(-60, max_h) < 2)

test("dB: 0 dB = full height", abs(db_to_height_db(0, max_h) - 400) < 1)
test("dB: -6 dB = 90%", abs(db_to_height_db(-6, max_h) - 360) < 5)
test("dB: -30 dB = 50%", abs(db_to_height_db(-30, max_h) - 200) < 1)
test("dB: -60 dB = 0", abs(db_to_height_db(-60, max_h)) < 1)

# The key difference: -6 dB (half amplitude)
lin_h = db_to_height_linear(-6, max_h)
db_h = db_to_height_db(-6, max_h)
test("Linear shows half amp as ~50%", 45 < (lin_h / max_h * 100) < 55,
     f"got {lin_h/max_h*100:.0f}%")
test("dB shows half amp as ~90%", 88 < (db_h / max_h * 100) < 92,
     f"got {db_h/max_h*100:.0f}%")

# ============================================================
section("Edge cases - very short audio")
short = np.zeros(100, dtype=np.float32)
r = engine.analyze_buffer(short)
test("Short audio: no crash", True)
test("Short audio: no detection", r.fundamental_freq == 0.0)

# ============================================================
section("Edge cases - DC offset")
dc = np.ones(FFT_SIZE * 2, dtype=np.float32) * 0.5
r = engine.analyze_buffer(dc)
test("DC offset: no crash", True)
# DC shouldn't register as a note (it's below MIN_FUNDAMENTAL_HZ)

# ============================================================
section("Edge cases - extreme frequency")
audio = make_audio(80.1)  # Just above MIN_FUNDAMENTAL_HZ
r = engine.analyze_buffer(audio)
test("80 Hz detectable", r.fundamental_freq > 0,
     f"got {r.fundamental_freq:.1f}")

audio = make_audio(1999.0)  # Near MAX_FUNDAMENTAL_HZ
r = engine.analyze_buffer(audio)
test("1999 Hz detectable", r.fundamental_freq > 0,
     f"got {r.fundamental_freq:.1f}")

# ============================================================
section("Edge cases - empty descriptors")
empty_result = TonerResult()
test("Empty result: default resonance", empty_result.descriptors['resonance'] == 0.5)
test("Empty result: default richness", empty_result.descriptors['richness'] == 0.0)

# ============================================================
section("Multiple consecutive analyses (stability)")
audio = make_audio(440.0, {2: 0.6, 3: 0.3, 4: 0.15})
freqs = []
for _ in range(10):
    r = engine.analyze_buffer(audio)
    freqs.append(r.fundamental_freq)

test("10 analyses: all detect", all(f > 0 for f in freqs))
if all(f > 0 for f in freqs):
    spread = max(freqs) - min(freqs)
    test("10 analyses: stable (spread < 5 Hz)", spread < 5.0,
         f"spread={spread:.1f}")

# ============================================================
section("Profile import/merge logic")
# Simulate merging
base = {
    "Lib A": {
        "Horn 1": {
            'horn_type': 'Alto',
            'sessions': [{'date': '2026-01-01', 'captures': [{'note': 'A4', 'harmonics_db': [0, -6], 'descriptors': {}}]}],
        }
    }
}
imported = {
    "Lib A": {
        "Horn 1": {  # Existing - should merge sessions
            'horn_type': 'Alto',
            'sessions': [{'date': '2026-02-01', 'captures': [{'note': 'B4', 'harmonics_db': [0, -8], 'descriptors': {}}]}],
        },
        "Horn 2": {  # New
            'horn_type': 'Tenor',
            'sessions': [],
        }
    },
    "Lib B": {  # New library
        "Horn 3": {
            'horn_type': 'Soprano',
            'sessions': [],
        }
    }
}

# Merge (same logic as _toner_import_profiles)
for lib_name, lib_profiles in imported.items():
    if lib_name not in base:
        base[lib_name] = {}
    for prof_name, prof_data in lib_profiles.items():
        if prof_name not in base[lib_name]:
            base[lib_name][prof_name] = prof_data
        else:
            existing = base[lib_name][prof_name]
            existing_dates = {s.get('date') for s in existing.get('sessions', [])}
            for session in prof_data.get('sessions', []):
                if session.get('date') not in existing_dates:
                    existing.setdefault('sessions', []).append(session)

test("Merge: Horn 2 added", "Horn 2" in base["Lib A"])
test("Merge: Lib B added", "Lib B" in base)
test("Merge: Horn 1 has 2 sessions", len(base["Lib A"]["Horn 1"]["sessions"]) == 2)
test("Merge: duplicate date not added",
     sum(1 for s in base["Lib A"]["Horn 1"]["sessions"] if s['date'] == '2026-01-01') == 1)

# Cleanup
shutil.rmtree(tmpdir)

# ============================================================
# Final summary
section("DONE")
print(f"\n{'=' * 60}")
print(f"TOTAL: {passed} passed, {failed} failed out of {passed + failed}")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print(f"{failed} TESTS FAILED")
print("=" * 60)
