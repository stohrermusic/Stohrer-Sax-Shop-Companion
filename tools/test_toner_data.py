"""Characterization tests for the tone analyzer's DATA side (headless).

Pins what toner_engine does today with stored presets, fingerprints,
descriptors, offline WAV analysis and note names, plus the config-side
migrations of the tone library file. No Tk root, no audio device, no WAV
corpus. The golden numbers were computed by running the code as it stood on
2026-11-10; a change to any of them is a change in what the Analyze tool
shows for a user's existing recordings, so it should be deliberate.
"""

import json
import math
import os
import subprocess
import sys
import tempfile
import time
import wave

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Never the user's real settings folder (tools/test_config_isolation.py).
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-toner-data-")

from i18n import init_translation
init_translation("en")

import numpy as np
import toner_engine as te
from audio_utils import synthetic_tone

passed = 0
failed = 0
_t0 = time.time()

def check(name, fn):
    global passed, failed
    try:
        ok = fn()
    except Exception as e:
        ok = False
        name = f"{name}  [{type(e).__name__}: {e}]"
    if ok is False:
        failed += 1
        print(f"  FAIL: {name}")
    else:
        passed += 1
        print(f"  PASS: {name}")

def near(a, b, tol=1e-6):
    assert a is not None and abs(a - b) <= tol, f"got {a!r}, expected {b!r}"

def near_dict(d, golden):
    for k, v in golden.items():
        if isinstance(v, bool):
            assert d[k] is v, f"{k}: got {d[k]!r}, expected {v!r}"
        else:
            near(d[k], v)

def near_list(got, exp, tol=1e-6):
    """Element by element within tol: an average or a rounding can differ in the last place between
    numpy versions (CI's locked 2.4.6 on Python 3.11 against an unlocked 2.4.2 here, 2026-11-11), so no computed float
    is compared exactly."""
    assert len(got) == len(exp), f"got {got!r}, expected {exp!r}"
    for g, e in zip(got, exp):
        near(g, e, tol)


TMP = tempfile.mkdtemp(prefix="ssc-toner-data-files-")

STD = [0.0, -3.0, -6.0, -10.0, -14.0, -18.0, -22.0, -26.0, -30.0, -34.0]
THIN = [0.0, -20.0, -35.0, -50.0]
WARM = [0.0, -1.0, -4.0, -6.0, -9.0, -12.0, -15.0, -18.0, -21.0, -24.0]
BUMPY = [0.0, -3.0, -9.0, -7.0, -15.0, -12.0, -21.0, -17.0, -27.0, -22.0]

def cap(note, h, f, method="free", **extra):
    """A capture dict shaped as the app writes it."""
    d = {"note": note, "fundamental_freq": f, "harmonics_db": list(h),
         "harmonic_cents": [0.0] * len(h), "spectral_centroid": 512.3,
         "signal_level": 0.42, "timestamp": "2026-11-10T10:15:00",
         "n_frames": 30, "method": method}
    d.update(extra)
    return d


# ============================================================
# 1. STORAGE
# ============================================================
print("\n=== 1. Storage ===")

LIB = {
    "Shop Horns": {
        "Mark VI 145xxx": {
            "horn_type": "Alto", "player": "Matt", "mouthpiece": "Meyer 5M",
            "reed": "Vandoren 3", "mic_type": "condenser", "mic_model": "AT2020",
            "notes": "rich horn", "sandbox": False, "suppress_rolloff_warning": True,
            "sessions": [
                {"date": "2026-11-01", "mic_type": "condenser", "rolloff_rate": 2.1,
                 "recording_file": "C:/rec/mark6_20261101.wav",
                 "captures": [cap("A3", STD, 220.0),
                              cap("C#4", WARM, 277.18, method="calibration",
                                  written_note="A#4", detected_as="C#4")]},
                {"date": "2026-11-02", "mic_type": "condenser",
                 "captures": [cap("D4", THIN, 293.66, method="file",
                                  source_file="mark6.wav")]},
            ]},
        "Contact mic test": {"sandbox": True, "sessions": []},
    },
}

def t_round_trip():
    p = os.path.join(TMP, "round.json")
    assert te.save_tone_presets(LIB, p) is True, "te.save_tone_presets(LIB, p) is True"
    assert not os.path.exists(p + ".tmp"), "not os.path.exists(p + \".tmp\")"
    assert te.load_tone_presets(p) == LIB, "te.load_tone_presets(p) == LIB"
check("nested library with 2 sessions, free/calibration/file captures and optional keys round-trips", t_round_trip)

def t_key_heights_link():
    """Round 11 (Matt, 2026-11-11): a preset may link to one Key Heights library entry; a session
    snapshots the link and a copy of the heights at its creation; nothing on disk is migrated."""
    key_lib = {"Shop": {"Mark VI 145xxx": {"make": "Selmer", "model": "Mark VI", "size": "Alto", "serial": "",
                                           "notes": "", "units": "mm",
                                           "heights": {"B": "2.5", "F": "7.0", "Low C": "9.0"}}}}
    link = {"library": "Shop", "set": "Mark VI 145xxx"}
    assert te.key_heights_link({"key_heights": link}) == link, te.key_heights_link({"key_heights": link})
    assert te.key_heights_link({"key_heights": link}) is not link, "a copy, not the preset's own dict"
    for bad in ({}, {"key_heights": None}, {"key_heights": "Shop/Mark VI"}, {"key_heights": {"library": "Shop"}},
                {"key_heights": {"library": "", "set": "x"}}, {"key_heights": {"library": 1, "set": "x"}}, "not a dict"):
        assert te.key_heights_link(bad) is None, bad
    assert te.key_heights_entry(link, key_lib) == key_lib["Shop"]["Mark VI 145xxx"], "the entry is not found"
    assert te.key_heights_entry({"library": "Shop", "set": "gone"}, key_lib) is None, "a gone set must read as None"
    assert te.key_heights_entry(link, {}) is None and te.key_heights_entry(None, key_lib) is None, "empty cases"
    snap = te.key_heights_snapshot(link, key_lib)
    assert snap == {"library": "Shop", "set": "Mark VI 145xxx", "units": "mm",
                    "heights": {"B": "2.5", "F": "7.0", "Low C": "9.0"}}, snap
    assert snap["heights"] is not key_lib["Shop"]["Mark VI 145xxx"]["heights"], "the snapshot must copy the heights"
    assert te.key_heights_snapshot({"library": "Shop", "set": "gone"}, key_lib) == {"library": "Shop", "set": "gone"}, \
        "a missing entry snapshots the link alone"
    assert te.key_heights_snapshot(None, key_lib) is None, "no link, no snapshot"
    # storage: the link travels as any other preset field, and so does a session's snapshot
    lib = {"L": {"P": {"horn_type": "Alto", "key_heights": link,
                       "sessions": [{"date": "2026-11-11 10:00:00", "key_heights": snap, "captures": []}]}}}
    p = os.path.join(TMP, "kh.json")
    assert te.save_tone_presets(lib, p) is True and te.load_tone_presets(p) == lib, "the link did not round-trip"
    # the import unwrapper does not know the field and does not need to
    assert te.unwrap_tone_import(lib) == lib, "unwrap changed a plain library"
    assert te.unwrap_tone_import({"format": "wrapped-tone-presets", "format_version": 1, "libraries": lib}) == lib, \
        "unwrap changed a wrapped library"
check("the Key heights link: read from a preset (anything unreadable is none), its entry looked up, a session's "
      "snapshot with a copy of the heights (the link alone when the entry is gone); the link and the snapshot travel "
      "through save/load and the import unwrapper untouched", t_key_heights_link)

def t_flat_migrates():
    p = os.path.join(TMP, "flat.json")
    flat = {"Old Tenor": {"horn_type": "Tenor", "sessions": [{"captures": [cap("A3", STD, 220.0)]}]}}
    with open(p, "w") as f:
        json.dump(flat, f)
    assert te.load_tone_presets(p) == {"My Presets": flat}, "te.load_tone_presets(p) == {\"My Presets\": flat}"
    with open(p) as f:
        assert json.load(f) == {"My Presets": flat}, "flat file was not rewritten as nested"
check("flat file migrates to {'My Presets': ...} and is rewritten nested on disk", t_flat_migrates)

def t_preset_named_sessions():
    p = os.path.join(TMP, "named.json")
    data = {"Lib": {"sessions": {"horn_type": "Alto", "sessions": []}}}
    with open(p, "w") as f:
        json.dump(data, f)
    assert te.load_tone_presets(p) == data, "te.load_tone_presets(p) == data"
check("a preset literally named 'sessions' does not trigger the flat migration", t_preset_named_sessions)

check("missing file -> {'My Presets': {}}",
      lambda: te.load_tone_presets(os.path.join(TMP, "absent.json")) == {"My Presets": {}})

def t_top_list():
    p = os.path.join(TMP, "list.json")
    with open(p, "w") as f:
        json.dump([1, 2], f)
    assert te.load_tone_presets(p) == {"My Presets": {}}, "te.load_tone_presets(p) == {\"My Presets\": {}}"
    assert not os.path.exists(p + ".corrupt.bak"), "not os.path.exists(p + \".corrupt.bak\")"
check("top-level JSON list -> default, no .corrupt.bak", t_top_list)

def t_corrupt():
    p = os.path.join(TMP, "corrupt.json")
    with open(p, "w") as f:
        f.write('{"My Presets": {"A": ')
    assert te.load_tone_presets(p) == {"My Presets": {}}, "te.load_tone_presets(p) == {\"My Presets\": {}}"
    with open(p + ".corrupt.bak") as f:
        assert f.read() == '{"My Presets": {"A": ', "f.read() == '{\"My Presets\": {\"A\": '"
    with open(p, "w") as f:
        f.write("second garbage")
    assert te.load_tone_presets(p) == {"My Presets": {}}, "te.load_tone_presets(p) == {\"My Presets\": {}}"
    with open(p + ".corrupt.bak") as f:
        assert f.read() == '{"My Presets": {"A": ', "backup was overwritten"
check("corrupt JSON -> default and .corrupt.bak written once (second corruption keeps the first)", t_corrupt)

def t_dir_path():
    d = os.path.join(TMP, "a_dir.json")
    os.mkdir(d)
    assert te.load_tone_presets(d) == {"My Presets": {}}, "te.load_tone_presets(d) == {\"My Presets\": {}}"
check("a path that is a directory loads as the default", t_dir_path)

def t_unwritable():
    p = os.path.join(TMP, "no_such_dir", "x.json")
    assert te.save_tone_presets(LIB, p) is False, "te.save_tone_presets(LIB, p) is False"
    assert not os.path.exists(p + ".tmp"), "not os.path.exists(p + \".tmp\")"
check("save to a path in a missing folder -> False, no .tmp", t_unwritable)

def t_save_onto_dir():
    d = os.path.join(TMP, "save_dir")
    os.mkdir(d)
    assert te.save_tone_presets(LIB, d) is False, "te.save_tone_presets(LIB, d) is False"
    assert os.path.exists(d + ".tmp"), "today the .tmp is left behind (see report)"
check("save onto a directory -> False and a stale .tmp is LEFT (pinned as-is)", t_save_onto_dir)

def t_save_unserializable():
    p = os.path.join(TMP, "badval.json")
    assert te.save_tone_presets({"L": {"P": {1, 2}}}, p) is False, "te.save_tone_presets({\"L\": {\"P\": {1, 2}}}, p) is False"
    assert not os.path.exists(p) and os.path.exists(p + ".tmp"), "not os.path.exists(p) and os.path.exists(p + \".tmp\")"
check("save of unserializable data -> False, target untouched, partial .tmp LEFT (pinned as-is)",
      t_save_unserializable)

check("flatten_presets: duplicate name across libraries gets '[lib] name'; non-dict libraries skipped",
      lambda: te.flatten_presets({"A": {"x": 1, "y": 2}, "B": {"x": 3}, "C": "bad"})
      == {"x": 1, "y": 2, "[B] x": 3})


# ============================================================
# 2. CONFIG MIGRATIONS (subprocess: config runs them at import)
# ============================================================
print("\n=== 2. Config migrations ===")

def run_config_import(cfg_files, cwd_files):
    cfg = tempfile.mkdtemp(prefix="ssc-toner-cfg-")
    cwd = tempfile.mkdtemp(prefix="ssc-toner-cwd-")
    for folder, files in ((cfg, cfg_files), (cwd, cwd_files)):
        for name, text in files.items():
            with open(os.path.join(folder, name), "w") as f:
                f.write(text)
    env = dict(os.environ, SAXSHOP_CONFIG_DIR=cfg)
    code = f"import sys; sys.path.insert(0, {ROOT!r}); import config"
    r = subprocess.run([sys.executable, "-c", code], cwd=cwd, env=env,
                       capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr
    return cfg, cwd


OLD = '{"My Presets": {"Old": {"sessions": []}}}'

def t_rename():
    cfg, _cwd = run_config_import({"tone_profiles.json": OLD}, {})
    with open(os.path.join(cfg, "toner_data.json")) as f:
        assert f.read() == OLD, "f.read() == OLD"
    assert not os.path.exists(os.path.join(cfg, "tone_profiles.json")), "expected os.rename (old file gone)"
check("config import renames tone_profiles.json -> toner_data.json (os.rename: old file gone)", t_rename)

def t_rename_skipped():
    cfg, _cwd = run_config_import({"tone_profiles.json": OLD, "toner_data.json": "{}"}, {})
    with open(os.path.join(cfg, "toner_data.json")) as f:
        assert f.read() == "{}", "f.read() == \"{}\""
    assert os.path.exists(os.path.join(cfg, "tone_profiles.json")), "os.path.exists(os.path.join(cfg, \"tone_profiles.json\"))"
check("no rename when toner_data.json already exists (both files kept)", t_rename_skipped)

def t_legacy_cwd():
    cfg, cwd = run_config_import({}, {"tone_profiles.json": OLD})
    with open(os.path.join(cfg, "toner_data.json")) as f:
        assert f.read() == OLD, "f.read() == OLD"
    assert not os.path.exists(os.path.join(cfg, "tone_profiles.json")), "not os.path.exists(os.path.join(cfg, \"tone_profiles.json\"))"
    assert os.path.exists(os.path.join(cwd, "tone_profiles.json")), "legacy copy uses copy2"
check("legacy cwd tone_profiles.json: copy2 into config dir, then renamed; cwd file kept", t_legacy_cwd)

def t_legacy_cwd_new_name():
    cfg, cwd = run_config_import({}, {"toner_data.json": OLD})
    with open(os.path.join(cfg, "toner_data.json")) as f:
        assert f.read() == OLD, "f.read() == OLD"
    assert os.path.exists(os.path.join(cwd, "toner_data.json")), "os.path.exists(os.path.join(cwd, \"toner_data.json\"))"
check("legacy cwd toner_data.json is copied into the config dir", t_legacy_cwd_new_name)


# ============================================================
# 3. FINGERPRINTS AND DESCRIPTORS
# ============================================================
print("\n=== 3. Fingerprints and descriptors ===")

G_STD = {"richness": 0.11229521096727454, "warmth": 0.75, "even_odd": 0.7030597868308686,
         "rolloff_shape": 0.0, "low_harmonic_data": False}
G_THIN = {"richness": 0.05, "warmth": 0.0, "even_odd": 1.0, "rolloff_shape": 0.0, "low_harmonic_data": False}
G_WARM = {"richness": 0.4760317345070095, "warmth": 0.9166666666666666, "even_odd": 0.6629043820956366,
          "rolloff_shape": 0.0, "low_harmonic_data": False}
G_BUMPY = {"richness": 0.48017595839592825, "warmth": 0.75, "even_odd": 1.0,
           "rolloff_shape": 0.46604938271604934, "low_harmonic_data": False}
check("descriptors STD: richness 0.112295211, warmth 0.75, even_odd 0.703059787, rolloff_shape 0",
      lambda: near_dict(te.descriptors_from_harmonics(STD, 220.0, "Alto"), G_STD))
check("descriptors THIN: richness 0.05 (one significant upper), warmth 0, even_odd 1.0",
      lambda: near_dict(te.descriptors_from_harmonics(THIN, 220.0, "Alto"), G_THIN))
check("descriptors WARM: richness 0.476031735, warmth 0.916666667, even_odd 0.662904382",
      lambda: near_dict(te.descriptors_from_harmonics(WARM, 220.0, "Alto"), G_WARM))
check("descriptors BUMPY: richness 0.480175958, even_odd 1.0, rolloff_shape 0.466049383",
      lambda: near_dict(te.descriptors_from_harmonics(BUMPY, 220.0, "Alto"), G_BUMPY))
check("descriptors do not depend on sax_type (STD on Tenor == Alto)",
      lambda: near_dict(te.descriptors_from_harmonics(STD, 220.0, "Tenor"), G_STD))
check("descriptors [0,-3,-40]: low_harmonic_data True, richness 0.05",
      lambda: near_dict(te.descriptors_from_harmonics([0.0, -3.0, -40.0], 220.0, "Alto"),
                        {"richness": 0.05, "warmth": 0.75, "low_harmonic_data": True}))
ZERO = {"richness": 0.0, "warmth": 0.0, "even_odd": 0.0, "rolloff_shape": 0.0, "low_harmonic_data": False}
check("descriptors of [] or f0=0 are all zero",
      lambda: te.descriptors_from_harmonics([], 220.0) == ZERO and te.descriptors_from_harmonics(STD, 0) == ZERO)

def t_rolloff():
    near(te.compute_rolloff_rate(STD), 3.875)
    near(te.compute_rolloff_rate(WARM), 2.875)
    near(te.compute_rolloff_rate(STD + [-38.0, -42.0, -46.0]), 3.9)  # stops at H12
    assert te.compute_rolloff_rate(THIN) is None, "te.compute_rolloff_rate(THIN) is None"
    assert te.compute_rolloff_rate([0, -1, -2, -3, -4, -5, -6]) is None, "te.compute_rolloff_rate([0, -1, -2, -3, -4, -5, -6]) is None"
    near(te.compute_rolloff_rate([0, -1, -2, -3, -4, -5, -6, -8]), 7 / 6)
check("rolloff_rate: STD 3.875, WARM 2.875, 13 entries 3.9 (H2..H12), <8 entries None", t_rolloff)

S1 = {"date": "2026-11-01", "mic_type": "condenser",
      "captures": [cap("A3", STD, 220.0), cap("A3", WARM, 220.0), cap("D4", THIN, 293.66)]}

def t_fp_per_note_first():
    fp = te.compute_fingerprint([S1], "Alto")
    assert fp["note_count"] == 2 and fp["capture_count"] == 3, "fp[\"note_count\"] == 2 and fp[\"capture_count\"] == 3"
    assert sorted(fp["per_note"]) == ["A3", "D4"], "sorted(fp[\"per_note\"]) == [\"A3\", \"D4\"]"
    near_list(fp["per_note"]["A3"]["harmonics_db"], [0.0, -2.0, -5.0, -8.0, -11.5, -15.0, -18.5, -22.0, -25.5, -29.0])
    near(fp["per_note"]["A3"]["descriptors"]["richness"], 0.29259972374826704)
    # Equal weight per note: mean of A3 (0.2926) and D4 (0.05), not of 3 captures
    near(fp["descriptors"]["richness"], (0.29259972374826704 + 0.05) / 2)
    near_dict(fp["descriptors"], {"richness": 0.1712998618741335, "warmth": 0.4166666666666667,
                                  "even_odd": 0.8423270852756358, "rolloff_shape": 0.0, "evenness": 0.5})
    # Horn-level harmonics are the mean of all CAPTURES (H2: (-3-1-20)/3 = -8)
    near_list(fp["harmonics_db"], [0.0, -8.0, -15.0, -22.0, -11.5, -15.0, -18.5, -22.0, -25.5, -29.0])
    near(fp["rolloff_rate"], 2.625)
    assert fp["mic_type"] == "condenser", "fp[\"mic_type\"] == \"condenser\""
check("fingerprint: per-note mean first (richness 0.171299862 = mean of A3 0.292599724 and D4 0.05), "
      "warmth 0.416666667, even_odd 0.842327085, rolloff_rate 2.625", t_fp_per_note_first)

def t_fp_filter_and_mic():
    s2 = {"mic_type": "dynamic", "captures": [cap("E4", WARM, 329.63), cap("C7", STD, 2093.0),
                                              cap("F4", [0.0, 25.0, -5.0], 349.23)]}
    s3 = {"mic_type": "dynamic", "captures": []}
    fp = te.compute_fingerprint([S1, s2, s3], "Alto")
    assert fp["note_count"] == 3 and fp["capture_count"] == 4, "fp[\"note_count\"] == 3 and fp[\"capture_count\"] == 4"  # C7 out of range, +25 dB dropped
    assert fp["mic_type"] == "dynamic", "majority over sessions, empty sessions included"
    near_dict(fp["descriptors"], {"richness": 0.27287715275175883, "warmth": 0.5833333333333334,
                                  "even_odd": 0.7825195175489693, "evenness": 0.5})
    near(fp["rolloff_rate"], 2.6354166666666665)
check("fingerprint: out-of-range and +25 dB captures dropped; mic_type is the session majority 'dynamic'",
      t_fp_filter_and_mic)

S5 = {"date": "2026-11-05", "mic_type": "condenser",
      "captures": [cap(n, h, f) for n, h, f in [("A3", STD, 220.0), ("B3", WARM, 246.94), ("C#4", THIN, 277.18),
                                                 ("D4", STD, 293.66), ("E4", WARM, 329.63)]]}
check("fingerprint with 5 notes: evenness 0.525662436, richness 0.245330778, rolloff_rate 2.925",
      lambda: (near_dict(te.compute_fingerprint([S5], "Alto")["descriptors"],
                         {"evenness": 0.5256624360067756, "richness": 0.2453307781897136,
                          "warmth": 0.6666666666666666, "even_odd": 0.7463856675706021}),
               near(te.compute_fingerprint([S5], "Alto")["rolloff_rate"], 2.925)))

def t_fp_march_labels():
    """Matt, 2026-11-11: "Yes, apply it." The fingerprint groups a capture with the March 2026 label (exactly
    two transpositions below its pitch, the corrected note in the horn's range) under the note it is, through
    the chart's one rule (toner_engine.corrected_capture); a slip an octave above that again is left where it
    was; an ordinary preset's fingerprint is identical to the one computed without the rule."""
    # a tenor's written E4 is concert D3 (midi 50); the March presets stored it as A#0 (22), 28 semitones low
    d3, dsharp3 = te.midi_to_hz(50), te.midi_to_hz(51)
    march = {"captures": [cap("A#0", STD, d3), cap("A#0", WARM, d3 * 1.004), cap("B0", THIN, dsharp3)]}
    fp = te.compute_fingerprint([march], "Tenor")
    assert sorted(fp["per_note"]) == ["D#3", "D3"], f"the March labels were not grouped under their notes: {sorted(fp['per_note'])}"
    assert fp["corrected"] == 3 and fp["capture_count"] == 3 and fp["note_count"] == 2, (fp["corrected"], fp["capture_count"], fp["note_count"])
    near_list(fp["per_note"]["D3"]["harmonics_db"], [0.0, -2.0, -5.0, -8.0, -11.5, -15.0, -18.5, -22.0, -25.5, -29.0])
    # the same captures under their true labels give the same fingerprint, key for key
    true = {"captures": [cap("D3", STD, d3), cap("D3", WARM, d3 * 1.004), cap("D#3", THIN, dsharp3)]}
    fp_true = te.compute_fingerprint([true], "Tenor")
    assert fp_true["corrected"] == 0, fp_true["corrected"]
    assert {k: v for k, v in fp.items() if k != "corrected"} == {k: v for k, v in fp_true.items() if k != "corrected"}, "corrected and true-labelled fingerprints differ"
    # the raw captures are not touched
    assert [c["note"] for c in march["captures"]] == ["A#0", "A#0", "B0"], "a capture dict was modified"
    # a slip an octave above the 2T (the detector's own, on the lowest asked notes) is not corrected; its label
    # A#0 is out of the tenor's range and it is dropped, exactly as before
    slip = {"captures": [cap("A#0", STD, te.midi_to_hz(50 + 12))]}
    fps = te.compute_fingerprint([slip], "Tenor")
    assert fps["corrected"] == 0 and fps["capture_count"] == 0 and fps["per_note"] == {}, (fps["corrected"], fps["capture_count"], sorted(fps["per_note"]))
    # a slip whose wrong label happens to lie in range is filed under that label, as before: the rule is not wider
    slip_in = {"captures": [cap("C3", STD, te.midi_to_hz(48 + 28 + 12))]}
    fpi = te.compute_fingerprint([slip_in], "Tenor")
    assert fpi["corrected"] == 0 and sorted(fpi["per_note"]) == ["C3"], (fpi["corrected"], sorted(fpi["per_note"]))
    # a corrected note still goes through the filter: an impossible harmonic keeps it out
    bad = {"captures": [cap("A#0", [0.0, 25.0, -5.0], d3)]}
    fpb = te.compute_fingerprint([bad], "Tenor")
    assert fpb["corrected"] == 1 and fpb["capture_count"] == 0, (fpb["corrected"], fpb["capture_count"])
    # an ordinary preset: identical to the fingerprint computed with the correction stood in as a no-op
    orig = te.corrected_capture
    te.corrected_capture = lambda c, sax, reference_hz=te.A4_HZ: (c, False)
    try:
        plain1, plain5 = te.compute_fingerprint([S1], "Alto"), te.compute_fingerprint([S5], "Alto")
    finally:
        te.corrected_capture = orig
    assert te.compute_fingerprint([S1], "Alto") == plain1 and te.compute_fingerprint([S5], "Alto") == plain5, "an ordinary preset's fingerprint changed under the rule"
    assert plain1["corrected"] == 0 and plain5["corrected"] == 0, (plain1["corrected"], plain5["corrected"])
    # one rule, one place: the chart's run corrects the same captures through the same function
    import intonation as I
    run = I.run_from_captures([march], "Tenor")
    assert run.corrected == 3 and run.notes == ["E4", "F4"], (run.corrected, run.notes)
    assert I.corrected_label is te.corrected_label and I.corrected_capture is te.corrected_capture, "the chart and the fingerprint do not share the rule"
    assert I.OFF_NOTE_CENTS == te.OFF_NOTE_CENTS == 50.0 and I.midi_to_hz is te.midi_to_hz, "the bound or the pitch helper is written twice"
check("fingerprint: the March 2026 labels (exactly two transpositions low, corrected note in range) are grouped under "
      "the note they are through the chart's one rule, the raw captures untouched; a slip an octave above stays out "
      "(or under its in-range label) as before; a corrected note still meets the filter; an ordinary preset is "
      "identical with the rule stood in as a no-op", t_fp_march_labels)

def t_plausible():
    assert te._capture_is_plausible(cap("C3", STD, 130.81), "Alto") is True, "te._capture_is_plausible(cap(\"C3\", STD, 130.81), \"Alto\") is True"
    assert te._capture_is_plausible(cap("B2", STD, 123.47), "Alto") is False, "te._capture_is_plausible(cap(\"B2\", STD, 123.47), \"Alto\") is False"
    assert te._capture_is_plausible(cap("A4", [0.0, 20.0], 440.0), "Alto") is True, "te._capture_is_plausible(cap(\"A4\", [0.0, 20.0], 440.0), \"Alto\") is True"
    assert te._capture_is_plausible(cap("A4", [0.0, 20.5], 440.0), "Alto") is False, "te._capture_is_plausible(cap(\"A4\", [0.0, 20.5], 440.0), \"Alto\") is False"
    assert te._capture_is_plausible(cap("A4", [0.0, float("-inf")], 440.0), "Alto") is False, "te._capture_is_plausible(cap(\"A4\", [0.0, float(\"-inf\")], 440.0), \"Alto\") is False"
    assert te._capture_is_plausible(cap("", STD, 440.0), "Alto") is False, "te._capture_is_plausible(cap(\"\", STD, 440.0), \"Alto\") is False"
    assert te._capture_is_plausible(cap("", STD, 440.0), "Kazoo") is True, "te._capture_is_plausible(cap(\"\", STD, 440.0), \"Kazoo\") is True"  # unknown type: no range check
check("_capture_is_plausible: range edges, +20 dB kept, +20.5 and -inf dropped, unknown type skips range",
      t_plausible)

def t_average():
    a = te.average_captures([{"harmonics_db": [0, -4, -8], "harmonic_cents": [0, 2], "fundamental_freq": 220},
                             {"harmonics_db": [0, -6], "freq": 222},
                             {"harmonics_db": [0, -2, -10, -20], "fundamental_freq": 0}])
    assert a == {"harmonics_db": [0.0, -4.0, -9.0, -20.0], "fundamental_freq": 221.0, "harmonic_cents": [0.0, 2.0]}, "a == {\"harmonics_db\": [0.0, -4.0, -9.0, -20.0], \"fundamental_freq\": 221.0, \"harmonic_cents\": [0.0, 2.0]}"
    assert te.average_captures([]) is None, "te.average_captures([]) is None"
check("average_captures: per-index mean over mixed lengths, 'freq' fallback, 0 Hz ignored, cents averaged",
      t_average)

check("find_out_of_range_notes Alto: ['B2', 'A#5', 'C7'] (C3, A5, A4 in; unparseable ignored)",
      lambda: te.find_out_of_range_notes(["B2", "C3", "A5", "A#5", "C7", "X", "A4"], "Alto")
      == ["B2", "A#5", "C7"] and te.find_out_of_range_notes(["C9"], "Kazoo") == [])

EMPTY = {"date": "2026-11-03", "captures": []}

def t_session_fp():
    assert te.compute_session_fingerprint(EMPTY) is None and te.compute_session_fingerprint({}) is None, "te.compute_session_fingerprint(EMPTY) is None and te.compute_session_fingerprint({}) is None"
    assert te.compute_session_fingerprint(S1, "Alto") == te.compute_fingerprint([S1], "Alto"), "te.compute_session_fingerprint(S1, \"Alto\") == te.compute_fingerprint([S1], \"Alto\")"
check("compute_session_fingerprint: None without captures, else compute_fingerprint([s])", t_session_fp)

def t_variation():
    assert te.compute_session_variation([S1, EMPTY], "Alto") is None, "te.compute_session_variation([S1, EMPTY], \"Alto\") is None"
    v = te.compute_session_variation([S1, S5, EMPTY], "Alto")
    assert v["session_count"] == 2 and [d for d, _ in v["session_fingerprints"]] == ["2026-11-01", "2026-11-05"], "v[\"session_count\"] == 2 and [d for d, _ in v[\"session_fingerprints\"]] == [\"2026-11-01\", \"2026-11-05\"]"
    st = v["descriptor_stats"]
    near_dict(st["richness"], {"mean": 0.20831532003192355, "stdev": 0.037015458157790046,
                               "min": 0.1712998618741335, "max": 0.2453307781897136})
    near_dict(st["warmth"], {"mean": 0.5416666666666666, "stdev": 0.125})
    near_dict(st["even_odd"], {"mean": 0.794356376423119, "stdev": 0.04797070885251686})
    assert st["richness"]["n"] == 2, st["richness"]
    near(st["rolloff_shape"]["stdev"], 0.0)
check("session variation: 1 session -> None; 2 -> richness mean 0.208315320 stdev 0.037015458 (population)",
      t_variation)

def t_group():
    g = te.compute_group_fingerprint([("P1", {"horn_type": "Alto", "sessions": [S1]}),
                                      ("P5", {"horn_type": "Alto", "sessions": [S5]}),
                                      ("E", {"sessions": [EMPTY]})])
    assert g["preset_count"] == 2 and g["total_captures"] == 8 and g["note_count"] == 5, "g[\"preset_count\"] == 2 and g[\"total_captures\"] == 8 and g[\"note_count\"] == 5"
    assert [n for n, _ in g["per_preset"]] == ["P1", "P5"], "[n for n, _ in g[\"per_preset\"]] == [\"P1\", \"P5\"]"
    near_dict(g["descriptors"], {"richness": 0.20831532003192355, "warmth": 0.5416666666666666,
                                 "even_odd": 0.794356376423119, "rolloff_shape": 0.0})
    assert "evenness" not in g["descriptors"], "\"evenness\" not in g[\"descriptors\"]"
    for got, exp in zip(g["harmonics_db"], [0.0, -6.8, -13.0, -19.2, -11.5, -15.0, -18.5, -22.0, -25.5, -29.0]):
        near(got, exp)
    assert g["per_note"]["A3"] == {"harmonics_db": [0.0, -2.5, -5.5, -9.0, -12.75, -16.5, -20.25, -24.0,
                                                    -27.75, -31.5], "fundamental_freq": 220.0}, "g[\"per_note\"][\"A3\"] == {\"harmonics_db\": [0.0, -2.5, -5.5, -9.0, -12.75, -16.5, -20.25, -24.0, -27.75, -31.5..."
    e = te.compute_group_fingerprint([])
    assert e["preset_count"] == 0 and e["descriptors"] == {"richness": 0.0, "warmth": 0.0}, "e[\"preset_count\"] == 0 and e[\"descriptors\"] == {\"richness\": 0.0, \"warmth\": 0.0}"
    assert "per_note" not in e and "note_count" not in e, "empty result lacks per_note/note_count today"
check("group fingerprint: equal weight per preset (richness 0.208315320), H2 -6.8, empty group shape", t_group)

def t_population():
    pop = te.compute_population_stats(
        [("L", "P1", {"horn_type": "Alto", "sessions": [S1]}),
         ("L", "P5", {"horn_type": "Alto", "sessions": [S5]}),
         ("L", "T", {"horn_type": "Tenor", "sessions": [S5]}),
         ("L", "E", {"horn_type": "Alto", "sessions": [EMPTY]}),
         ("L", "W", {"horn_type": "Alto", "sessions": [{"captures": [cap("E4", WARM, 329.63)]}]})], "Alto")
    assert pop["count"] == 3 and pop["sax_type"] == "Alto", (pop["count"], pop["sax_type"])
    dv = pop["descriptor_values"]
    for got, exp in zip(dv["richness"], [0.1712998618741335, 0.2453307781897136, 0.4760317345070095]):
        near(got, exp)
    near_list(dv["evenness"], [0.5, 0.5, 0.5256624360067756])
    near_list(pop["rolloff_values"], [2.625, 2.875, 2.925])
    # the rank of a value that IS in the population: asked with the population's own value, not a
    # 17-digit literal of it (under numpy 2.4.6 on Python 3.11, CI's, the middle richness is
    # 0.24533077818971352, one last place off the literal, and bisect_left put the literal after it: 67)
    assert te.percentile_rank(dv["richness"][1], dv["richness"]) == 33, (dv["richness"], te.percentile_rank(dv["richness"][1], dv["richness"]))
    assert te.percentile_rank(0.5, dv["richness"]) == 100, dv["richness"]
check("population stats Alto: 3 presets (Tenor and empty skipped), sorted lists, rolloff [2.625, 2.875, 2.925]",
      t_population)

check("percentile_rank (bisect_left, rounded): [0,0,25,75,75,100] over [.1,.2,.2,.5]; empty -> None",
      lambda: [te.percentile_rank(x, [0.1, 0.2, 0.2, 0.5]) for x in (0.0, 0.1, 0.2, 0.3, 0.5, 0.9)]
      == [0, 0, 25, 75, 75, 100] and te.percentile_rank(1, []) is None)


# ============================================================
# 4. OFFLINE WAV ANALYSIS
# ============================================================
print("\n=== 4. Offline WAV analysis ===")

HARM = (0.0, -3.0, -6.0, -10.0, -14.0, -18.0)

def build(parts, sr):
    out = []
    for freq, dur in parts:
        n = int(round(dur * sr))
        out.append(np.zeros(n, np.float32) if freq is None else synthetic_tone(0, n, freq, sr, HARM, 0.3))
    return np.concatenate(out)

def write_wav(name, parts, sr=44100, channels=1):
    pcm = (np.clip(build(parts, sr), -1, 1) * 32767).astype("<i2")
    if channels == 2:
        pcm = np.repeat(pcm, 2)
    path = os.path.join(TMP, name)
    with wave.open(path, "wb") as w:
        w.setnchannels(channels)
        w.setsampwidth(2)
        w.setframerate(sr)
        w.writeframes(pcm.tobytes())
    return path, len(pcm) // channels

def expected_segments(parts):
    """The code's rule, replayed: per-hop notes (each window's end time), runs of one note, the
    whole-window frames of each run (toner_engine.whole_window: the windows holding the onset or
    the release go), >= FREE_MIN_FRAMES. Until 2026-11-10 the rule was "minus three attack frames"."""
    x = (np.clip(build(parts, 44100), -1, 1) * 32767).astype("<i2").astype(np.float32) / 32768.0
    eng, res = te.TonerEngine(), []
    for s in range(0, len(x) - te.FFT_SIZE, te.FILE_HOP_SAMPLES):
        r = eng.analyze_buffer(x[s:s + te.FFT_SIZE])
        if r.fundamental_freq > 0 and r.harmonics:
            res.append((r.fundamental_note, (s + te.FFT_SIZE) / te.SAMPLE_RATE))
    runs = []
    for n, t in res:
        if runs and runs[-1][0] == n:
            runs[-1][1].append(t)
        else:
            runs.append([n, [t]])
    out = []
    for n, ts in runs:
        k = sum(te.whole_window(ts, ts[0], ts[-1]))
        if k >= te.FREE_MIN_FRAMES:
            out.append((n, k))
    return out


TWO = [(220.0, 1.5), (None, 0.5), (293.66, 1.5)]
_p2, _n2 = write_wav("two.wav", TWO)
CAPS = te.analyze_audio_file(_p2, te.TonerEngine())

def t_two_notes():
    assert [c["note"] for c in CAPS] == ["A3", "D4"], "[c[\"note\"] for c in CAPS] == [\"A3\", \"D4\"]"
    for c in CAPS:
        assert set(c) == {"note", "fundamental_freq", "harmonics_db", "harmonic_cents", "timestamp",
                          "n_frames", "method", "spectral_centroid", "level_dbfs"}, sorted(c)
        assert c["method"] == "file" and c["timestamp"] == "", "c[\"method\"] == \"file\" and c[\"timestamp\"] == \"\""
        # the level against full scale (2026-11-10): the generated
        # tones sit well under full scale (measured -16.7 dBFS here), and are not silence
        assert isinstance(c["level_dbfs"], float) and -25.0 < c["level_dbfs"] < -5.0, c["level_dbfs"]
    near(CAPS[0]["fundamental_freq"], 220.0)
    near(CAPS[1]["fundamental_freq"], 293.66)
    near(CAPS[0]["spectral_centroid"], 536.0)
    near(CAPS[1]["spectral_centroid"], 706.3)
check("two held notes -> exactly [A3, D4], method 'file', timestamp '', no signal_level, "
      "f0 220.0 / 293.66 Hz (the notes as generated), centroid 536.0 / 706.3", t_two_notes)

def t_cents_and_h2():
    # Before the whole-window rule (2026-11-10) A3 read -1.09 c and D4 +0.35 c: the windows holding
    # the release were averaged in (the attack was skipped, the release was not).
    cents = [1200 * math.log2(c["fundamental_freq"] / f) for c, f in zip(CAPS, (220.0, 293.66))]
    assert abs(cents[0]) < 0.05 and abs(cents[1]) < 0.05, cents
    assert CAPS[0]["harmonics_db"][1] == -3.34 and CAPS[1]["harmonics_db"][1] == -3.09, "CAPS[0][\"harmonics_db\"][1] == -3.34 and CAPS[1][\"harmonics_db\"][1] == -3.09"
    assert all(abs(c["harmonics_db"][1] + 3.0) <= 1.0 for c in CAPS), "all(abs(c[\"harmonics_db\"][1] + 3.0) <= 1.0 for c in CAPS)"
check("fundamental: A3 and D4 within 0.05 cents (were -1.09 and +0.35 with the release averaged in); "
      "H2 -3.34 / -3.09 dB", t_cents_and_h2)

def t_n_frames():
    exp = expected_segments(TWO)
    assert exp == [("A3", 32), ("D4", 32)], exp
    assert [(c["note"], c["n_frames"]) for c in CAPS] == exp, "[(c[\"note\"], c[\"n_frames\"]) for c in CAPS] == exp"
    assert (_n2 - te.FFT_SIZE) // te.FILE_HOP_SAMPLES + 1 == 135, "(_n2 - te.FFT_SIZE) // te.FILE_HOP_SAMPLES + 1 == 135"
    assert te.FFT_SIZE // te.FILE_HOP_SAMPLES == 16, "a window is 16 hops: 16 frames go at each end of a run"
check("n_frames 32 / 32 = detected run (64 hops of 1024) less the 16 hops of a window at either end "
      "(was 61: minus 3 attack frames); 135 windows", t_n_frames)

def t_short():
    p, _n = write_wav("short.wav", [(220.0, 0.5)])
    assert te.analyze_audio_file(p, te.TonerEngine()) == [], "te.analyze_audio_file(p, te.TonerEngine()) == []"
    p, _n = write_wav("one_second.wav", [(220.0, 1.0)])
    assert te.analyze_audio_file(p, te.TonerEngine()) == [], "a 1 s note has fewer than 10 whole-window frames"
    # the shortest note the file analyzer keeps: 10 frames plus a window (16 hops) at either end, 42 hops,
    # (FFT_SIZE + 41 * 1024) / 44100 = 1.32 s
    p, _n = write_wav("long_enough.wav", [(220.0, 1.4)])
    caps = te.analyze_audio_file(p, te.TonerEngine())
    assert [c["note"] for c in caps] == ["A3"] and caps[0]["n_frames"] >= te.FREE_MIN_FRAMES, caps
check("a note held 0.5 s or 1.0 s gives no capture (fewer than 10 whole-window frames); 1.4 s does", t_short)

def t_stereo_48k():
    sig = [(c["note"], c["n_frames"], c["fundamental_freq"]) for c in CAPS]
    p, _n = write_wav("stereo.wav", TWO, channels=2)
    assert [(c["note"], c["n_frames"], c["fundamental_freq"]) for c in te.analyze_audio_file(p, te.TonerEngine())] == sig, "[(c[\"note\"], c[\"n_frames\"], c[\"fundamental_freq\"]) for c in te.analyze_audio_file(p, te.TonerEngine())] == sig"
    p, _n = write_wav("48k.wav", TWO, sr=48000)
    assert [(c["note"], c["n_frames"], c["fundamental_freq"]) for c in te.analyze_audio_file(p, te.TonerEngine())] == sig, "[(c[\"note\"], c[\"n_frames\"], c[\"fundamental_freq\"]) for c in te.analyze_audio_file(p, te.TonerEngine())] == sig"
check("stereo and 48 kHz files give the same notes, n_frames and f0", t_stereo_48k)

def t_progress():
    p, n = write_wav("long.wav", [(220.0, 10.0)])
    calls = []
    caps = te.analyze_audio_file(p, te.TonerEngine(), progress_cb=lambda i, t: calls.append((i, t)))
    windows = (n - te.FFT_SIZE) // te.FILE_HOP_SAMPLES + 1
    assert windows == 415, "windows == 415"
    # every 200 windows and once at the end, with the exact total (fixed 2026-11-10: the total was one
    # short, nothing reported completion, and a file under 200 windows never called)
    assert calls == [(200, 415), (400, 415), (415, 415)], calls
    assert [(c["note"], c["n_frames"]) for c in caps] == [("A3", 383)], "[(c[\"note\"], c[\"n_frames\"]) for c in caps] == [(\"A3\", 383)]"   # 415 windows less 16 at either end
    calls.clear()
    te.analyze_audio_file(_p2, te.TonerEngine(), progress_cb=lambda i, t: calls.append((i, t)))
    _samples, _rate, _ch = te._read_wav_file(_p2)
    short_windows = len(range(0, len(_samples) // _ch - te.FFT_SIZE, te.FILE_HOP_SAMPLES))
    assert 0 < short_windows < 200 and calls == [(short_windows, short_windows)], (calls, short_windows)
    # the asked-notes reader reports the same way
    calls.clear()
    te.analyze_asked_notes(p, te.TonerEngine(), [], progress_cb=lambda i, t: calls.append((i, t)))
    assert calls == [(200, 415), (400, 415), (415, 415)], calls
check("progress_cb every 200 windows and once at the end with the exact total: [(200, 415), (400, 415), "
      "(415, 415)] for 415 windows; a short file reports once; analyze_asked_notes the same", t_progress)


# ---- the file path against known recordings (generated; 2026-11-10) ----
import struct   # noqa: E402

def hz(midi, cents=0.0):
    return 440.0 * 2 ** ((midi - 69 + cents / 100.0) / 12.0)

def cents_off(f, midi):
    return 1200.0 * math.log2(f / hz(midi))

def t_scale_known_pitches():
    # a chromatic scale A3..E4 (concert), a second each with a short silence between: every note found, in
    # order, within 2 cents of the pitch generated
    midis = list(range(57, 65))
    parts = []
    for m in midis:
        parts += [(hz(m), 1.0), (None, 0.3)]
    p, _n = write_wav("scale.wav", parts)
    caps = te.analyze_audio_file(p, te.TonerEngine())
    assert [c["note"] for c in caps] == [te.PITCH_CLASSES[m % 12] + str(m // 12 - 1) for m in midis], [c["note"] for c in caps]
    for c, m in zip(caps, midis):
        assert abs(cents_off(c["fundamental_freq"], m)) < 2.0, (c["note"], c["fundamental_freq"])
        assert c["n_frames"] >= te.FREE_MIN_FRAMES, "c[\"n_frames\"] >= te.FREE_MIN_FRAMES"
check("a chromatic scale of known pitches: every note found in order, each within 2 cents", t_scale_known_pitches)

def t_asked_notes():
    # a note played 60 cents flat: by ear it lands on its neighbour; asked for, it keeps its name and its
    # pitch is read as it was played
    parts = [(hz(57, -60), 1.5), (None, 0.5), (hz(62), 1.5)]
    p, _n = write_wav("asked.wav", parts)
    by_ear = te.analyze_audio_file(p, te.TonerEngine())
    assert [c["note"] for c in by_ear] == ["G#3", "D4"], [c["note"] for c in by_ear]
    asked = [{"note": "A3", "written_note": "B4", "wav_start_s": 0.0, "wav_end_s": 1.5},
             {"note": "D4", "written_note": "E5", "wav_start_s": 2.0, "wav_end_s": 3.5},
             {"note": "C4", "written_note": "D5", "wav_start_s": 1.5, "wav_end_s": 2.0},      # the silence: nothing read
             {"note": "C4", "written_note": "D5"}]                                           # no stretch: nothing read
    got = te.analyze_asked_notes(p, te.TonerEngine(), asked)
    assert len(got) == 4 and got[2] is None and got[3] is None, got
    a, d = got[0], got[1]
    assert a["note"] == "A3" and a["written_note"] == "B4" and a["method"] == "calibration" and a["detected_note"] == "G#3", "a[\"note\"] == \"A3\" and a[\"written_note\"] == \"B4\" and a[\"method\"] == \"calibration\" and a[\"detected_note\"] == ..."
    assert -62.0 < cents_off(a["fundamental_freq"], 57) < -58.0, a["fundamental_freq"]
    assert d["note"] == "D4" and d["detected_note"] == "D4" and abs(cents_off(d["fundamental_freq"], 62)) < 2.0, "d[\"note\"] == \"D4\" and d[\"detected_note\"] == \"D4\" and abs(cents_off(d[\"fundamental_freq\"], 62)) < 2.0"
    assert a["wav_start_s"] == 0.0 and a["wav_end_s"] == 1.5 and d["wav_start_s"] == 2.0, "a[\"wav_start_s\"] == 0.0 and a[\"wav_end_s\"] == 1.5 and d[\"wav_start_s\"] == 2.0"
    # only the windows wholly inside the stretch count (the rule the live capture keeps its frames by)
    assert a["n_frames"] == len([t for t in (k * te.FILE_HOP_SAMPLES / te.SAMPLE_RATE + te.WINDOW_S
                                             for k in range(0, 200)) if te.WINDOW_S <= t <= 1.5]), a["n_frames"]
    assert isinstance(a["level_dbfs"], float) and a["signal_level"] > 0 and "harmonic_cents" in a, "isinstance(a[\"level_dbfs\"], float) and a[\"signal_level\"] > 0 and \"harmonic_cents\" in a"
    assert set(a) >= {"spectral_centroid", "harmonics_db", "timestamp"}, "set(a) >= {\"spectral_centroid\", \"harmonics_db\", \"timestamp\"}"
check("analyze_asked_notes: a note played 60 cents flat lands on its neighbour by ear but keeps the asked name "
      "and its pitch when read under the stretch it was played in; nothing from silence or without a stretch",
      t_asked_notes)

def t_octave_prone():
    # a timbre with the 2nd harmonic louder than the fundamental: read at its fundamental, not an octave up
    loud_h2 = (0.0, 6.0, -6.0, -12.0, -18.0, -24.0)
    sr = 44100
    x = synthetic_tone(0, int(sr * 2.0), 220.0, sr, loud_h2, 0.3)
    pcm = (np.clip(x, -1, 1) * 32767).astype("<i2")
    p = os.path.join(TMP, "h2.wav")
    with wave.open(p, "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(sr)
        w.writeframes(pcm.tobytes())
    caps = te.analyze_audio_file(p, te.TonerEngine())
    assert [c["note"] for c in caps] == ["A3"], [c["note"] for c in caps]
    assert abs(cents_off(caps[0]["fundamental_freq"], 57)) < 2.0, "abs(cents_off(caps[0][\"fundamental_freq\"], 57)) < 2.0"
    asked = te.analyze_asked_notes(p, te.TonerEngine(), [{"note": "A3", "wav_start_s": 0.0, "wav_end_s": 2.0}])
    assert asked[0] and asked[0]["detected_note"] == "A3", "asked[0] and asked[0][\"detected_note\"] == \"A3\""
check("an octave-prone timbre (H2 louder than H1) is read at its fundamental, by ear and asked", t_octave_prone)

def t_short_and_silence():
    p, _n = write_wav("short.wav", [(hz(60), 0.5)])                 # shorter than the whole-window minimum
    assert te.analyze_audio_file(p, te.TonerEngine()) == [], "te.analyze_audio_file(p, te.TonerEngine()) == []"
    # asked for, a short answer is still an answer (as the live guided capture keeps whatever it heard):
    # the six windows wholly inside the half second, no minimum
    short = te.analyze_asked_notes(p, te.TonerEngine(), [{"note": "C4", "wav_start_s": 0.0, "wav_end_s": 0.5}])
    assert short[0]["note"] == "C4" and short[0]["n_frames"] == 6, short
    assert te.analyze_asked_notes(p, te.TonerEngine(), [{"note": "C4", "wav_start_s": 0.0, "wav_end_s": 0.3}]) == [None], "te.analyze_asked_notes(p, te.TonerEngine(), [{\"note\": \"C4\", \"wav_start_s\": 0.0, \"wav_end_s\": 0.3}]) == [None]"
    p, n = write_wav("silence.wav", [(None, 2.0)])
    calls = []
    assert te.analyze_audio_file(p, te.TonerEngine(), progress_cb=lambda i, t: calls.append((i, t))) == [], "te.analyze_audio_file(p, te.TonerEngine(), progress_cb=lambda i, t: calls.append((i, t))) == []"
    windows = len(range(0, n - te.FFT_SIZE, te.FILE_HOP_SAMPLES))
    assert calls == [(windows, windows)], calls
    assert te.analyze_asked_notes(p, te.TonerEngine(), [{"note": "C4", "wav_start_s": 0.0, "wav_end_s": 2.0}]) == [None], "te.analyze_asked_notes(p, te.TonerEngine(), [{\"note\": \"C4\", \"wav_start_s\": 0.0, \"wav_end_s\": 2.0}]) == [None]"
check("a note too short gives no capture; silence gives none and still reports its progress once", t_short_and_silence)

def _write_pcm(path, x, sampwidth, sr=44100, float_fmt=False):
    """A mono WAV by hand: 16/24/32-bit PCM or IEEE float32 (the wave module writes no float)."""
    tag = 1
    if float_fmt:
        data, tag, width = x.astype("<f4").tobytes(), 3, 4
    elif sampwidth == 2:
        data, width = (np.clip(x, -1, 1) * 32767).astype("<i2").tobytes(), 2
    elif sampwidth == 3:
        i32 = (np.clip(x, -1, 1) * 8388607).astype("<i4")
        data, width = np.frombuffer(i32.tobytes(), np.uint8).reshape(-1, 4)[:, :3].tobytes(), 3
    else:
        data, width = (np.clip(x, -1, 1) * 2147483647).astype("<i4").tobytes(), 4
    fmt = struct.pack("<HHIIHH", tag, 1, sr, sr * width, width, width * 8)
    with open(path, "wb") as f:
        f.write(b"RIFF" + struct.pack("<I", 36 + len(data)) + b"WAVE")
        f.write(b"fmt " + struct.pack("<I", len(fmt)) + fmt)
        f.write(b"data" + struct.pack("<I", len(data)) + data)

def t_bit_depths():
    x = build(TWO, 44100)
    got = {}
    for label, kw in (("16", dict(sampwidth=2)), ("24", dict(sampwidth=3)), ("32", dict(sampwidth=4)),
                      ("f32", dict(sampwidth=4, float_fmt=True))):
        p = os.path.join(TMP, f"depth{label}.wav")
        _write_pcm(p, x, **kw)
        got[label] = [(c["note"], c["n_frames"], c["fundamental_freq"]) for c in te.analyze_audio_file(p, te.TonerEngine())]
    assert [g[:2] for g in got["16"]] == [("A3", 32), ("D4", 32)], got["16"]
    for label in ("24", "32", "f32"):
        assert [(n, k) for n, k, _f in got[label]] == [(n, k) for n, k, _f in got["16"]], (label, got[label])
        for (_n, _k, f16), (_n2, _k2, f) in zip(got["16"], got[label]):
            assert abs(f - f16) < 0.02, (label, f16, f)
check("16-, 24- and 32-bit PCM and float32 WAVs of the same sound give the same notes, frames and pitch", t_bit_depths)

def t_gap():
    p, _n = write_wav("gap_d4.wav", [(293.66, 1.5), (None, 0.5), (293.66, 1.5)])
    caps = te.analyze_audio_file(p, te.TonerEngine())
    assert [(c["note"], c["n_frames"]) for c in caps] == [("D4", 97)], "[(c[\"note\"], c[\"n_frames\"]) for c in caps] == [(\"D4\", 97)]"
    p, _n = write_wav("gap_a3.wav", [(220.0, 1.5), (None, 0.5), (220.0, 1.5)])
    caps = te.analyze_audio_file(p, te.TonerEngine())
    assert [(c["note"], c["n_frames"]) for c in caps] == [("A3", 32), ("A3", 31)], "[(c[\"note\"], c[\"n_frames\"]) for c in caps] == [(\"A3\", 32), (\"A3\", 31)]"
check("same note across a silence gap: D4 joins into ONE capture (97 frames: the silent windows are not "
      "frames); A3 splits in two (a release window reads G3/G#3)", t_gap)


# ============================================================
# 5. TRANSPOSITION AND NOTE NAMES
# ============================================================
print("\n=== 5. Transposition and note names ===")

check("transpose_note: C4 Alto A4, A#2 Tenor C4, C4 Sopranino A3, C4 Bass D6, unknown type and '' unchanged",
      lambda: [te.transpose_note(n, s) for n, s in [("C4", "Alto"), ("A#2", "Tenor"), ("C4", "Sopranino"),
                                                    ("C4", "Bass"), ("C4", "Nope"), ("", "Alto"), ("Cx", "Alto")]]
      == ["A4", "C4", "A3", "D6", "C4", "", "Cx"])
check("reverse_transpose_note: A4 Alto -> C4, C4 Tenor -> A#2; round trip for every sax type",
      lambda: te.reverse_transpose_note("A4", "Alto") == "C4" and te.reverse_transpose_note("C4", "Tenor") == "A#2"
      and all(te.reverse_transpose_note(te.transpose_note("F#3", s), s) == "F#3" for s in te.SAX_TYPES))

def t_freq_midi():
    near(te.note_to_freq("A4"), 440.0)
    near(te.note_to_freq("C4"), 261.6255653005986)
    near(te.note_to_freq("A4", 442.0), 442.0)
    assert te.note_to_freq("H4") == 0.0 and te.note_to_freq("") == 0.0, "te.note_to_freq(\"H4\") == 0.0 and te.note_to_freq(\"\") == 0.0"
    assert te.note_name_to_midi("C#4") == 61 and te.note_name_to_midi("C-1") is None, "te.note_name_to_midi(\"C#4\") == 61 and te.note_name_to_midi(\"C-1\") is None"
    assert te.note_name_to_midi("Bb3") is None, "flats are not parsed"
    assert te.note_name_to_midi("A#10") is None, "two-digit octaves are not parsed"
check("note_to_freq A4 440 / C4 261.625565301 / ref 442; midi C#4 61; flats and 2-digit octaves -> None",
      t_freq_midi)

def t_cal_counts():
    counts = {s: len([n for n in te.CALIBRATION_NOTES
                      if te.note_to_freq(te.reverse_transpose_note(n, s)) >= te.MIN_FUNDAMENTAL_HZ])
              for s in te.SAX_TYPES}
    assert len(te.CALIBRATION_NOTES) == 32, "len(te.CALIBRATION_NOTES) == 32"
    assert counts == {"Bass": 28, "Baritone": 32, "Tenor": 32, "C Melody": 32, "Alto": 32,
                      "F Mezzo": 32, "Soprano": 32, "Sopranino": 32}, counts
check("CALIBRATION_NOTES >= MIN_FUNDAMENTAL_HZ (65 Hz): Bass 28, every other type 32", t_cal_counts)


print(f"\n{passed}/{passed + failed} passed  ({time.time() - _t0:.1f} s)")
sys.exit(0 if failed == 0 else 1)
