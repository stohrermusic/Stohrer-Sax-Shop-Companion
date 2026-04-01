"""Test script for WAV recording — verifies save_recording produces valid files."""

import sys
import os
import tempfile
import wave
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import TonerEngine, SAMPLE_RATE

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


# ================================================
print("=== WAV Recording Tests ===")
print()

# --- Test 1: save_recording produces a valid WAV file ---
print("--- save_recording basics ---")
with tempfile.TemporaryDirectory() as tmpdir:
    filepath = os.path.join(tmpdir, "test.wav")
    # Simulate 1 second of audio in chunks (like the real callback)
    chunk_size = 1024
    n_chunks = SAMPLE_RATE // chunk_size
    chunks = [np.random.uniform(-0.5, 0.5, chunk_size).astype(np.float32)
              for _ in range(n_chunks)]

    TonerEngine.save_recording(chunks, filepath)

    test("WAV file created", os.path.isfile(filepath))
    test("WAV file non-empty", os.path.getsize(filepath) > 0)

    # Verify it's a valid WAV
    with wave.open(filepath, 'rb') as wf:
        test("channels = 1", wf.getnchannels() == 1)
        test("sample width = 2 (16-bit)", wf.getsampwidth() == 2)
        test("sample rate correct", wf.getframerate() == SAMPLE_RATE)
        expected_frames = chunk_size * n_chunks
        test("frame count matches input", wf.getnframes() == expected_frames)

print()

# --- Test 2: save_recording with empty chunks does nothing ---
print("--- save_recording with empty/None ---")
with tempfile.TemporaryDirectory() as tmpdir:
    filepath = os.path.join(tmpdir, "empty.wav")
    TonerEngine.save_recording([], filepath)
    test("empty chunks -> no file", not os.path.isfile(filepath))

    TonerEngine.save_recording(None, filepath)
    test("None chunks -> no file", not os.path.isfile(filepath))

print()

# --- Test 3: start/stop recording lifecycle ---
print("--- start/stop recording lifecycle ---")
engine = TonerEngine.__new__(TonerEngine)
engine._recording_chunks = None

# Before start, stop returns None
result = engine.stop_recording()
test("stop before start returns None", result is None)

# Start then stop with no data
engine.start_recording()
test("after start, _recording_chunks is list", isinstance(engine._recording_chunks, list))
result = engine.stop_recording()
test("stop returns empty list (no audio)", result == [])
test("after stop, _recording_chunks is None", engine._recording_chunks is None)

# Start, simulate callbacks, stop
engine.start_recording()
fake_indata = np.zeros((1024, 1), dtype=np.float32)
# Simulate the audio callback appending data
engine._recording_chunks.append(fake_indata[:, 0].copy())
engine._recording_chunks.append(fake_indata[:, 0].copy())
result = engine.stop_recording()
test("stop returns 2 chunks", len(result) == 2)

print()

# --- Test 4: save with path containing spaces and special chars ---
print("--- save with tricky file paths ---")
with tempfile.TemporaryDirectory() as tmpdir:
    tricky_dir = os.path.join(tmpdir, "My Music", "Sax Recordings")
    os.makedirs(tricky_dir)
    filepath = os.path.join(tricky_dir, "Conn Virtuoso_2026-03-30_14-30-00.wav")
    chunks = [np.zeros(1024, dtype=np.float32)]
    TonerEngine.save_recording(chunks, filepath)
    test("file with spaces in path created", os.path.isfile(filepath))

print()

# --- Test 5: _toner_get_recording_dir default logic ---
print("--- default recording dir logic ---")

# Simulate the default dir computation (same logic as _toner_get_recording_dir)
home = os.path.expanduser('~')
music = os.path.join(home, 'Music')
if not os.path.isdir(music):
    music = os.path.join(home, 'Documents')
expected_default = os.path.join(music, 'StohrerSaxShopCompanion')

print(f"  Home dir: {home}")
print(f"  Expected default recording dir: {expected_default}")
test("home dir exists", os.path.isdir(home))

# Check that the fallback parent exists
parent_exists = os.path.isdir(os.path.join(home, 'Music')) or os.path.isdir(os.path.join(home, 'Documents'))
test("Music or Documents folder exists", parent_exists)

# Actually create the dir to verify makedirs works
with tempfile.TemporaryDirectory() as tmpdir:
    # Simulate with a fake home
    fake_music = os.path.join(tmpdir, "Music")
    os.makedirs(fake_music)
    rec_dir = os.path.join(fake_music, "StohrerSaxShopCompanion")
    os.makedirs(rec_dir, exist_ok=True)
    test("makedirs creates recording subdir", os.path.isdir(rec_dir))

    # Test saving a file there
    filepath = os.path.join(rec_dir, "test_session.wav")
    chunks = [np.zeros(1024, dtype=np.float32)]
    TonerEngine.save_recording(chunks, filepath)
    test("WAV saved in created recording dir", os.path.isfile(filepath))

print()

# --- Test 6: verify _toner_save_wav_recording skips on empty chunks ---
print("--- empty chunk edge cases ---")
# The condition in _toner_stop_capture is:
#   if chunks and self._toner_active_session:
# An empty list [] is falsy, so save is skipped
chunks_empty = []
test("empty list is falsy (save skipped)", not chunks_empty)

# A list with one empty array
chunks_one_empty = [np.array([], dtype=np.float32)]
test("list with empty array is truthy (save attempted)", bool(chunks_one_empty))

# save_recording handles this gracefully (concatenate of empty arrays)
with tempfile.TemporaryDirectory() as tmpdir:
    filepath = os.path.join(tmpdir, "edge.wav")
    TonerEngine.save_recording(chunks_one_empty, filepath)
    # np.concatenate of [empty_array] gives empty array, wave writes 0 frames
    if os.path.isfile(filepath):
        with wave.open(filepath, 'rb') as wf:
            test("empty-array chunk produces 0-frame WAV", wf.getnframes() == 0)
    else:
        test("empty-array chunk handled (no crash)", True)

print()

# ================================================
print("=" * 50)
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
if failed:
    sys.exit(1)
