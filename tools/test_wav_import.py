"""Test script for WAV import — verifies _read_wav_file handles all WAV formats.

The stdlib `wave` module rejects format-tag 3 (IEEE float), but DAWs export
float WAVs by default. This suite synthesizes WAV files at the byte level for
each format we need to support and verifies _read_wav_file decodes them
correctly. Also smoke-tests analyze_audio_file end-to-end on a synthetic tone.
"""

import sys
import os
import struct
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from toner_engine import (
    _read_wav_file,
    analyze_audio_file,
    TonerEngine,
    SAMPLE_RATE,
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


# ================================================
# WAV byte-level writer (independent of production code)
# ================================================

# WAVE_FORMAT_IEEE_FLOAT SubFormat GUID
GUID_IEEE_FLOAT = bytes([
    0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00,
    0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71,
])
# WAVE_FORMAT_PCM SubFormat GUID
GUID_PCM = bytes([
    0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x00,
    0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71,
])


def write_wav(path, samples, framerate, n_channels, fmt_tag, bits,
              extensible=False, list_chunk=None):
    """Write a WAV file with arbitrary format tag.

    samples: 1-D numpy array in interleaved order (already cast to target dtype)
    fmt_tag: 1 = PCM, 3 = IEEE_FLOAT, 0xFFFE = EXTENSIBLE
    extensible: if True, the SubFormat GUID will be IEEE_FLOAT or PCM matching
                what the underlying samples actually are
    list_chunk: optional bytes to insert as a LIST chunk before the data chunk
                (simulates DAW metadata chunks like LIST/INFO)
    """
    sampwidth = bits // 8
    block_align = sampwidth * n_channels
    byte_rate = framerate * block_align

    if fmt_tag == 0xFFFE:
        # EXTENSIBLE: 40-byte fmt chunk
        guid = GUID_IEEE_FLOAT if extensible == 'float' else GUID_PCM
        fmt_chunk = struct.pack(
            '<HHIIHHHHI',
            0xFFFE, n_channels, framerate, byte_rate,
            block_align, bits,
            22,           # cbSize
            bits,         # valid bits per sample
            0x3,          # channel mask (front L + R)
        ) + guid
    else:
        fmt_chunk = struct.pack(
            '<HHIIHH',
            fmt_tag, n_channels, framerate, byte_rate,
            block_align, bits,
        )

    data_bytes = samples.tobytes()

    chunks = b''
    chunks += b'fmt ' + struct.pack('<I', len(fmt_chunk)) + fmt_chunk
    if list_chunk is not None:
        chunks += b'LIST' + struct.pack('<I', len(list_chunk)) + list_chunk
        if len(list_chunk) & 1:
            chunks += b'\x00'  # pad
    chunks += b'data' + struct.pack('<I', len(data_bytes)) + data_bytes

    riff = b'RIFF' + struct.pack('<I', 4 + len(chunks)) + b'WAVE' + chunks

    with open(path, 'wb') as f:
        f.write(riff)


def make_sine(freq, duration_s, framerate, amplitude=0.5, n_channels=1):
    """Generate a sine wave as float32 mono or stereo (interleaved)."""
    n_frames = int(duration_s * framerate)
    t = np.arange(n_frames, dtype=np.float64) / framerate
    mono = (amplitude * np.sin(2 * np.pi * freq * t)).astype(np.float32)
    if n_channels == 1:
        return mono
    # Interleaved stereo: copy mono into both channels
    out = np.empty(n_frames * n_channels, dtype=np.float32)
    for c in range(n_channels):
        out[c::n_channels] = mono
    return out


# ================================================
print("=== WAV Import Tests ===")
print()

# All format tests use a 1-second 440 Hz tone at the engine's sample rate.
TEST_FREQ = 440.0
TEST_DURATION = 1.0
TEST_FRAMERATE = SAMPLE_RATE

with tempfile.TemporaryDirectory() as tmpdir:

    # --- PCM 16-bit mono (regression: existing path) ---
    print("--- PCM int formats (regression) ---")
    f = make_sine(TEST_FREQ, TEST_DURATION, TEST_FRAMERATE, n_channels=1)
    pcm16 = (f * 32767).astype('<i2')
    path = os.path.join(tmpdir, 'pcm16_mono.wav')
    write_wav(path, pcm16, TEST_FRAMERATE, 1, fmt_tag=1, bits=16)
    result = _read_wav_file(path)
    test("PCM 16 mono: returned non-None", result is not None)
    if result:
        samples, fr, nc = result
        test("PCM 16 mono: framerate matches", fr == TEST_FRAMERATE)
        test("PCM 16 mono: channels = 1", nc == 1)
        test("PCM 16 mono: sample count matches",
             len(samples) == int(TEST_DURATION * TEST_FRAMERATE))
        test("PCM 16 mono: peak amplitude ~0.5",
             0.45 <= float(np.abs(samples).max()) <= 0.55)

    # --- PCM 24-bit stereo (less common, exercises mono mixdown) ---
    f_stereo = make_sine(TEST_FREQ, TEST_DURATION, TEST_FRAMERATE, n_channels=2)
    # Pack to 24-bit little-endian: take int32, drop the lowest byte
    pcm24_int = (f_stereo * (2**23 - 1)).astype('<i4')
    pcm24_bytes = bytearray()
    for v in pcm24_int:
        b = int(v).to_bytes(4, 'little', signed=True)
        pcm24_bytes.extend(b[:3])  # keep low 3 bytes (LE)
    pcm24_arr = np.frombuffer(bytes(pcm24_bytes), dtype=np.uint8)
    path = os.path.join(tmpdir, 'pcm24_stereo.wav')
    write_wav(path, pcm24_arr, TEST_FRAMERATE, 2, fmt_tag=1, bits=24)
    result = _read_wav_file(path)
    test("PCM 24 stereo: returned non-None", result is not None)
    if result:
        samples, fr, nc = result
        test("PCM 24 stereo: channels = 2", nc == 2)
        test("PCM 24 stereo: interleaved sample count correct",
             len(samples) == int(TEST_DURATION * TEST_FRAMERATE) * 2)

    # --- PCM 32-bit int mono ---
    pcm32 = (f * (2**31 - 1)).astype('<i4')
    path = os.path.join(tmpdir, 'pcm32_mono.wav')
    write_wav(path, pcm32, TEST_FRAMERATE, 1, fmt_tag=1, bits=32)
    result = _read_wav_file(path)
    test("PCM 32 int mono: returned non-None", result is not None)
    if result:
        samples, _, _ = result
        test("PCM 32 int mono: peak amplitude ~0.5",
             0.45 <= float(np.abs(samples).max()) <= 0.55)

    print()
    print("--- IEEE float formats (the bugfix) ---")

    # --- IEEE float 32-bit mono (THE BUG FIX) ---
    path = os.path.join(tmpdir, 'float32_mono.wav')
    write_wav(path, f, TEST_FRAMERATE, 1, fmt_tag=3, bits=32)
    result = _read_wav_file(path)
    test("Float32 mono: returned non-None (was crashing before fix)",
         result is not None)
    if result:
        samples, fr, nc = result
        test("Float32 mono: framerate matches", fr == TEST_FRAMERATE)
        test("Float32 mono: channels = 1", nc == 1)
        test("Float32 mono: sample count matches",
             len(samples) == int(TEST_DURATION * TEST_FRAMERATE))
        test("Float32 mono: peak amplitude ~0.5",
             0.45 <= float(np.abs(samples).max()) <= 0.55)
        # Float WAV should round-trip exactly
        test("Float32 mono: samples round-trip exactly",
             np.allclose(samples, f, atol=1e-7))

    # --- IEEE float 32-bit STEREO (mono mixdown on float path) ---
    path = os.path.join(tmpdir, 'float32_stereo.wav')
    write_wav(path, f_stereo, TEST_FRAMERATE, 2, fmt_tag=3, bits=32)
    result = _read_wav_file(path)
    test("Float32 stereo: returned non-None", result is not None)
    if result:
        samples, fr, nc = result
        test("Float32 stereo: channels = 2", nc == 2)
        test("Float32 stereo: interleaved sample count correct",
             len(samples) == int(TEST_DURATION * TEST_FRAMERATE) * 2)

    # --- IEEE float 64-bit mono ---
    f64 = f.astype('<f8')
    path = os.path.join(tmpdir, 'float64_mono.wav')
    write_wav(path, f64, TEST_FRAMERATE, 1, fmt_tag=3, bits=64)
    result = _read_wav_file(path)
    test("Float64 mono: returned non-None", result is not None)
    if result:
        samples, _, _ = result
        test("Float64 mono: peak amplitude ~0.5",
             0.45 <= float(np.abs(samples).max()) <= 0.55)

    print()
    print("--- WAVE_FORMAT_EXTENSIBLE (Pro Tools / Reaper exports) ---")

    # --- EXTENSIBLE wrapping IEEE float32 ---
    path = os.path.join(tmpdir, 'extensible_float.wav')
    write_wav(path, f, TEST_FRAMERATE, 1, fmt_tag=0xFFFE, bits=32,
              extensible='float')
    result = _read_wav_file(path)
    test("Extensible-float32 mono: returned non-None", result is not None)
    if result:
        samples, _, _ = result
        test("Extensible-float32 mono: samples round-trip",
             np.allclose(samples, f, atol=1e-7))

    # --- EXTENSIBLE wrapping PCM16 ---
    path = os.path.join(tmpdir, 'extensible_pcm.wav')
    write_wav(path, pcm16, TEST_FRAMERATE, 1, fmt_tag=0xFFFE, bits=16,
              extensible='pcm')
    result = _read_wav_file(path)
    test("Extensible-PCM16 mono: returned non-None", result is not None)

    print()
    print("--- Chunk parsing edge cases ---")

    # --- LIST chunk before data (common DAW metadata) ---
    list_payload = b'INFOICMT\x10\x00\x00\x00bounced from DAW'
    path = os.path.join(tmpdir, 'with_list_chunk.wav')
    write_wav(path, f, TEST_FRAMERATE, 1, fmt_tag=3, bits=32,
              list_chunk=list_payload)
    result = _read_wav_file(path)
    test("LIST chunk before data: parser skips it", result is not None)
    if result:
        samples, _, _ = result
        test("LIST chunk: data still parsed correctly",
             len(samples) == int(TEST_DURATION * TEST_FRAMERATE))

    # --- Non-RIFF file should return None, not crash ---
    path = os.path.join(tmpdir, 'not_a_wav.wav')
    with open(path, 'wb') as fp:
        fp.write(b'this is not a WAV file at all')
    result = _read_wav_file(path)
    test("Non-RIFF input: returns None gracefully", result is None)

    print()
    print("--- End-to-end: analyze_audio_file on a float WAV ---")

    # Synthesize a 220 Hz (~A3) tone with harmonics so the engine has
    # something to detect, save as float32 WAV, and verify analyze_audio_file
    # produces at least one capture.
    duration = 3.0
    n = int(duration * TEST_FRAMERATE)
    t = np.arange(n, dtype=np.float64) / TEST_FRAMERATE
    fund = 220.0
    tone = (
        0.5 * np.sin(2 * np.pi * fund * t)
        + 0.25 * np.sin(2 * np.pi * 2 * fund * t)
        + 0.15 * np.sin(2 * np.pi * 3 * fund * t)
        + 0.08 * np.sin(2 * np.pi * 4 * fund * t)
    ).astype(np.float32)
    path = os.path.join(tmpdir, 'tone_float32.wav')
    write_wav(path, tone, TEST_FRAMERATE, 1, fmt_tag=3, bits=32)

    engine = TonerEngine()
    engine.sax_type = "Tenor"
    captures = analyze_audio_file(path, engine)
    test("analyze_audio_file on float32 WAV: at least one capture",
         len(captures) >= 1)
    if captures:
        # Frequency should be near 220 Hz (within a few Hz tolerance)
        freqs = [c.get('fundamental_freq', 0) for c in captures]
        avg_freq = sum(freqs) / len(freqs)
        test(f"analyze_audio_file: detected freq ~220 Hz (got {avg_freq:.1f})",
             215 <= avg_freq <= 225)

print()
print("=" * 50)
print(f"Results: {passed} passed, {failed} failed out of {passed + failed}")
sys.exit(0 if failed == 0 else 1)
