"""
Non-interactive test suite for AudioRingBuffer from audio_utils.py.

Tests: basic write/read, ring buffer wrapping, thread safety,
stale detection, empty buffer, partial writes.

Usage:
    python tools/test_audio_utils.py
"""

import os
import sys
import threading
import time

# Add parent directory to path so we can import project modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
from audio_utils import AudioRingBuffer


# =============================================================================
# TEST FRAMEWORK
# =============================================================================

_results = []

def run_test(name, fn):
    """Run a test function and record PASS/FAIL."""
    try:
        fn()
        _results.append(("PASS", name, None))
        print(f"  PASS  {name}")
    except AssertionError as e:
        _results.append(("FAIL", name, str(e)))
        print(f"  FAIL  {name} -- {e}")
    except Exception as e:
        _results.append(("ERROR", name, str(e)))
        print(f"  ERROR {name} -- {type(e).__name__}: {e}")


def assert_true(condition, msg=""):
    if not condition:
        raise AssertionError(msg or "Expected True but got False")

def assert_false(condition, msg=""):
    if condition:
        raise AssertionError(msg or "Expected False but got True")

def assert_equal(a, b, msg=""):
    if a != b:
        raise AssertionError(msg or f"Expected {a!r} == {b!r}")


# =============================================================================
# BASIC WRITE/READ
# =============================================================================

def test_basic_write_read():
    """Write data into buffer, read it back, verify match."""
    buf = AudioRingBuffer(10)
    data = np.array([1, 2, 3, 4, 5], dtype=np.float32)
    buf.write(data)
    result = buf.read()
    assert_true(result is not None, "read() returned None after write")
    # Buffer is size 10, wrote 5 samples. The read should return all 10
    # with zeros for the unwritten portion (oldest) and data at the end.
    assert_equal(len(result), 10, f"Expected length 10, got {len(result)}")
    # The last 5 elements should be our data (chronological order)
    np.testing.assert_array_equal(result[5:], data)
    # The first 5 should be zeros
    np.testing.assert_array_equal(result[:5], np.zeros(5, dtype=np.float32))


def test_basic_write_exact_size():
    """Write exactly buffer-size amount of data."""
    buf = AudioRingBuffer(5)
    data = np.array([10, 20, 30, 40, 50], dtype=np.float32)
    buf.write(data)
    result = buf.read()
    assert_true(result is not None, "read() returned None after write")
    np.testing.assert_array_equal(result, data)


def test_multiple_writes():
    """Multiple sequential writes, read returns all in order."""
    buf = AudioRingBuffer(8)
    buf.write(np.array([1, 2, 3], dtype=np.float32))
    buf.write(np.array([4, 5, 6], dtype=np.float32))
    result = buf.read()
    assert_true(result is not None)
    # Wrote 6 samples total into size-8 buffer. First 2 are zeros.
    expected = np.array([0, 0, 1, 2, 3, 4, 5, 6], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


# =============================================================================
# RING BUFFER WRAPPING
# =============================================================================

def test_wrap_around():
    """Write more than buffer size, verify only most recent data kept."""
    buf = AudioRingBuffer(5)
    buf.write(np.array([1, 2, 3], dtype=np.float32))
    buf.write(np.array([4, 5, 6], dtype=np.float32))
    # Total 6 written, buffer size 5. Most recent 5 should be [2,3,4,5,6].
    result = buf.read()
    assert_true(result is not None)
    # After first write: [1,2,3,0,0], write_pos=3
    # After second write: [1,2,3,4,5] then wraps: [6,2,3,4,5], write_pos=1
    # read() does np.roll(buffer, -write_pos) = roll([6,2,3,4,5], -1) = [2,3,4,5,6]
    expected = np.array([2, 3, 4, 5, 6], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


def test_write_larger_than_buffer():
    """Write a chunk bigger than the buffer in one call."""
    buf = AudioRingBuffer(4)
    data = np.array([1, 2, 3, 4, 5, 6, 7], dtype=np.float32)
    buf.write(data)
    result = buf.read()
    assert_true(result is not None)
    # When n >= buffer size, it takes data[-4:] = [4,5,6,7], write_pos=0
    expected = np.array([4, 5, 6, 7], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


def test_multiple_wraps():
    """Write many small chunks that wrap around multiple times."""
    buf = AudioRingBuffer(4)
    for i in range(10):
        buf.write(np.array([float(i)], dtype=np.float32))
    result = buf.read()
    assert_true(result is not None)
    # Last 4 writes: 6, 7, 8, 9
    expected = np.array([6, 7, 8, 9], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


# =============================================================================
# THREAD SAFETY
# =============================================================================

def test_thread_safety():
    """Write from one thread, read from another, no crashes."""
    buf = AudioRingBuffer(1024)
    errors = []
    stop_event = threading.Event()

    def writer():
        try:
            i = 0
            while not stop_event.is_set():
                chunk = np.full(64, float(i % 100), dtype=np.float32)
                buf.write(chunk)
                i += 1
        except Exception as e:
            errors.append(f"Writer: {e}")

    def reader():
        try:
            while not stop_event.is_set():
                result = buf.read()
                if result is not None:
                    assert_equal(len(result), 1024,
                                 f"Read returned wrong length: {len(result)}")
        except Exception as e:
            errors.append(f"Reader: {e}")

    w = threading.Thread(target=writer)
    r = threading.Thread(target=reader)
    w.start()
    r.start()

    time.sleep(0.3)
    stop_event.set()
    w.join(timeout=2)
    r.join(timeout=2)

    assert_true(len(errors) == 0, f"Thread errors: {errors}")


def test_concurrent_writers():
    """Multiple writers, one reader, no crashes."""
    buf = AudioRingBuffer(512)
    errors = []
    stop_event = threading.Event()

    def writer(thread_id):
        try:
            while not stop_event.is_set():
                chunk = np.full(32, float(thread_id), dtype=np.float32)
                buf.write(chunk)
        except Exception as e:
            errors.append(f"Writer {thread_id}: {e}")

    def reader():
        try:
            while not stop_event.is_set():
                buf.read()
        except Exception as e:
            errors.append(f"Reader: {e}")

    threads = [threading.Thread(target=writer, args=(i,)) for i in range(4)]
    threads.append(threading.Thread(target=reader))
    for t in threads:
        t.start()

    time.sleep(0.3)
    stop_event.set()
    for t in threads:
        t.join(timeout=2)

    assert_true(len(errors) == 0, f"Thread errors: {errors}")


# =============================================================================
# STALE DETECTION
# =============================================================================

def test_stale_before_any_activity():
    """Fresh buffer: is_stale should be True (write_count == last_read_count == 0)."""
    buf = AudioRingBuffer(10)
    assert_true(buf.is_stale(), "Fresh buffer should be stale (no writes since last read)")


def test_not_stale_after_write():
    """After writing, buffer should not be stale (write_count > last_read_count)."""
    buf = AudioRingBuffer(10)
    buf.write(np.array([1.0], dtype=np.float32))
    assert_false(buf.is_stale(), "Should not be stale after write without read")


def test_stale_after_read():
    """Write, then read. Now stale because no new writes since last read."""
    buf = AudioRingBuffer(10)
    buf.write(np.array([1.0], dtype=np.float32))
    buf.read()
    assert_true(buf.is_stale(), "Should be stale after read catches up to writes")


def test_not_stale_after_new_write():
    """Write, read (stale), write again (not stale)."""
    buf = AudioRingBuffer(10)
    buf.write(np.array([1.0], dtype=np.float32))
    buf.read()
    assert_true(buf.is_stale())
    buf.write(np.array([2.0], dtype=np.float32))
    assert_false(buf.is_stale(), "Should not be stale after new write")


def test_stale_write_count_mechanism():
    """Verify the write_count increments and last_read_count tracks it."""
    buf = AudioRingBuffer(10)
    assert_equal(buf.write_count, 0)
    assert_equal(buf.last_read_count, 0)

    buf.write(np.array([1.0], dtype=np.float32))
    assert_equal(buf.write_count, 1)
    assert_equal(buf.last_read_count, 0)

    buf.write(np.array([2.0], dtype=np.float32))
    assert_equal(buf.write_count, 2)
    assert_equal(buf.last_read_count, 0)

    buf.read()
    assert_equal(buf.last_read_count, 2)

    buf.write(np.array([3.0], dtype=np.float32))
    assert_equal(buf.write_count, 3)
    assert_equal(buf.last_read_count, 2)
    assert_false(buf.is_stale())


# =============================================================================
# EMPTY BUFFER
# =============================================================================

def test_empty_buffer_returns_none():
    """Reading from a fresh buffer returns None."""
    buf = AudioRingBuffer(10)
    result = buf.read()
    assert_true(result is None, f"Expected None from empty buffer, got {type(result)}")


def test_cleared_buffer_returns_none():
    """After clear(), read returns None."""
    buf = AudioRingBuffer(10)
    buf.write(np.array([1, 2, 3], dtype=np.float32))
    buf.clear()
    result = buf.read()
    assert_true(result is None, "Expected None after clear()")


def test_clear_resets_counters():
    """clear() resets write_count and last_read_count."""
    buf = AudioRingBuffer(10)
    buf.write(np.array([1.0], dtype=np.float32))
    buf.read()
    buf.clear()
    assert_equal(buf.write_count, 0)
    assert_equal(buf.last_read_count, 0)
    assert_equal(buf.write_pos, 0)
    assert_false(buf.has_data)


# =============================================================================
# PARTIAL WRITES
# =============================================================================

def test_partial_write_zero_padded():
    """Write less than full buffer, verify zeros fill the rest."""
    buf = AudioRingBuffer(8)
    buf.write(np.array([10, 20, 30], dtype=np.float32))
    result = buf.read()
    assert_true(result is not None)
    # First 5 elements are zeros (oldest), last 3 are data
    expected = np.array([0, 0, 0, 0, 0, 10, 20, 30], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


def test_single_sample_write():
    """Write a single sample."""
    buf = AudioRingBuffer(4)
    buf.write(np.array([42.0], dtype=np.float32))
    result = buf.read()
    assert_true(result is not None)
    expected = np.array([0, 0, 0, 42], dtype=np.float32)
    np.testing.assert_array_equal(result, expected)


def test_read_returns_copy():
    """Verify read() returns a copy, not a view into the internal buffer."""
    buf = AudioRingBuffer(4)
    buf.write(np.array([1, 2, 3, 4], dtype=np.float32))
    result = buf.read()
    result[:] = 0  # Modify the returned array
    result2 = buf.read()
    expected = np.array([1, 2, 3, 4], dtype=np.float32)
    np.testing.assert_array_equal(result2, expected)


# =============================================================================
# MAIN
# =============================================================================

if __name__ == "__main__":
    print("\n=== AudioRingBuffer Test Suite ===\n")

    print("-- Basic Write/Read --")
    run_test("basic_write_read", test_basic_write_read)
    run_test("basic_write_exact_size", test_basic_write_exact_size)
    run_test("multiple_writes", test_multiple_writes)

    print("\n-- Ring Buffer Wrapping --")
    run_test("wrap_around", test_wrap_around)
    run_test("write_larger_than_buffer", test_write_larger_than_buffer)
    run_test("multiple_wraps", test_multiple_wraps)

    print("\n-- Thread Safety --")
    run_test("thread_safety", test_thread_safety)
    run_test("concurrent_writers", test_concurrent_writers)

    print("\n-- Stale Detection --")
    run_test("stale_before_any_activity", test_stale_before_any_activity)
    run_test("not_stale_after_write", test_not_stale_after_write)
    run_test("stale_after_read", test_stale_after_read)
    run_test("not_stale_after_new_write", test_not_stale_after_new_write)
    run_test("stale_write_count_mechanism", test_stale_write_count_mechanism)

    print("\n-- Empty Buffer --")
    run_test("empty_buffer_returns_none", test_empty_buffer_returns_none)
    run_test("cleared_buffer_returns_none", test_cleared_buffer_returns_none)
    run_test("clear_resets_counters", test_clear_resets_counters)

    print("\n-- Partial Writes --")
    run_test("partial_write_zero_padded", test_partial_write_zero_padded)
    run_test("single_sample_write", test_single_sample_write)
    run_test("read_returns_copy", test_read_returns_copy)

    # Summary
    passed = sum(1 for r in _results if r[0] == "PASS")
    failed = sum(1 for r in _results if r[0] == "FAIL")
    errors = sum(1 for r in _results if r[0] == "ERROR")
    total = len(_results)

    print(f"\n{'='*40}")
    print(f"  {passed}/{total} PASSED", end="")
    if failed:
        print(f", {failed} FAILED", end="")
    if errors:
        print(f", {errors} ERRORS", end="")
    print()
    print(f"{'='*40}\n")

    sys.exit(1 if (failed + errors) > 0 else 0)
