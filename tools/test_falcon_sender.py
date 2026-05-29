"""
Engine-level tests for falcon_sender.py — no real hardware required.

Uses a MockSerial that mimics a Grbl 1.1f controller (emits the boot
banner, answers status polls, ACKs lines, can be made to emit errors
or alarms). All tests run on the worker thread machinery the real
sender uses, so streaming/threading bugs are caught here.
"""

import sys
import os
import time
import threading

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import falcon_sender as fs

passed = 0
failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


def stream_lines_only(completed):
    """Strip the streamer's pre-stream safety lines (wake-up dwell and
    $X unlock) so tests can compare against the caller's intended
    stream lines."""
    return [ln for ln in completed if ln not in ('G4 P0.01', '$X')]


if not fs.HAS_PYSERIAL:
    print("\n!! pyserial not available - skipping falcon_sender tests")
    sys.exit(0)


# ============================================================
print("\n=== Status parser ===")
# ============================================================

s = fs.parse_status('<Idle|MPos:0.000,0.000,0.000|FS:0,0>')
check("Idle status state", s and s['state'] == 'Idle')
check("Idle status MPos", s and s['mpos'] == [0.0, 0.0, 0.0])
check("Idle status feed/spindle", s and s['feed'] == 0 and s['spindle'] == 0)

s = fs.parse_status('<Run|MPos:5.529,0.560,0.000|FS:200,500|Pn:S|Ov:100,100,100>')
check("Run status state", s and s['state'] == 'Run')
check("Run status MPos", s and s['mpos'] == [5.529, 0.560, 0.0])
check("Run status feed=200 spindle=500", s and s['feed'] == 200 and s['spindle'] == 500)
check("Run status pins", s and s.get('pins') == 'S')
check("Run status overrides", s and s.get('overrides') == [100, 100, 100])

s = fs.parse_status('<Alarm|MPos:0.000,0.000,0.000|WCO:0.000,0.000,0.000>')
check("Alarm status state", s and s['state'] == 'Alarm')
check("Alarm status WCO", s and s['wco'] == [0.0, 0.0, 0.0])

s = fs.parse_status('garbage')
check("non-status line returns None", s is None)


# ============================================================
print("\n=== list_serial_ports / auto_detect_falcon (smoke) ===")
# ============================================================

# These can't be meaningfully unit-tested without real hardware, but they
# shouldn't crash and should return reasonable types.
ports = fs.list_serial_ports()
check("list_serial_ports returns a list", isinstance(ports, list))
for p in ports:
    if 'device' not in p or 'description' not in p:
        check(f"port dict has device + description (got {p.keys()})", False)
        break
else:
    check("each port dict has device + description", True)

# auto_detect_falcon — if hardware is present, may return a string; else None
result = fs.auto_detect_falcon()
check("auto_detect_falcon returns str or None",
      result is None or isinstance(result, str))


# ============================================================
print("\n=== MockSerial-backed streaming ===")
# ============================================================

class MockSerial:
    """In-memory pseudo-serial that mimics a Grbl 1.1f controller.

    Buffers writes from the host. On each read(), processes pending writes
    and produces appropriate responses:
        - 'Grbl 1.1f ...' banner on first read after open or soft-reset
        - 'ok\\n' for each completed line
        - '<state|...>' for each '?' real-time byte
        - 'error:N\\n' or 'ALARM:N\\n' if explicitly programmed
    """
    def __init__(self):
        self.is_open = True
        self.in_buf = bytearray()       # what the host has written (from sender)
        self.out_buf = bytearray()       # what the controller will send back
        self.banner_pending = True
        self.completed_lines = []
        self.pending_alarm = None
        self.pending_error = None
        self._cur_state = 'Idle'
        self._mpos = [0.0, 0.0, 0.0]
        self._lock = threading.Lock()

    @property
    def in_waiting(self):
        with self._lock:
            self._process_input()
            return len(self.out_buf)

    def read(self, n):
        with self._lock:
            self._process_input()
            if self.banner_pending and not self.out_buf:
                self.out_buf.extend(b'\r\nGrbl 1.1f [\'$\' for help]\r\n')
                self.banner_pending = False
            data = bytes(self.out_buf[:n])
            del self.out_buf[:n]
            return data

    def write(self, data):
        with self._lock:
            # Handle real-time commands inline (don't queue them)
            for b in data:
                bb = bytes([b])
                if bb == fs.RT_SOFT_RESET:
                    self.in_buf.clear()
                    self.completed_lines.clear()
                    self.banner_pending = True
                elif bb == fs.RT_STATUS:
                    self.out_buf.extend(
                        f'<{self._cur_state}|MPos:'
                        f'{self._mpos[0]:.3f},{self._mpos[1]:.3f},{self._mpos[2]:.3f}'
                        f'|FS:0,0>\r\n'.encode('ascii')
                    )
                elif bb in (fs.RT_FEED_HOLD, fs.RT_CYCLE_START):
                    pass  # ignored for tests
                else:
                    self.in_buf.append(b)
            return len(data)

    def _process_input(self):
        # Pull any complete lines out of in_buf and ack them
        while b'\n' in self.in_buf:
            idx = self.in_buf.index(b'\n')
            line = bytes(self.in_buf[:idx]).decode('ascii', 'replace').strip()
            del self.in_buf[:idx + 1]
            if not line:
                self.out_buf.extend(b'ok\r\n')
                continue
            self.completed_lines.append(line)
            # FalconSender's stream prefix: a wake-up dwell + a $X
            # safety unlock. Both are implementation details (not
            # test-controlled stream lines), so always ack them and
            # reserve pending_error/pending_alarm for the caller's
            # actual stream commands.
            if line in ('G4 P0.01', '$X'):
                self.out_buf.extend(b'ok\r\n')
                continue
            if self.pending_error:
                self.out_buf.extend((self.pending_error + '\r\n').encode('ascii'))
                self.pending_error = None
            elif self.pending_alarm:
                self.out_buf.extend((self.pending_alarm + '\r\n').encode('ascii'))
                self.pending_alarm = None
                self._cur_state = 'Alarm'
            else:
                self.out_buf.extend(b'ok\r\n')

    def reset_input_buffer(self):
        with self._lock:
            self.out_buf.clear()

    def close(self):
        self.is_open = False


def _new_sender_with_mock():
    """Build a FalconSender backed by a MockSerial. Bypasses connect()."""
    mock = MockSerial()
    s = fs.FalconSender(port='MOCK', baud=115200)
    s._serial = mock
    s.status = {'state': 'Idle'}
    return s, mock


# --- Streaming basics ---
sender, mock = _new_sender_with_mock()
done_reasons = []
progress_calls = []
sender.on_done = lambda reason: done_reasons.append(reason)
sender.on_progress = lambda i, n: progress_calls.append((i, n))

lines = ['G90', 'G0X10', 'G0X20', 'G0X30', 'M2']
sender.start_stream(lines)
# Wait for the worker to finish
sender._worker.join(timeout=5.0)

check("worker thread exited", not sender._worker.is_alive())
check(f"reason was 'complete' (got {done_reasons})",
      done_reasons == ['complete'])
check(f"all {len(lines)} lines made it to mock controller",
      stream_lines_only(mock.completed_lines) == lines)
check("progress callbacks fired for each line",
      len(progress_calls) == len(lines)
      and progress_calls[-1] == (len(lines), len(lines)))
# Every stream sends a wake-up dwell and a $X unlock before the caller's
# G-code, so a sticky Alarm from a previous run doesn't force a Falcon
# restart. Verify both prefix lines were sent (in that order) ahead of
# the user's first G-code line.
check("wake-up dwell sent before user G-code",
      'G4 P0.01' in mock.completed_lines
      and mock.completed_lines.index('G4 P0.01') < mock.completed_lines.index('G90'))
check("$X unlock sent before user G-code",
      '$X' in mock.completed_lines
      and mock.completed_lines.index('$X') < mock.completed_lines.index('G90'))


# --- error mid-stream stops cleanly ---
sender, mock = _new_sender_with_mock()
mock.pending_error = 'error:9'   # G-code locked out during alarm or jog
errors = []
done_reasons = []
sender.on_error = lambda msg: errors.append(msg)
sender.on_done = lambda reason: done_reasons.append(reason)

sender.start_stream(['G0X10', 'G0X20', 'G0X30'])
sender._worker.join(timeout=5.0)

check("error mid-stream surfaced via on_error",
      any('error:9' in e for e in errors))
check("error mid-stream ended with reason 'error'",
      done_reasons == ['error'])
check("latest_error stored",
      sender.latest_error and 'error:9' in sender.latest_error)


# --- alarm mid-stream ---
sender, mock = _new_sender_with_mock()
mock.pending_alarm = 'ALARM:1'   # hard limit
alarms = []
done_reasons = []
sender.on_alarm = lambda msg: alarms.append(msg)
sender.on_done = lambda reason: done_reasons.append(reason)

sender.start_stream(['G0X10', 'G0X20'])
sender._worker.join(timeout=5.0)

check("alarm surfaced via on_alarm",
      any('ALARM:1' in a for a in alarms))
check("alarm sets reason='error'", done_reasons == ['error'])
check("latest_alarm stored", sender.latest_alarm == 'ALARM:1')


# --- stop() halts cleanly ---
sender, mock = _new_sender_with_mock()
done_reasons = []
sender.on_done = lambda reason: done_reasons.append(reason)

# Use many lines so stop() lands mid-stream
long_lines = [f'G0X{i}' for i in range(500)]
sender.start_stream(long_lines)
worker_ref = sender._worker  # stop() nulls _worker
time.sleep(0.05)
sender.stop()

check("stop() joined the worker", not worker_ref.is_alive())
# Either 'stopped' (top-of-loop check fired) or 'error' (soft-reset
# triggered a serial error in the middle of a write). Both are acceptable
# outcomes of an emergency stop — the important guarantee is that the
# stream was halted before completion.
check(f"stop() reports stopped or error (got {done_reasons})",
      done_reasons == ['stopped'] or done_reasons == ['error'])
check("stop() didn't send the full set",
      len(mock.completed_lines) < len(long_lines))


# --- Status callbacks fire ---
sender, mock = _new_sender_with_mock()
status_updates = []
sender.on_status = lambda s: status_updates.append(s)

sender.start_stream(['G0X10', 'G0X20', 'G0X30'])
sender._worker.join(timeout=5.0)

check("status callbacks fired during streaming",
      len(status_updates) >= 1)
check("status callbacks parse state correctly",
      all('state' in s for s in status_updates))


# --- Jog during stream doesn't corrupt the character-counting protocol ---
# Regression for the "error:8 — '$' command only valid when idle" case
# where mid-stream jog writes interleaved with stream writes on the wire
# AND jog acks got consumed as stream-line acks. With the fix, jog calls
# serialize via the same lock as stream writes, and jog `ok`s are tracked
# in _jogs_in_flight so they don't drain sent_lengths early.

class MockSerialNoCollide(MockSerial):
    """MockSerial that tracks atomicity: each write that touches in_buf
    must NOT interleave another's bytes. We verify this by checking
    that every newline-terminated line in in_buf is a complete G-code
    or jog command (no partial lines starting mid-token)."""
    def __init__(self):
        super().__init__()
        self.write_count = 0
        self.write_bytes_log = []  # what got written, in call order

    def write(self, data):
        with self._lock:
            self.write_count += 1
            self.write_bytes_log.append(bytes(data))
        return super().write(data)


sender, mock = _new_sender_with_mock()
# Promote mock to the tracking variant
tracking_mock = MockSerialNoCollide()
sender._serial = tracking_mock
sender.status = {'state': 'Idle'}
done_reasons = []
sender.on_done = lambda reason: done_reasons.append(reason)

# Fire off a stream + interleave jogs from "another thread" (simulated
# by calling jog() between waiting for the stream to complete).
stream_lines = [f'G0X{i}' for i in range(20)]
sender.start_stream(stream_lines)

# Interleave 10 jogs while the stream is running
for _ in range(10):
    sender.jog(x=1, y=0, feed=1000, relative=True)
sender._worker.join(timeout=5.0)

check("jog+stream worker exited", not sender._worker.is_alive())
check(f"jog+stream completes successfully (got {done_reasons})",
      done_reasons == ['complete'])
# Every write should have been a complete line (no interleaving bytes).
# A garbled write would show as a chunk lacking '\n' or merging two
# commands.
garbled = [b for b in tracking_mock.write_bytes_log
           if b'\n' in b and b.count(b'\n') > 1
           and not b.endswith(b'\n')]
check("no garbled multi-line writes (interleaving)", len(garbled) == 0)
# All 20 stream lines should have made it to the mock (plus the 10
# interleaved jog commands also went through — filter those out).
non_jog_lines = [ln for ln in stream_lines_only(tracking_mock.completed_lines)
                  if not ln.startswith('$J=')]
jog_lines = [ln for ln in tracking_mock.completed_lines
              if ln.startswith('$J=')]
check(f"all 20 stream lines delivered despite jogs (got "
      f"{len(non_jog_lines)} stream + {len(jog_lines)} jogs)",
      len(non_jog_lines) == 20)
check(f"all 10 jogs delivered (got {len(jog_lines)})",
      len(jog_lines) == 10)


# --- Stall watchdog: lost-ack USB hiccup is detected, not hung forever ---
# Regression for the 79%-and-Idle deadlock that wasted hours of basswood.
# Simulates Grbl silently dropping an `ok` byte on the USB link: the
# streamer thinks lines are still in flight but Grbl says Idle. Without
# the watchdog, this hangs the worker thread forever.

class MockSerialDropAcks(MockSerial):
    """Like MockSerial, but silently drops every ack after the Nth line.

    Mimics a lost-`ok` over USB — the controller still processes the
    G-code (Idle once done), but the host never sees the acks.
    """
    def __init__(self, drop_after):
        super().__init__()
        self.drop_after = drop_after

    def _process_input(self):
        while b'\n' in self.in_buf:
            idx = self.in_buf.index(b'\n')
            line = bytes(self.in_buf[:idx]).decode('ascii', 'replace').strip()
            del self.in_buf[:idx + 1]
            if not line:
                self.out_buf.extend(b'ok\r\n')
                continue
            self.completed_lines.append(line)
            if len(self.completed_lines) <= self.drop_after:
                self.out_buf.extend(b'ok\r\n')
            # else: silently swallow the ack — Grbl finished it but the byte was lost

_orig_timeout = fs.STALL_TIMEOUT_S
_orig_max = fs.MAX_STALL_RECOVERIES
fs.STALL_TIMEOUT_S = 0.5  # speed the watchdog up for the test
try:
    # 1. Single lost-ack burst — watchdog should recover and complete.
    sender = fs.FalconSender(port='MOCK', baud=115200)
    mock = MockSerialDropAcks(drop_after=2)
    sender._serial = mock
    sender.status = {'state': 'Idle'}
    errors = []
    done_reasons = []
    sender.on_error = lambda m: errors.append(m)
    sender.on_done = lambda r: done_reasons.append(r)
    # 20 short lines all fit in Grbl's 128-byte buffer so they all
    # go out before the watchdog fires. The mock drops every ack
    # after the 2nd, but since this is only ONE burst, the recovery
    # path should kick in once and the stream should complete.
    sender.start_stream([f'G0X{i}' for i in range(20)])
    sender._worker.join(timeout=4.0)
    check("recovery joined worker", not sender._worker.is_alive())
    check(f"recovery message surfaced (errors={errors})",
          any('recovered' in e.lower() for e in errors))
    check(f"recovery ends with reason 'complete' (got {done_reasons})",
          done_reasons == ['complete'])

    # 2. USB link genuinely broken — every ack lost forever. After
    # MAX_STALL_RECOVERIES the streamer should give up.
    class MockSerialDropAllAcks(MockSerial):
        def _process_input(self):
            while b'\n' in self.in_buf:
                idx = self.in_buf.index(b'\n')
                line = bytes(self.in_buf[:idx]).decode('ascii', 'replace').strip()
                del self.in_buf[:idx + 1]
                if not line:
                    self.out_buf.extend(b'ok\r\n')
                    continue
                self.completed_lines.append(line)
                # swallow every ack

    fs.MAX_STALL_RECOVERIES = 2  # speed up the "giving up" path
    sender = fs.FalconSender(port='MOCK', baud=115200)
    mock = MockSerialDropAllAcks()
    sender._serial = mock
    sender.status = {'state': 'Idle'}
    errors = []
    done_reasons = []
    sender.on_error = lambda m: errors.append(m)
    sender.on_done = lambda r: done_reasons.append(r)
    sender.start_stream([f'G0X{i}' for i in range(60)])
    sender._worker.join(timeout=10.0)
    check("give-up joined worker", not sender._worker.is_alive())
    check(f"give-up message surfaced (errors={errors})",
          any('giving up' in e.lower() for e in errors))
    check(f"give-up ends with reason 'error' (got {done_reasons})",
          done_reasons == ['error'])
finally:
    fs.STALL_TIMEOUT_S = _orig_timeout
    fs.MAX_STALL_RECOVERIES = _orig_max


# --- Character-counting: many lines stream cleanly without buffer overflow ---
sender, mock = _new_sender_with_mock()
# 30 fairly long lines — char-counting protocol means the streamer
# keeps the Grbl buffer full but never overflows it (Grbl would silently
# drop bytes if it did).
long_lines = ['G1X100.0Y100.0F1000'] * 30
sender.start_stream(long_lines)
worker_ref = sender._worker
worker_ref.join(timeout=5.0)
check("char-counting stream completes", not worker_ref.is_alive())
check(f"all 30 long lines delivered (got {len(stream_lines_only(mock.completed_lines))})",
      len(stream_lines_only(mock.completed_lines)) == 30)


# ============================================================
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
