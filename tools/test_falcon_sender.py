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
      mock.completed_lines == lines)
check("progress callbacks fired for each line",
      len(progress_calls) == len(lines)
      and progress_calls[-1] == (len(lines), len(lines)))


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
check(f"all 30 long lines delivered (got {len(mock.completed_lines)})",
      len(mock.completed_lines) == 30)


# ============================================================
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
