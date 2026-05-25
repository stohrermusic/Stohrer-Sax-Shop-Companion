"""Grbl/Falcon G-code streamer over USB serial.

Built for the Creality Falcon2 Pro 40W (Grbl v1.1f firmware, USB-C @
115200 baud) and any other stock Grbl v1.1+ controller. Implements the
character-counting streaming protocol from the official Grbl docs so the
controller's 128-byte serial buffer stays full and the lookahead planner
doesn't stutter mid-raster.

pyserial is an optional dependency. The rest of the app degrades
gracefully when it's not installed.

Public API:
    auto_detect_falcon()       -> port str or None
    list_serial_ports()        -> list of port info dicts
    FalconSender(port, baud)
        connect() / disconnect()
        start_stream(lines)    -> begins streaming on a worker thread
        pause() / resume() / stop()
        jog(x, y, feed)
        send_realtime(byte)
        status                 -> dict, updated by status polls
        latest_error           -> last "error:N" line, if any
        latest_alarm           -> last "ALARM:N" line, if any

Callbacks (provided via constructor, invoked from the worker thread):
    on_status(status_dict)
    on_progress(line_index, total_lines)
    on_error(err_str)
    on_done(reason)            reason: "complete" | "stopped" | "error"

Threading model:
    All serial I/O is on a single worker thread. UI callers must marshal
    callback values back to the Tk main thread via root.after_idle() or
    similar — this module never touches Tk.
"""

from __future__ import annotations

import re
import threading
import time
from collections import deque

try:
    import serial
    import serial.tools.list_ports
    HAS_PYSERIAL = True
except ImportError:  # pragma: no cover
    serial = None
    HAS_PYSERIAL = False


# =============================================================================
# Constants
# =============================================================================

DEFAULT_BAUD = 115200
GRBL_RX_BUFFER = 128          # bytes, fixed in Grbl firmware
STATUS_POLL_HZ = 4            # `?` polls per second during streaming
CONNECT_WAKEUP_DELAY_S = 1.5  # time to wait for Grbl banner after open
# Stall watchdog: when Grbl reports Idle continuously for this many
# seconds while we still have lines waiting on acks, the
# character-counting state has desynced — an `ok` was lost over USB.
# Grbl being Idle means its planner is empty AND its serial RX has
# been processed, so every line we sent has definitely executed. We
# can safely clear the in-flight queue and keep streaming.
STALL_TIMEOUT_S = 15.0
# After this many recoveries on a single stream, give up — the USB
# link is too unreliable to trust (probably a bad cable or hub).
MAX_STALL_RECOVERIES = 5

# Grbl real-time command bytes (single-byte, immediate, don't queue)
RT_STATUS = b'?'              # request status report
RT_FEED_HOLD = b'!'           # pause feed (laser keeps power; M9 to also kill it)
RT_CYCLE_START = b'~'         # resume from feed hold
RT_SOFT_RESET = b'\x18'       # Ctrl-X, clears planner + serial buffer
RT_SAFETY_DOOR = b'\x84'
RT_JOG_CANCEL = b'\x85'

# Falcon2 Pro 40W identifies as Espressif (ESP32-S3 onboard); older
# Falcons use CH340. Match by VID when possible to skip non-laser
# devices first; fall back to handshake on every port either way.
FALCON_VID_PID_HINTS = [
    (0x303A, 0x4001),  # Espressif ESP32-S3 (Falcon2 Pro 40W)
    (0x1A86, 0x7523),  # CH340
    (0x1A86, 0x55D4),  # CH9102
    (0x1A86, 0x7522),  # CH340N
]


# =============================================================================
# Helpers — port enumeration & auto-detect
# =============================================================================

def list_serial_ports():
    """Return a list of {device, description, hwid, vid, pid} dicts."""
    if not HAS_PYSERIAL:
        return []
    out = []
    for p in serial.tools.list_ports.comports():
        out.append({
            'device': p.device,
            'description': p.description or '',
            'hwid': p.hwid or '',
            'vid': p.vid,
            'pid': p.pid,
            'manufacturer': p.manufacturer or '',
        })
    return out


def _probe_for_grbl(port, baud=DEFAULT_BAUD, timeout_s=CONNECT_WAKEUP_DELAY_S):
    """Open a port, send a soft-reset, wait for the Grbl banner.

    A passive listen doesn't work for the Falcon2 Pro 40W's ESP32-S3
    controller — it boots silently and only emits the banner in response
    to ``\\x18`` (Ctrl-X soft-reset). The reset is harmless to send to a
    non-Grbl device; pyserial will just close the port and we'll move on.
    """
    if not HAS_PYSERIAL:
        return False
    try:
        s = serial.Serial(port, baud, timeout=0.2)
    except (serial.SerialException, OSError):
        return False
    try:
        # Give the port a moment to settle after enumeration, drain
        # whatever the OS buffered, then kick the controller and listen
        # for its boot banner.
        time.sleep(0.2)
        try:
            s.reset_input_buffer()
            s.write(RT_SOFT_RESET)
        except (serial.SerialException, OSError):
            return False
        deadline = time.monotonic() + timeout_s
        buf = b''
        while time.monotonic() < deadline:
            chunk = s.read(256)
            if chunk:
                buf += chunk
                if b'Grbl' in buf:
                    return True
        return b'Grbl' in buf
    finally:
        try:
            s.close()
        except Exception:
            pass


def auto_detect_falcon():
    """Try to find a Grbl-speaking serial port. Returns the port device or None.

    Order:
      1. Any port matching a known Falcon VID:PID — probe that first.
      2. Any port whose description mentions Falcon / CH340 — probe next.
      3. Any remaining serial port — handshake-probe in order.
    """
    if not HAS_PYSERIAL:
        return None
    ports = list_serial_ports()
    if not ports:
        return None

    # Build a priority-ordered list of candidates without duplicates
    priority = []
    seen = set()

    def _add(p):
        if p['device'] not in seen:
            priority.append(p)
            seen.add(p['device'])

    for p in ports:
        if p['vid'] and p['pid'] and (p['vid'], p['pid']) in FALCON_VID_PID_HINTS:
            _add(p)
    for p in ports:
        desc = (p['description'] + ' ' + p['hwid']).lower()
        if 'falcon' in desc or 'ch340' in desc or 'creality' in desc:
            _add(p)
    for p in ports:
        _add(p)

    for p in priority:
        if _probe_for_grbl(p['device']):
            return p['device']
    return None


# =============================================================================
# Status report parsing
# =============================================================================

# Grbl 1.1 status reports look like:
#   <Idle|MPos:0.000,0.000,0.000|FS:0,0>
#   <Run|MPos:5.529,0.560,0.000|FS:200,0|Pn:S|Ov:100,100,100>
#   <Alarm|MPos:0.000,0.000,0.000|WCO:0.000,0.000,0.000>
_STATUS_RE = re.compile(r'<([^>]+)>')


def parse_status(line):
    """Parse a Grbl 1.1 status report string into a dict.

    Returns None if `line` is not a status report.
    """
    m = _STATUS_RE.search(line)
    if not m:
        return None
    fields = m.group(1).split('|')
    state = fields[0]
    out = {'state': state, 'raw': line.strip()}
    for f in fields[1:]:
        if ':' not in f:
            continue
        key, val = f.split(':', 1)
        if key in ('MPos', 'WPos'):
            try:
                xyz = [float(v) for v in val.split(',')]
                out[key.lower()] = xyz
            except ValueError:
                pass
        elif key == 'FS':
            try:
                fs = [float(v) for v in val.split(',')]
                out['feed'] = fs[0] if len(fs) > 0 else 0.0
                out['spindle'] = fs[1] if len(fs) > 1 else 0.0
            except ValueError:
                pass
        elif key == 'WCO':
            try:
                out['wco'] = [float(v) for v in val.split(',')]
            except ValueError:
                pass
        elif key == 'Pn':
            out['pins'] = val
        elif key == 'Ov':
            try:
                out['overrides'] = [int(v) for v in val.split(',')]
            except ValueError:
                pass
        else:
            out[key.lower()] = val
    return out


# =============================================================================
# Sender
# =============================================================================

class FalconSender:
    """Stream G-code to a Grbl controller over USB serial.

    Construct, connect, start_stream() with a list of G-code lines (each
    line without trailing newline). Callbacks are invoked from the worker
    thread — UI consumers should marshal to the Tk main thread.
    """

    def __init__(self, port=None, baud=DEFAULT_BAUD,
                  on_status=None, on_progress=None,
                  on_error=None, on_done=None,
                  on_alarm=None):
        if not HAS_PYSERIAL:
            raise RuntimeError("pyserial is not installed")
        self.port = port
        self.baud = baud
        self.on_status = on_status
        self.on_progress = on_progress
        self.on_error = on_error
        self.on_done = on_done
        self.on_alarm = on_alarm

        self._serial = None
        self._worker = None
        self._stop_flag = threading.Event()
        self._pause_flag = threading.Event()
        self._lock = threading.RLock()
        # Outstanding jogs whose ``ok`` Grbl hasn't yet returned. Each
        # ``$J=`` ack must be consumed BEFORE we treat any ``ok`` as a
        # stream-line ack — otherwise the character-counting protocol
        # over-counts and either terminates the stream early or leaves
        # the planner over-loaded. Read/written under self._lock.
        self._jogs_in_flight = 0

        self.status = {'state': 'Disconnected'}
        self.latest_error = None
        self.latest_alarm = None
        self._total_lines = 0
        self._sent_lines = 0

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self):
        """Open the serial port and confirm Grbl is responsive.

        Strategy: send ``?`` (real-time status query) and check for a
        ``<...>`` reply. Grbl responds to ``?`` in every state
        including Alarm, so this is a non-destructive liveness probe.
        Only if the probe gets nothing do we fall back to the heavier
        soft-reset (which DOES elicit the boot banner — but also
        puts the controller into Alarm state if homing is required,
        causing a continuous beep).
        """
        with self._lock:
            if self._serial is not None and self._serial.is_open:
                return
            self._serial = serial.Serial(self.port, self.baud, timeout=0.05)
            time.sleep(0.1)  # let port settle

            # Probe with ? — non-destructive
            try:
                self._serial.reset_input_buffer()
                self._serial.write(RT_STATUS)
            except (serial.SerialException, OSError):
                pass
            time.sleep(0.2)
            try:
                buf = self._serial.read(self._serial.in_waiting or 256)
            except (serial.SerialException, OSError):
                buf = b''
            if b'<' in buf and b'>' in buf:
                # Grbl alive — done, no soft-reset needed
                self._serial.reset_input_buffer()
                self.status = {'state': 'Idle',
                                'raw': buf.decode('ascii', 'replace')}
                return

            # Probe failed — wait for the boot banner (Falcon sometimes
            # needs a beat after USB enumeration before it'll talk).
            deadline = time.monotonic() + CONNECT_WAKEUP_DELAY_S
            while time.monotonic() < deadline:
                chunk = self._serial.read(256)
                if chunk:
                    buf += chunk
                    if b'Grbl' in buf:
                        break

            # Last resort: soft-reset. This DOES wake the controller but
            # puts it back into Alarm if homing is required ($22=1).
            # Only get here on a really stuck connection.
            if b'Grbl' not in buf and (b'<' not in buf or b'>' not in buf):
                self._serial.write(RT_SOFT_RESET)
                time.sleep(CONNECT_WAKEUP_DELAY_S)
                buf += self._serial.read(self._serial.in_waiting or 256)
            self._serial.reset_input_buffer()
            self.status = {'state': 'Idle',
                            'raw': buf.decode('ascii', 'replace')}

    def disconnect(self):
        """Close the serial port. Stops any in-flight stream first."""
        self.stop()
        with self._lock:
            if self._serial is not None:
                try:
                    self._serial.close()
                except Exception:
                    pass
                self._serial = None
            self.status = {'state': 'Disconnected'}

    def is_connected(self):
        with self._lock:
            return self._serial is not None and self._serial.is_open

    # ------------------------------------------------------------------
    # Real-time commands
    # ------------------------------------------------------------------

    def send_realtime(self, byte):
        """Send a single real-time command byte. Doesn't queue."""
        with self._lock:
            if self._serial is None or not self._serial.is_open:
                return
            try:
                self._serial.write(byte)
            except (serial.SerialException, OSError):
                pass

    def pause(self):
        """Feed hold — pauses motion; laser stays in current state."""
        self._pause_flag.set()
        self.send_realtime(RT_FEED_HOLD)

    def resume(self):
        """Resume from feed hold."""
        self._pause_flag.clear()
        self.send_realtime(RT_CYCLE_START)

    def stop(self):
        """Emergency stop: soft-reset clears planner buffer and stops motion.

        The worker thread is signalled to exit. Safe to call from any thread.

        Only sends RT_SOFT_RESET when there's an active stream to stop.
        Without the guard, the soft-reset puts the Falcon into Alarm
        state every time disconnect() is called — including the
        connect-poll-disconnect cycle used by status polling — which
        produces a continuous beep.
        """
        self._stop_flag.set()
        self._pause_flag.clear()
        if self._worker is not None and self._worker.is_alive():
            self.send_realtime(RT_SOFT_RESET)
            self._worker.join(timeout=2.0)
        self._worker = None

    def unlock(self):
        """Send `$X` to clear an alarm state."""
        with self._lock:
            if self._serial is None:
                return
            try:
                self._serial.write(b'$X\n')
            except (serial.SerialException, OSError):
                pass

    def home(self, timeout_s=60.0):
        """Send ``$H`` and block until Grbl returns ``ok`` (homing done)
        or an ``error:N`` / ``ALARM:N`` arrives.

        Returns a ``(success, message)`` tuple. ``message`` is the raw
        Grbl response (e.g. ``"ok"`` or ``"error:5"``). Caller must
        ensure no stream is active — this races for reads with the
        worker thread.

        ``timeout_s`` covers the whole homing cycle. On a Falcon2 Pro
        40W with $130=400/$131=415, homing typically completes in 15-25
        seconds; 60s is comfortable headroom.
        """
        with self._lock:
            if self._serial is None or not self._serial.is_open:
                return (False, "not connected")
            try:
                self._serial.reset_input_buffer()
                self._serial.write(b'$H\n')
            except (serial.SerialException, OSError) as e:
                return (False, f"write failed: {e}")
            deadline = time.monotonic() + timeout_s
            buf = b''
            while time.monotonic() < deadline:
                try:
                    chunk = self._serial.read(self._serial.in_waiting or 1)
                except (serial.SerialException, OSError) as e:
                    return (False, f"read failed: {e}")
                if chunk:
                    buf += chunk
                    while b'\n' in buf:
                        line, buf = buf.split(b'\n', 1)
                        text = line.decode('ascii', 'replace').strip()
                        if not text:
                            continue
                        if text == 'ok':
                            return (True, 'ok')
                        if text.startswith('error:'):
                            self.latest_error = text
                            return (False, text)
                        if text.startswith('ALARM:'):
                            self.latest_alarm = text
                            return (False, text)
                else:
                    time.sleep(0.05)
            return (False, f"timeout after {timeout_s:.0f}s")

    def get_status(self, timeout=0.3):
        """Synchronous ``?`` query. Returns a parsed status dict or None.

        Caller must not invoke this during an active stream — the worker
        thread reads the port and will race for the response. Intended
        for between-job UIs (origin calibration, jog dialogs) where the
        sender is connected but idle.
        """
        with self._lock:
            if self._serial is None or not self._serial.is_open:
                return None
            try:
                self._serial.reset_input_buffer()
                self._serial.write(RT_STATUS)
            except (serial.SerialException, OSError):
                return None
            deadline = time.monotonic() + timeout
            buf = b''
            while time.monotonic() < deadline:
                try:
                    chunk = self._serial.read(256)
                except (serial.SerialException, OSError):
                    return None
                if chunk:
                    buf += chunk
                    if b'>' in buf and b'<' in buf:
                        break
            for line in buf.decode('ascii', 'replace').splitlines():
                line = line.strip()
                if line.startswith('<') and line.endswith('>'):
                    parsed = parse_status(line)
                    if parsed:
                        self.status = parsed
                        return parsed
            return None

    # ------------------------------------------------------------------
    # Jog
    # ------------------------------------------------------------------

    def jog(self, x=None, y=None, z=None, feed=2000, relative=False):
        """Send a jog command. Uses Grbl 1.1 `$J=` syntax (cancellable).

        Safe to call during streaming: the write is serialized with
        the streamer's writes via self._lock, and the jog's eventual
        ``ok`` from Grbl is consumed by the stream reader without
        being counted as a stream-line ack (see ``_jogs_in_flight``).
        Without that bookkeeping, a mid-stream jog corrupts the
        character-counting protocol and Grbl can return error:8.
        """
        parts = []
        if relative:
            parts.append('G91')
        else:
            parts.append('G90')
        if x is not None:
            parts.append(f'X{x}')
        if y is not None:
            parts.append(f'Y{y}')
        if z is not None:
            parts.append(f'Z{z}')
        parts.append(f'F{feed}')
        cmd = '$J=' + ''.join(parts) + '\n'
        with self._lock:
            if self._serial is None:
                return
            try:
                self._serial.write(cmd.encode('ascii'))
                self._jogs_in_flight += 1
            except (serial.SerialException, OSError):
                pass  # do NOT count an unack-able jog

    # ------------------------------------------------------------------
    # Streaming
    # ------------------------------------------------------------------

    def start_stream(self, gcode_lines):
        """Start streaming a list of G-code lines on a worker thread.

        Lines should be strings WITHOUT trailing newlines (the streamer
        appends `\\n`). Comment-only / blank lines are skipped in
        the streamer loop (Grbl doesn't ack them).
        """
        # Old worker may still be exiting after _notify_done — give it
        # a brief join window before refusing. Without this, fast
        # restart-stream loops (calibration framing) hit a race where
        # is_alive() is still True for milliseconds after the worker
        # has logically completed.
        if self._worker is not None and self._worker.is_alive():
            self._worker.join(timeout=0.5)
            if self._worker.is_alive():
                raise RuntimeError("A stream is already in progress")
        self._stop_flag.clear()
        self._pause_flag.clear()
        self._total_lines = len(gcode_lines)
        self._sent_lines = 0
        self.latest_error = None
        self.latest_alarm = None
        # Reset jog tracking so any leftover (an ack arrived after the
        # last stream ended but before start_stream() was called)
        # doesn't bleed into this stream's character-counting state.
        with self._lock:
            self._jogs_in_flight = 0
        self._worker = threading.Thread(
            target=self._stream_worker, args=(list(gcode_lines),),
            name='FalconSender-stream', daemon=True,
        )
        self._worker.start()

    def is_streaming(self):
        return self._worker is not None and self._worker.is_alive()

    def _stream_worker(self, lines):
        """Run on the worker thread. Implements character-counting protocol.

        Sends as many lines as fit in Grbl's 128-byte RX buffer at once.
        Each "ok" pops the oldest line length from the queue and frees
        that much room. Status polls are interleaved at STATUS_POLL_HZ.
        """
        try:
            self._stream_loop(lines)
        except Exception as e:
            self._notify_error(f"streamer crashed: {e}")
            self._notify_done("error")

    def _stream_loop(self, lines):
        ser = self._serial
        if ser is None:
            self._notify_error("no serial connection")
            self._notify_done("error")
            return

        # Wake-ping. Empirically the Falcon2 Pro 40W can sit unresponsive
        # at the start of a stream if there's been an extended idle gap
        # since the last command (e.g. between the auto-frame seed move
        # closing and the user clicking "Start Frame →"). Sending `?`
        # nudges the controller without affecting state — a no-op when
        # it's already awake. The status response is harmlessly read
        # by the loop below.
        self.send_realtime(RT_STATUS)
        time.sleep(0.05)

        sent_lengths = deque()  # lengths of in-flight lines waiting on "ok"
        ack_count = 0           # how many lines have been acknowledged
        line_idx = 0
        last_poll = 0.0
        rx_buf = b''
        poll_interval = 1.0 / STATUS_POLL_HZ
        finished_sending = False
        reason = "complete"
        last_ack_time = time.monotonic()      # for stall watchdog
        idle_since = None                      # when state first became Idle
        stall_recoveries = 0                   # count per stream

        while True:
            if self._stop_flag.is_set():
                reason = "stopped"
                break

            # 1) Send next line if there's room in Grbl's buffer
            if (not finished_sending
                and not self._pause_flag.is_set()
                and line_idx < len(lines)):
                raw = lines[line_idx].strip()
                # Grbl 1.1 does NOT send `ok` for comment-only or
                # blank lines. Sending them anyway would leave the
                # character-counting streamer waiting forever for an
                # ack that never arrives. Skip them client-side and
                # advance ack_count alongside line_idx so the
                # done-check (ack_count >= line_idx) still fires.
                if not raw or raw.startswith(';'):
                    line_idx += 1
                    ack_count += 1
                    self._sent_lines = line_idx
                    self._notify_progress(line_idx, self._total_lines)
                    continue
                next_line = raw + '\n'
                next_bytes = next_line.encode('ascii', errors='replace')
                if sum(sent_lengths) + len(next_bytes) < GRBL_RX_BUFFER:
                    # Serialize with jog() / send_realtime() / etc. so
                    # the bytes can't interleave with another writer.
                    # Concurrent writes garble Grbl's line parser and
                    # produce things like error:8.
                    try:
                        with self._lock:
                            ser.write(next_bytes)
                    except (serial.SerialException, OSError) as e:
                        self._notify_error(f"serial write failed: {e}")
                        reason = "error"
                        break
                    sent_lengths.append(len(next_bytes))
                    line_idx += 1
                    self._sent_lines = line_idx
                    self._notify_progress(line_idx, self._total_lines)
                else:
                    pass  # buffer is full, wait for an ack

            if line_idx >= len(lines):
                finished_sending = True

            # 2) Read everything available; split on newlines
            try:
                chunk = ser.read(ser.in_waiting or 1)
            except (serial.SerialException, OSError) as e:
                self._notify_error(f"serial read failed: {e}")
                reason = "error"
                break
            if chunk:
                rx_buf += chunk
                while b'\n' in rx_buf:
                    line, rx_buf = rx_buf.split(b'\n', 1)
                    text = line.decode('ascii', errors='replace').strip()
                    if not text:
                        continue
                    if text == 'ok':
                        # Each jog command Grbl receives also produces an
                        # `ok`. Consume those FIRST so we don't mistake a
                        # jog ack for a stream-line ack (which would drain
                        # sent_lengths early and let the streamer over-
                        # fill Grbl's RX buffer).
                        with self._lock:
                            if self._jogs_in_flight > 0:
                                self._jogs_in_flight -= 1
                                continue
                        if sent_lengths:
                            sent_lengths.popleft()
                        ack_count += 1
                        last_ack_time = time.monotonic()
                    elif text.startswith('error:'):
                        # If jogs are in flight, this error is more
                        # likely from the jog than from the stream
                        # (Grbl returns errors in command-receive
                        # order). Consume it against the jog queue
                        # and keep streaming — surfacing the error
                        # so the UI can show it, but not killing the
                        # stream over a rejected jog.
                        with self._lock:
                            if self._jogs_in_flight > 0:
                                self._jogs_in_flight -= 1
                                self.latest_error = text
                                self._notify_error(text)
                                continue
                        self.latest_error = text
                        self._notify_error(text)
                        # Per Grbl docs, an error means the controller
                        # entered an alarm state. Force-complete the
                        # streamer so the done-condition fires on the
                        # next loop iteration.
                        reason = "error"
                        finished_sending = True
                        ack_count = line_idx
                        sent_lengths.clear()
                    elif text.startswith('ALARM:'):
                        self.latest_alarm = text
                        if self.on_alarm:
                            try:
                                self.on_alarm(text)
                            except Exception:
                                pass
                        reason = "error"
                        finished_sending = True
                        ack_count = line_idx
                        sent_lengths.clear()
                    elif text.startswith('<'):
                        parsed = parse_status(text)
                        if parsed:
                            self.status = parsed
                            # Track Idle continuity for the stall watchdog
                            if parsed.get('state') == 'Idle':
                                if idle_since is None:
                                    idle_since = time.monotonic()
                            else:
                                idle_since = None
                            if self.on_status:
                                try:
                                    self.on_status(parsed)
                                except Exception:
                                    pass

            # 3) Poll status periodically
            now = time.monotonic()
            if now - last_poll >= poll_interval:
                self.send_realtime(RT_STATUS)
                last_poll = now

            # 3b) Stall watchdog with auto-recovery. If Grbl has been
            # Idle for STALL_TIMEOUT_S while we still have lines
            # marked in-flight (sent_lengths non-empty) and no ack
            # has arrived in that window, an `ok` byte was lost over
            # USB. Because Idle means Grbl's planner is empty AND
            # its serial RX has been processed, every line we sent
            # is definitively done — we just lost the ack(s). Safely
            # resync the character-count and continue. Cap recoveries
            # per stream so a truly broken USB link still fails out
            # instead of looping forever.
            if (sent_lengths
                    and idle_since is not None
                    and (now - idle_since) > STALL_TIMEOUT_S
                    and (now - last_ack_time) > STALL_TIMEOUT_S):
                pct = (100 * line_idx // self._total_lines
                       if self._total_lines else 0)
                recovered = len(sent_lengths)
                stall_recoveries += 1
                if stall_recoveries > MAX_STALL_RECOVERIES:
                    self._notify_error(
                        f"streamer giving up at line {line_idx} / "
                        f"{self._total_lines} ({pct}%) — "
                        f"{MAX_STALL_RECOVERIES} lost-ack recoveries "
                        f"in one job. Check USB cable / hub."
                    )
                    reason = "error"
                    break
                # Resync. Grbl is done with every sent line.
                sent_lengths.clear()
                ack_count = line_idx
                last_ack_time = now
                idle_since = None  # require fresh Idle observation
                self._notify_error(
                    f"recovered lost ack(s) at line {line_idx} / "
                    f"{self._total_lines} ({pct}%) — {recovered} "
                    f"line(s) confirmed done via Idle state. "
                    f"Continuing."
                )

            # 4) Are we done? All sent AND all acknowledged.
            if finished_sending and ack_count >= line_idx:
                # Wait for the controller to actually finish motion
                self._wait_for_idle()
                break

            if not chunk:
                # Brief sleep when nothing came back; prevents 100% CPU spin
                time.sleep(0.005)

        # If stop() was called after streaming completed but before the
        # idle wait finished, report "stopped" rather than "complete" so
        # the UI reflects the user's intent.
        if self._stop_flag.is_set() and reason == "complete":
            reason = "stopped"
        self._notify_done(reason)

    def _wait_for_idle(self, timeout_s=120.0):
        """After the last `ok`, the planner still has motion queued. Poll
        until state is Idle (or alarm) or timeout."""
        deadline = time.monotonic() + timeout_s
        while time.monotonic() < deadline:
            if self._stop_flag.is_set():
                return
            self.send_realtime(RT_STATUS)
            time.sleep(0.1)
            try:
                chunk = self._serial.read(self._serial.in_waiting or 1)
            except (serial.SerialException, OSError):
                return
            if chunk:
                text = chunk.decode('ascii', errors='replace')
                for line in text.splitlines():
                    if line.startswith('<'):
                        parsed = parse_status(line)
                        if parsed:
                            self.status = parsed
                            if self.on_status:
                                try:
                                    self.on_status(parsed)
                                except Exception:
                                    pass
                            if parsed.get('state') in ('Idle', 'Alarm'):
                                return

    # ------------------------------------------------------------------
    # Callback shims
    # ------------------------------------------------------------------

    def _notify_progress(self, sent, total):
        if self.on_progress:
            try:
                self.on_progress(sent, total)
            except Exception:
                pass

    def _notify_error(self, msg):
        if self.on_error:
            try:
                self.on_error(msg)
            except Exception:
                pass

    def _notify_done(self, reason):
        if self.on_done:
            try:
                self.on_done(reason)
            except Exception:
                pass
