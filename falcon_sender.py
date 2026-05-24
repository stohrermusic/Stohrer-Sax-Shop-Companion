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

# Grbl real-time command bytes (single-byte, immediate, don't queue)
RT_STATUS = b'?'              # request status report
RT_FEED_HOLD = b'!'           # pause feed (laser keeps power; M9 to also kill it)
RT_CYCLE_START = b'~'         # resume from feed hold
RT_SOFT_RESET = b'\x18'       # Ctrl-X, clears planner + serial buffer
RT_SAFETY_DOOR = b'\x84'
RT_JOG_CANCEL = b'\x85'

# Falcon2 Pro identifies as a CH340 USB-serial adapter on most installs.
# Match by VID:PID when possible; fall back to handshake.
FALCON_VID_PID_HINTS = [
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
    """Open a port, wait for Grbl banner, return True if seen."""
    if not HAS_PYSERIAL:
        return False
    try:
        # Open with DTR low to avoid resetting some boards on connect;
        # for Grbl we DO want a reset to get the banner, so default DTR is fine.
        s = serial.Serial(port, baud, timeout=0.2)
    except (serial.SerialException, OSError):
        return False
    try:
        # Some boards take a moment to enumerate after open. Read for the
        # full timeout window and look for a Grbl banner.
        deadline = time.monotonic() + timeout_s
        buf = b''
        while time.monotonic() < deadline:
            chunk = s.read(256)
            if chunk:
                buf += chunk
                # The banner looks like "Grbl 1.1f ['$' for help]"
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

        self.status = {'state': 'Disconnected'}
        self.latest_error = None
        self.latest_alarm = None
        self._total_lines = 0
        self._sent_lines = 0

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------

    def connect(self):
        """Open the serial port and wait for the Grbl banner."""
        with self._lock:
            if self._serial is not None and self._serial.is_open:
                return
            self._serial = serial.Serial(self.port, self.baud, timeout=0.05)
            # Wait for the boot banner. If we don't see it within the window,
            # try a soft-reset; that always elicits the banner.
            deadline = time.monotonic() + CONNECT_WAKEUP_DELAY_S
            buf = b''
            while time.monotonic() < deadline:
                chunk = self._serial.read(256)
                if chunk:
                    buf += chunk
                    if b'Grbl' in buf:
                        break
            if b'Grbl' not in buf:
                self._serial.write(RT_SOFT_RESET)
                time.sleep(CONNECT_WAKEUP_DELAY_S)
                buf += self._serial.read(self._serial.in_waiting or 256)
            self._serial.reset_input_buffer()
            self.status = {'state': 'Idle', 'raw': buf.decode('ascii', 'replace')}

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
        """
        self._stop_flag.set()
        self._pause_flag.clear()
        self.send_realtime(RT_SOFT_RESET)
        if self._worker is not None and self._worker.is_alive():
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

    # ------------------------------------------------------------------
    # Jog
    # ------------------------------------------------------------------

    def jog(self, x=None, y=None, z=None, feed=2000, relative=False):
        """Send a jog command. Uses Grbl 1.1 `$J=` syntax (cancellable)."""
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
            except (serial.SerialException, OSError):
                pass

    # ------------------------------------------------------------------
    # Streaming
    # ------------------------------------------------------------------

    def start_stream(self, gcode_lines):
        """Start streaming a list of G-code lines on a worker thread.

        Lines should be strings WITHOUT trailing newlines (the streamer
        appends `\\n`). Comment-only / blank lines are kept (we send
        them too, so progress numbers match user-visible line counts).
        """
        if self._worker is not None and self._worker.is_alive():
            raise RuntimeError("A stream is already in progress")
        self._stop_flag.clear()
        self._pause_flag.clear()
        self._total_lines = len(gcode_lines)
        self._sent_lines = 0
        self.latest_error = None
        self.latest_alarm = None
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

        sent_lengths = deque()  # lengths of in-flight lines waiting on "ok"
        ack_count = 0           # how many lines have been acknowledged
        line_idx = 0
        last_poll = 0.0
        rx_buf = b''
        poll_interval = 1.0 / STATUS_POLL_HZ
        finished_sending = False
        reason = "complete"

        while True:
            if self._stop_flag.is_set():
                reason = "stopped"
                break

            # 1) Send next line if there's room in Grbl's buffer
            if (not finished_sending
                and not self._pause_flag.is_set()
                and line_idx < len(lines)):
                next_line = lines[line_idx].strip() + '\n'
                next_bytes = next_line.encode('ascii', errors='replace')
                if sum(sent_lengths) + len(next_bytes) < GRBL_RX_BUFFER:
                    try:
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
                        if sent_lengths:
                            sent_lengths.popleft()
                        ack_count += 1
                    elif text.startswith('error:'):
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
