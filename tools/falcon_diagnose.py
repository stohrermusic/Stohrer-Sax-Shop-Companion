"""Send a handful of diagnostic commands to the Falcon and print the
responses. Run this with the Falcon plugged in via USB and idle (not
mid-job in LightBurn).

    python tools/falcon_diagnose.py

Paste the output back into the conversation.
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import falcon_sender

try:
    import serial
except ImportError:
    print("ERROR: pyserial not installed. pip install pyserial")
    sys.exit(1)


def banner(label):
    print()
    print("=" * 60)
    print(label)
    print("=" * 60)


def send_and_read(ser, cmd, wait_s, label):
    """Send a single command, read whatever comes back for ``wait_s``
    seconds, print it all."""
    banner(f"> {label}  ({cmd!r})")
    try:
        ser.reset_input_buffer()
    except Exception:
        pass
    ser.write(cmd.encode('ascii') + b'\n')
    deadline = time.monotonic() + wait_s
    buf = b''
    while time.monotonic() < deadline:
        try:
            chunk = ser.read(ser.in_waiting or 1)
        except Exception as e:
            print(f"  (read failed: {e})")
            return
        if chunk:
            buf += chunk
        else:
            time.sleep(0.05)
    text = buf.decode('ascii', errors='replace').strip()
    if not text:
        print("  (no response)")
    else:
        for line in text.splitlines():
            print(f"  {line}")


def main():
    # 1. Find a Falcon port
    port = falcon_sender.auto_detect_falcon()
    if not port:
        print("ERROR: no Falcon detected. Is it plugged in via USB?")
        print("Available ports:")
        try:
            from serial.tools import list_ports
            for p in list_ports.comports():
                print(f"  {p.device}  {p.description}")
        except Exception:
            pass
        sys.exit(1)
    print(f"Falcon detected on {port}")

    # 2. Open serial port directly (bypass FalconSender so we can
    #    control timing of each command manually)
    try:
        ser = serial.Serial(port, 115200, timeout=0.05)
    except Exception as e:
        print(f"ERROR: could not open {port}: {e}")
        sys.exit(1)

    # 3. Wake the controller (Falcon doesn't always send a banner on
    #    connect — Ctrl-X soft-reset elicits one)
    banner("Wakeup (Ctrl-X soft-reset)")
    ser.write(b'\x18')
    time.sleep(1.5)
    try:
        wake_buf = ser.read(ser.in_waiting or 256)
    except Exception:
        wake_buf = b''
    for line in wake_buf.decode('ascii', errors='replace').splitlines():
        if line.strip():
            print(f"  {line}")

    # 4. Clear any alarm so subsequent commands aren't blocked
    send_and_read(ser, '$X', 0.5, 'Unlock (clear alarm)')

    # 5. Diagnostics
    send_and_read(ser, '$I', 0.5, 'Build info — Grbl version + build options')
    send_and_read(ser, '?',  0.5, 'Current state + MPos')
    send_and_read(ser, '$$', 1.0,
                   'All settings — LOOK FOR $22 (1=homing enabled)')

    # 6. The big one — try to home. Long timeout because $H only acks
    #    after the homing cycle completes.
    print()
    print("About to send $H — homing cycle. If your Falcon's homing")
    print("is enabled, the head will move toward the home switches at")
    print("full speed. Make sure nothing is in its path. If it isn't")
    print("enabled, you'll just see an error.")
    print()
    input("Press Enter to send $H (Ctrl-C to skip)...")
    send_and_read(ser, '$H', 60.0, 'Home — long timeout (60s)')

    # 7. Where did we end up?
    send_and_read(ser, '?', 0.5, 'State + MPos after homing')

    ser.close()
    print()
    print("Done. Paste everything above back to the chat.")


if __name__ == '__main__':
    main()
