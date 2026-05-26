"""Cross-platform sleep / display-lock prevention for long machine
operations (calibration card engrave, large pad-sheet cuts, etc.).

Without this, Windows can put the system to sleep mid-engrave when
the user steps away — killing the USB serial connection and aborting
the job. Same problem on macOS (system sleep) and Linux (display
manager sleep, which doesn't kill the connection but might lock the
screen and confuse a returning user).

Usage:

    from sleep_lock import prevent_sleep, allow_sleep
    prevent_sleep()
    try:
        run_long_job()
    finally:
        allow_sleep()

The functions are idempotent — pairing is "best effort" rather than
strict refcounting. Multiple prevent_sleep() calls before a single
allow_sleep() will still release the lock; that's intentional for
simplicity. Long jobs should bracket themselves with try/finally.

Platforms:
    - Windows: SetThreadExecutionState with ES_CONTINUOUS |
      ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED.
    - macOS: spawns `caffeinate -dis` and tracks the PID; the
      subprocess prevents sleep + display sleep + idle.
    - Linux: no universal API (depends on the desktop environment),
      so this is a no-op. Users on Linux should disable sleep in
      their DE settings during long jobs.
"""

import platform
import subprocess

_state = {'pid': None}  # macOS caffeinate PID

# Windows execution-state constants (from winbase.h)
_ES_CONTINUOUS = 0x80000000
_ES_SYSTEM_REQUIRED = 0x00000001
_ES_DISPLAY_REQUIRED = 0x00000002


def prevent_sleep():
    """Block system + display sleep until allow_sleep() is called.
    Safe to call multiple times in a row (effectively idempotent)."""
    system = platform.system()
    if system == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(
                _ES_CONTINUOUS | _ES_SYSTEM_REQUIRED
                | _ES_DISPLAY_REQUIRED)
        except Exception:
            pass
    elif system == "Darwin":
        if _state['pid'] is None:
            try:
                proc = subprocess.Popen(
                    ['caffeinate', '-dis'],
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL)
                _state['pid'] = proc
            except Exception:
                pass
    # Linux: no-op; user must disable sleep in their DE.


def allow_sleep():
    """Release the sleep block set by prevent_sleep()."""
    system = platform.system()
    if system == "Windows":
        try:
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(_ES_CONTINUOUS)
        except Exception:
            pass
    elif system == "Darwin":
        proc = _state['pid']
        if proc is not None:
            try:
                proc.terminate()
                proc.wait(timeout=2)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
            _state['pid'] = None
