"""Frame pacing for the tuner's and the Analyzer's animation loops.

No tkinter here: the loops ask FramePacer how long to wait and hand that to
Tk's after(). Moved out of tuner_tab.py on 2026-11-25 when the Analyzer's
loop (toner_tab._toner_animate), which had the same fixed after(interval)
and ran 21 fps on its 30 setting, started using it too.
"""
import logging
import sys

_log = logging.getLogger(__name__)


class FramePacer:
    """Paces an `after()` loop against a clock, so "60" is sixty frames a second.

    The loop used to ask Tk for a fixed 16 ms after each frame's work. That
    adds the work to every period, and on Windows the wait itself is rounded
    up to the system timer tick (15.6 ms unless raised, see fine_timer), so
    4 ms of work + after(16) was 31 ms: 32 fps on the 60 setting, 63 on the
    120 setting (measured in JustATuner and here, 2026-11-25). At 32 fps the
    fine outer rings step more than half a stripe per frame from about 8
    cents off and appear to run backwards.

    Call next_delay_ms once the frame's work AND its redraws are done (the
    tab calls it from an idle callback): it returns how long to wait for the
    next frame's start on a fixed timeline, one period after the last start.
    A frame that overran its period is followed at once, and the timeline
    restarts from there: no debt is carried, so there is never a burst of
    catch-up frames, and a loop that cannot keep up simply runs back to
    back. The wait is never under 1 ms, so Tk reaches its idle work (the
    redraws) before the next frame is due.
    """

    def __init__(self):
        self._deadline = None

    def reset(self):
        self._deadline = None

    def next_delay_ms(self, period_s, now=None):
        if now is None:
            import time
            now = time.perf_counter()
        if self._deadline is None:
            self._deadline = now
        self._deadline = max(self._deadline + period_s, now)
        return max(1, int(round((self._deadline - now) * 1000.0)))


_fine_timer_holds = 0


def fine_timer(on):
    """Windows: hold (True) or release (False) a 1 ms system timer tick for
    this process. The default tick is 15.6 ms and every Tk `after()` is
    rounded up to it, so no frame rate between 32 and 64 fps is reachable
    without this. Held only while the tuner runs (a fine tick costs a little
    power); calls nest. Elsewhere timers are already fine: a no-op.
    Returns whether a hold is in effect."""
    global _fine_timer_holds
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        winmm = ctypes.windll.winmm
        if on:
            if winmm.timeBeginPeriod(1) == 0:      # TIMERR_NOERROR
                _fine_timer_holds += 1
        elif _fine_timer_holds > 0:
            winmm.timeEndPeriod(1)
            _fine_timer_holds -= 1
    except (OSError, AttributeError) as e:
        _log.warning("Could not set the timer resolution: %s", e)
    return _fine_timer_holds > 0
