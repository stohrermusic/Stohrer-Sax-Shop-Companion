"""
Tests for the framing loop's handoff to the cut (FalconRunDialog, loop mode).

Frame & Cut runs the framing trace in a loop. Between passes the dialog
waits for Grbl to report Idle (a jog may still be finishing), then streams
the next pass. "Looks Good — Cut!" sets _cut_requested and lets the current
pass finish; the pass's on_done then closes the dialog and the cut starts.

Matt hit the gap on 2026-10-26 (v2.80, still in v2.86): clicking Cut during
the pause BETWEEN passes. Both the Idle wait and the restart bailed out on
_cut_requested, nothing else was going to end the loop, and the dialog sat on
"Cut requested — finishing current framing pass..." with the head stopped.
Stop did nothing either: it soft-resets the stream and waits for the
sender's on_done, but between passes nothing is streaming, so none came.
Home Laser had the same blind spot; homing freed the dialog only because it
made the window respond again.

These drive the handlers on an instance built via __new__ with a scripted
fake sender and a fake after() queue — no Tk, no serial port, so this runs
headless anywhere. Each case was run against the pre-fix code first and the
handoff and Stop cases failed there.

Run:
    python tools/test_framing_loop_handoff.py
"""
import inspect
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import i18n  # noqa: E402

i18n.init_translation('en')

from ui_dialogs import FalconRunDialog  # noqa: E402

results = []


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


class FakeWidget:
    def config(self, **kw):
        pass


class FakeVar:
    def __init__(self):
        self.value = ""

    def set(self, v):
        self.value = v


class FakeSender:
    """Scripted Grbl: get_status pops from `states` (the last one repeats)."""

    def __init__(self, states=("Idle",), streaming=False):
        self.states = list(states)
        self.streaming = streaming
        self.started = 0
        self.stopped = 0

    def get_status(self, timeout=0.2):
        state = self.states.pop(0) if len(self.states) > 1 else self.states[0]
        return {'state': state}

    def is_streaming(self):
        return self.streaming

    def start_stream(self, lines):
        self.started += 1
        self.streaming = True

    def stop(self):
        self.stopped += 1
        self.streaming = False


def make_framing_dialog(sender):
    """A framing-loop dialog (cut button shown) mid-run, plus its after() queue."""
    d = FalconRunDialog.__new__(FalconRunDialog)
    queue = []
    d.after = lambda ms, fn, *args: queue.append((fn, args))
    d.winfo_exists = lambda: True
    d._destroyed = 0

    def destroy():
        d._destroyed += 1
    d.destroy = destroy
    d._finished_calls = []
    d._on_finished = lambda reason: d._finished_calls.append(reason)
    d._sender = sender
    d._loop = True
    d._show_cut_button = True
    d._stop_needs_confirm = False
    d._is_jog_only_mode = False
    d._finished = False
    d._final_reason = None
    d._cut_requested = False
    d._pass_count = 0
    d._sent = 0
    d._wait_idle_retries = 0
    d._gcode_provider = None
    d._gcode_lines = ["G92 X0 Y0", "G0 X0 Y0", "M5"]
    d._last_error = None
    d._progress = {'value': 0}
    d._state_var = FakeVar()
    d._cut_btn = d._stop_btn = d._pause_btn = d._resume_btn = FakeWidget()
    d._done_button_label = "Close"
    d._has_jogged = True
    return d, queue


def run_queue(queue, limit=100):
    """Run deferred after() calls in order until the queue drains."""
    n = 0
    while queue and n < limit:
        fn, args = queue.pop(0)
        fn(*args)
        n += 1
    assert not queue, "after() queue never drained"


def end_pass(d):
    """The sender reports a framing pass finished cleanly."""
    d._sender.streaming = False
    d._on_done("complete")


def main_test():
    print("Framing Loop -> Cut Handoff Tests")
    print("=" * 60)

    def cut_between_passes_hands_off():
        sender = FakeSender(states=["Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)                       # pass 1 done; Idle wait scheduled
        assert queue, "the loop should be waiting for Idle"
        d._on_cut_clicked()               # the click lands in the pause
        run_queue(queue)
        assert d._destroyed == 1, "dialog must close so the cut can start"
        assert d._final_reason == "complete", d._final_reason
        assert d._finished_calls == ["complete"], d._finished_calls
        assert sender.started == 0, "no extra framing pass after Cut"
    check("Cut clicked between passes ends the loop and hands off to the cut",
          cut_between_passes_hands_off)

    def cut_waits_for_idle_first():
        # A jog in flight when Cut is clicked: the cut's first line would
        # hit error:9 in Jog state, so the handoff must still wait.
        sender = FakeSender(states=["Jog", "Jog", "Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)
        d._on_cut_clicked()
        fn, args = queue.pop(0)
        fn(*args)                         # first poll: Jog
        assert d._destroyed == 0 and not d._finished, "handed off while jogging"
        run_queue(queue)
        assert d._destroyed == 1 and d._final_reason == "complete"
        assert sender.started == 0
    check("Cut between passes still waits for Idle before handing off",
          cut_waits_for_idle_first)

    def cut_during_settle_delay():
        # Idle reached, restart scheduled 150 ms out — Cut lands in that gap.
        sender = FakeSender(states=["Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)
        fn, args = queue.pop(0)
        fn(*args)                         # Idle → schedules _restart_stream
        assert queue and queue[0][0] == d._restart_stream
        d._on_cut_clicked()
        run_queue(queue)
        assert d._destroyed == 1 and d._final_reason == "complete"
        assert sender.started == 0, "restart must not stream after Cut"
    check("Cut clicked during the restart settle delay also hands off",
          cut_during_settle_delay)

    def no_cut_keeps_looping():
        sender = FakeSender(states=["Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)
        run_queue(queue)
        assert sender.started == 1, "the next framing pass must start"
        assert d._destroyed == 0 and not d._finished
    check("Without Cut the loop still restarts the next pass", no_cut_keeps_looping)

    def cut_mid_pass_unchanged():
        sender = FakeSender(streaming=True)
        d, queue = make_framing_dialog(sender)
        d._on_cut_clicked()
        assert d._destroyed == 0, "mid-pass: must let the pass finish"
        end_pass(d)
        assert d._destroyed == 1 and d._final_reason == "complete"
        assert not queue and sender.started == 0
    check("Cut clicked mid-pass still waits for that pass, then hands off",
          cut_mid_pass_unchanged)

    def stop_between_passes_ends_run():
        sender = FakeSender(states=["Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)
        d._on_stop_clicked()
        assert d._finished and d._final_reason == "stopped", (d._finished, d._final_reason)
        assert d._destroyed == 0, "Stop leaves the dialog open with Close"
        run_queue(queue)
        assert sender.started == 0, "no pass may start after Stop"
    check("Stop between passes ends the run (it used to do nothing)",
          stop_between_passes_ends_run)

    def stop_after_cut_in_gap():
        # The exact state Matt was stuck in, before the fix's handoff fires.
        sender = FakeSender(states=["Idle"])
        d, queue = make_framing_dialog(sender)
        end_pass(d)
        d._on_cut_clicked()
        d._on_stop_clicked()
        assert d._finished and d._final_reason == "stopped"
        run_queue(queue)
        assert d._destroyed == 0 and d._final_reason == "stopped", \
            "Stop wins over a pending Cut: the cut must not start"
        assert sender.started == 0
    check("Stop after a pending Cut stops; the cut does not start",
          stop_after_cut_in_gap)

    def stop_mid_pass_waits_for_sender():
        # Streaming: the soft-reset's own on_done finishes the dialog.
        # Finishing here too would report twice.
        sender = FakeSender(streaming=True)
        d, queue = make_framing_dialog(sender)
        d._on_stop_clicked()
        assert sender.stopped == 1
        assert not d._finished, "mid-pass Stop must leave finishing to on_done"
    check("Stop mid-pass leaves finishing to the sender's on_done",
          stop_mid_pass_waits_for_sender)

    def home_uses_the_same_stop():
        src = inspect.getsource(FalconRunDialog._on_home_clicked)
        assert "self._stop_stream()" in src and "self._sender.stop()" not in src, \
            "Home must abort through _stop_stream so it works between passes"
    check("Home Laser aborts the run through the same stop path",
          home_uses_the_same_stop)

    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main_test())
