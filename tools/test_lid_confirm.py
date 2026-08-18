"""
Tests for the laser lid reminder (FalconRunDialog._remind_lid_closed).

Firing the laser with the lid open drops the Falcon into Door/Alarm state,
which costs a machine power cycle AND an app restart before SSC sees the
laser again. FalconRunDialog shows a plain "Is the lid closed?" reminder --
a single-OK popup, not a yes/no gate -- at every boundary that hands off to
something that fires the laser: Start Frame (the jog-to-position dialog's
advance button, gated by confirm_lid_on_advance), Cut ("Looks Good — Cut!"
during a loop with a cut button), and Done during a framing-only loop (no
cut button -- e.g. Camera Calibration's continuous framing trace).

All three used to be gated by _is_door_open(), which read Grbl's status for
Door state or a Pn: 'D' pin and only asked when it couldn't tell. Retired
after live-hardware testing on 2026-08-05: the Falcon2 Pro 40W (ESP32-S3
controller) never sends a Pn: field at all, lid open or closed, so
detection was unverifiable and added a pointless synchronous status poll.
Matt asked for the simple version: same reminder, every time, OK proceeds.

These drive the handlers on an instance built via __new__, setting only the
attributes each path touches — no Tk, no serial port, so this runs headless
anywhere.

Run:
    python tools/test_lid_confirm.py
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import i18n  # noqa: E402

i18n.init_translation('en')

import ui_dialogs  # noqa: E402
from ui_dialogs import FalconRunDialog  # noqa: E402

results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


class FakeWidget:
    """Stands in for a tk.Button/tk.StringVar — config()/set() are no-ops."""

    def config(self, **kw):
        pass

    def set(self, *a, **kw):
        pass


def make_dialog(**attrs):
    """A FalconRunDialog with only the attributes a given handler touches."""
    d = FalconRunDialog.__new__(FalconRunDialog)
    d._destroyed = 0

    def fake_destroy():
        d._destroyed += 1
    d.destroy = fake_destroy
    for k, v in attrs.items():
        setattr(d, k, v)
    return d


class PatchShowInfo:
    """Capture showinfo calls (the reminder has no return value to fake)."""

    def __init__(self):
        self.calls = []

    def __enter__(self):
        self._orig = ui_dialogs.messagebox.showinfo

        def fake(title, message, **kwargs):
            self.calls.append((title, message))
        ui_dialogs.messagebox.showinfo = fake
        return self

    def __exit__(self, *exc):
        ui_dialogs.messagebox.showinfo = self._orig
        return False


class FailOnAskYesNo:
    """The reminder must not be a yes/no gate -- fail loudly if it asks one."""

    def __enter__(self):
        self._orig = ui_dialogs.messagebox.askyesno

        def fake(*a, **kw):
            raise AssertionError("lid reminder must not use askyesno")
        ui_dialogs.messagebox.askyesno = fake
        return self

    def __exit__(self, *exc):
        ui_dialogs.messagebox.askyesno = self._orig
        return False


def main_test():
    print("Laser Lid Reminder Tests")
    print("=" * 60)

    # ------------------------------------------------------------
    # The reminder itself
    # ------------------------------------------------------------

    def reminder_wording_is_constant():
        d = make_dialog()
        with PatchShowInfo() as p:
            d._remind_lid_closed()
        title, message = p.calls[0]
        assert title == "Lid Closed?", f"unexpected title {title!r}"
        assert message.strip() == "Is the lid closed?", (
            f"reminder should be the bare question, got {message!r}")
    check("Reminder uses the plain 'Lid Closed?' wording", reminder_wording_is_constant)

    def reminder_is_not_a_yes_no_gate():
        # Regression guard for the design Matt explicitly asked to move
        # away from: a decision popup that can be answered "No."
        d = make_dialog()
        with FailOnAskYesNo(), PatchShowInfo():
            d._remind_lid_closed()
    check("Reminder never calls askyesno (single OK button only)",
          reminder_is_not_a_yes_no_gate)

    def no_status_poll_happens():
        # Detection is gone entirely -- showing the reminder must not
        # touch self._sender at all, let alone poll it.
        d = make_dialog()

        class ExplodingSender:
            def get_status(self, *a, **kw):
                raise AssertionError("reminder must not poll sender status")
        d._sender = ExplodingSender()
        with PatchShowInfo():
            d._remind_lid_closed()
    check("Reminder never touches the sender", no_status_poll_happens)

    # ------------------------------------------------------------
    # Start Frame → (_on_advance_clicked)
    # ------------------------------------------------------------

    def advance_never_prompts():
        # Start Frame hands off to the framing loop, which runs at low
        # power. Prompting here made Frame & Cut ask twice in one run,
        # which just trains the user to click it away.
        d = make_dialog()
        with PatchShowInfo() as p:
            d._on_advance_clicked()
        assert not p.calls, f"advance must not prompt: {p.calls}"
        assert d._destroyed == 1, "dialog should close straight through"
    check("Advance (Start Frame) closes with no prompt", advance_never_prompts)

    def stop_button_finished_path_does_not_prompt():
        d = make_dialog(_finished=True)
        with PatchShowInfo() as p:
            d._on_stop_clicked()
        assert not p.calls, f"finished-Stop must not prompt: {p.calls}"
        assert d._destroyed == 1
    check("Stop button's finished path does not prompt",
          stop_button_finished_path_does_not_prompt)

    # ------------------------------------------------------------
    # Cut ("Looks Good — Cut!")
    # ------------------------------------------------------------

    def cut_click_shows_reminder():
        d = make_dialog(_cut_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchShowInfo() as p:
            d._on_cut_clicked()
        assert len(p.calls) == 1, f"expected one reminder, got {len(p.calls)}"
        assert d._cut_requested is True
    check("Cut click shows the reminder and flags the cut request",
          cut_click_shows_reminder)

    # ------------------------------------------------------------
    # Done during a framing-only loop (no cut button)
    # ------------------------------------------------------------

    def graceful_stop_does_not_remind():
        # Done during a framing-only loop just lets the in-flight pass
        # finish and exits -- nothing new is about to fire.
        d = make_dialog(_finished=False, _loop=True, _show_cut_button=False,
                         _stop_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchShowInfo() as p:
            d._on_stop_clicked()
        assert not p.calls, f"graceful Done must not prompt: {p.calls}"
        assert d._cut_requested is True
    check("Graceful Done (framing-only loop) does not prompt",
          graceful_stop_does_not_remind)

    def cut_is_the_only_reminder_in_a_frame_and_cut_run():
        # The whole point of the change: one prompt per run, on the cut.
        prompts = []
        jog = make_dialog()
        with PatchShowInfo() as p:
            jog._on_advance_clicked()          # Start Frame
        prompts += p.calls
        framing = make_dialog(_loop=True, _show_cut_button=True,
                               _cut_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchShowInfo() as p:
            framing._on_cut_clicked()          # Looks Good — Cut!
        prompts += p.calls
        assert len(prompts) == 1, \
            f"a Frame & Cut run must prompt exactly once, got {len(prompts)}"
    check("A full Frame & Cut run prompts exactly once (on the cut)",
          cut_is_the_only_reminder_in_a_frame_and_cut_run)

    def aggressive_stop_does_not_remind():
        # Cut-context Stop is an abort, not a "fires the laser" handoff --
        # it must go straight to the existing Stop-confirmation, not the
        # lid reminder.
        calls = []

        class Sender:
            def stop(self):
                calls.append('stopped')
        d = make_dialog(_finished=False, _loop=False, _show_cut_button=True,
                         _stop_needs_confirm=False, _sender=Sender())
        with PatchShowInfo() as p:
            d._on_stop_clicked()
        assert not p.calls, f"aggressive stop should not show the lid reminder: {p.calls}"
        assert calls == ['stopped']
    check("Aggressive Stop (mid-cut abort) does not show the lid reminder",
          aggressive_stop_does_not_remind)

    # ------------------------------------------------------------
    # Wiring
    # ------------------------------------------------------------

    def only_the_cut_path_calls_the_reminder():
        import inspect
        src = inspect.getsource(FalconRunDialog)
        # One definition plus exactly one call site (_on_cut_clicked).
        assert src.count("_remind_lid_closed") == 2, (
            "the lid reminder should have exactly one call site (the cut); "
            f"found {src.count('_remind_lid_closed') - 1}")
        cut_src = inspect.getsource(FalconRunDialog._on_cut_clicked)
        assert "_remind_lid_closed" in cut_src, "the cut path must still remind"
    check("Only the cut path calls the reminder", only_the_cut_path_calls_the_reminder)

    def advance_flag_is_gone():
        import inspect
        import main
        sig = inspect.signature(FalconRunDialog.__init__)
        assert 'confirm_lid_on_advance' not in sig.parameters, \
            "the advance-prompt flag should have been removed, not left unused"
        src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
        assert "confirm_lid_on_advance" not in src, \
            "Frame & Cut still passes the retired flag"
        assert "Start Frame" in src, "expected the jog dialog in this method"
    check("The advance-prompt flag is gone from both sides", advance_flag_is_gone)

    def old_names_are_gone():
        # Regression guard: the detection-based helpers should not come back.
        for name in ('_is_door_open', '_confirm_door_closed', '_confirm_lid_before_advance'):
            assert not hasattr(FalconRunDialog, name), \
                f"{name} should have been removed, not just unused"
        assert hasattr(FalconRunDialog, '_remind_lid_closed')
    check("Retired detection helpers are actually gone", old_names_are_gone)

    print("=" * 60)
    passed = sum(1 for r in results if r)
    print(f"Summary: {passed}/{len(results)} passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main_test())
