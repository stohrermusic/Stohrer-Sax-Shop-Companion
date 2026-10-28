"""
Tests for the laser lid question (FalconRunDialog._remind_lid_closed).

Firing the laser with the lid open drops the Falcon into Door/Alarm state,
which costs a machine power cycle AND an app restart before SSC sees the
laser again. FalconRunDialog asks "Is the lid closed?" at exactly ONE
boundary: **Start Frame**, the jog-to-position dialog's advance button,
opted into by the caller via remind_lid_on_advance.

The question has two real buttons, "Yes, Start Frame" and "No, Cancel"
(LidClosedDialog; Matt, 2026-10-26). Yes advances; No leaves the jog
dialog open so the user can shut the lid and click Start Frame again, with
nothing consumed. From 2026-08-05 to 2026-10-26 it was a single-OK
reminder with nothing to decline -- the thinking being that a yes/no gate
just gets clicked through -- but in the shop the lone button read as "the
app thinks the lid is closed" rather than as a question.

Start Frame is where the laser first comes on. Framing runs at low power,
but low power still energises the beam, so Grbl enters Door state right
there. It briefly prompted on the *cut* instead (2026-08-11), on the
reasoning that framing "isn't about to start burning" -- that confused
marking the material with firing the laser, and by the time the cut prompt
appeared the user had already eaten the alarm. Corrected 2026-08-18.

It used to be gated by _is_door_open(), which read Grbl's status for Door
state or a Pn: 'D' pin and only asked when it couldn't tell. Retired after
live-hardware testing on 2026-08-05: the Falcon2 Pro 40W (ESP32-S3
controller) never sends a Pn: field at all, lid open or closed, so
detection was unverifiable and added a pointless synchronous status poll.
Same question every time.

These drive the handlers on an instance built via __new__, setting only the
attributes each path touches, with the dialog class replaced by a fake that
answers from a script -- no Tk, no serial port, so this runs headless
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


class PatchLidDialog:
    """Replace LidClosedDialog with a fake that records what it was asked
    and answers from a script (True = Yes, Start Frame; False = No,
    Cancel). Default answer: Yes."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = []

    def __enter__(self):
        outer = self
        self._orig = ui_dialogs.LidClosedDialog

        class Fake:
            def __init__(self, parent, title, message, yes_label, no_label):
                outer.calls.append((title, message, yes_label, no_label))
                self.result = outer.answers.pop(0) if outer.answers else True
        ui_dialogs.LidClosedDialog = Fake
        return self

    def __exit__(self, *exc):
        ui_dialogs.LidClosedDialog = self._orig
        return False


class FailOnMessageBox:
    """The question is the custom dialog, never a tk messagebox -- fail
    loudly if the reminder reaches for one (showinfo has no answer to
    give, askyesno can't label its buttons)."""

    def __enter__(self):
        self._orig = (ui_dialogs.messagebox.showinfo, ui_dialogs.messagebox.askyesno)

        def fake(*a, **kw):
            raise AssertionError("lid question must use LidClosedDialog, not a messagebox")
        ui_dialogs.messagebox.showinfo = fake
        ui_dialogs.messagebox.askyesno = fake
        return self

    def __exit__(self, *exc):
        ui_dialogs.messagebox.showinfo, ui_dialogs.messagebox.askyesno = self._orig
        return False


def main_test():
    print("Laser Lid Question Tests")
    print("=" * 60)

    # ------------------------------------------------------------
    # The question itself
    # ------------------------------------------------------------

    def question_wording_and_buttons_are_constant():
        d = make_dialog()
        with FailOnMessageBox(), PatchLidDialog(True) as p:
            assert d._remind_lid_closed() is True
        title, message, yes_label, no_label = p.calls[0]
        assert title == "Lid Closed?", f"unexpected title {title!r}"
        assert message.strip() == "Is the lid closed?", (
            f"the question should be the bare question, got {message!r}")
        assert yes_label == "Yes, Start Frame", yes_label
        assert no_label == "No, Cancel", no_label
    check("Question reads 'Is the lid closed?' with 'Yes, Start Frame' / 'No, Cancel'",
          question_wording_and_buttons_are_constant)

    def no_answers_false():
        d = make_dialog()
        with FailOnMessageBox(), PatchLidDialog(False):
            assert d._remind_lid_closed() is False
    check("'No, Cancel' comes back as False", no_answers_false)

    def no_status_poll_happens():
        # Detection is gone entirely -- asking must not touch
        # self._sender at all, let alone poll it.
        d = make_dialog()

        class ExplodingSender:
            def get_status(self, *a, **kw):
                raise AssertionError("reminder must not poll sender status")
        d._sender = ExplodingSender()
        with PatchLidDialog(True):
            d._remind_lid_closed()
    check("Question never touches the sender", no_status_poll_happens)

    # ------------------------------------------------------------
    # Start Frame → (_on_advance_clicked)
    # ------------------------------------------------------------

    def advance_asks_when_the_caller_opts_in():
        # Start Frame hands off to the framing loop, which FIRES THE
        # LASER. An open lid here is the alarm Matt hit in the shop.
        d = make_dialog(_remind_lid_on_advance=True, _final_reason="complete")
        with PatchLidDialog(True) as p:
            d._on_advance_clicked()
        assert len(p.calls) == 1, f"Start Frame must ask once: {p.calls}"
        assert d._destroyed == 1, "Yes should close the jog dialog so the caller proceeds"
        assert d._final_reason == "complete", "Yes must not clobber the caller's proceed signal"
    check("Advance (Start Frame) asks when the caller opts in; Yes proceeds",
          advance_asks_when_the_caller_opts_in)

    def no_keeps_the_jog_dialog_open():
        # The lid is up. Stay put: the user shuts it and clicks Start
        # Frame again. Nothing is destroyed, nothing is cancelled.
        d = make_dialog(_remind_lid_on_advance=True, _final_reason="complete")
        with PatchLidDialog(False) as p:
            d._on_advance_clicked()
        assert len(p.calls) == 1
        assert d._destroyed == 0, "No must leave the jog dialog open"
        assert d._final_reason == "complete", "No must not turn into a cancel"
        # Second click after closing the lid: asks again, then proceeds.
        with PatchLidDialog(True) as p2:
            d._on_advance_clicked()
        assert len(p2.calls) == 1 and d._destroyed == 1
    check("'No, Cancel' keeps the jog dialog open; the next Start Frame asks again",
          no_keeps_the_jog_dialog_open)

    def advance_stays_silent_by_default():
        # The same button on a dialog whose next step doesn't fire the
        # laser (plain "Close" after a finished stream) must not ask.
        d = make_dialog(_remind_lid_on_advance=False)
        with PatchLidDialog() as p:
            d._on_advance_clicked()
        assert not p.calls, f"plain advance must not ask: {p.calls}"
        assert d._destroyed == 1
    check("Advance stays silent when the caller did not opt in",
          advance_stays_silent_by_default)

    def stop_button_finished_path_does_not_ask():
        d = make_dialog(_finished=True, _remind_lid_on_advance=False)
        with PatchLidDialog() as p:
            d._on_stop_clicked()
        assert not p.calls, f"finished-Stop must not ask: {p.calls}"
        assert d._destroyed == 1
    check("Stop button's finished path does not ask",
          stop_button_finished_path_does_not_ask)

    # ------------------------------------------------------------
    # Cut ("Looks Good — Cut!")
    # ------------------------------------------------------------

    def cut_click_does_not_ask():
        # By the cut, framing has been running -- the lid is
        # demonstrably shut. Asking again is the second prompt that
        # trains the user to dismiss it.
        d = make_dialog(_cut_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchLidDialog() as p:
            d._on_cut_clicked()
        assert not p.calls, f"cut must not ask again: {p.calls}"
        assert d._cut_requested is True
    check("Cut click flags the request without asking again",
          cut_click_does_not_ask)

    # ------------------------------------------------------------
    # Done during a framing-only loop (no cut button)
    # ------------------------------------------------------------

    def graceful_stop_does_not_ask():
        # Done during a framing-only loop just lets the in-flight pass
        # finish and exits -- nothing new is about to fire.
        d = make_dialog(_finished=False, _loop=True, _show_cut_button=False,
                         _stop_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchLidDialog() as p:
            d._on_stop_clicked()
        assert not p.calls, f"graceful Done must not ask: {p.calls}"
        assert d._cut_requested is True
    check("Graceful Done (framing-only loop) does not ask",
          graceful_stop_does_not_ask)

    def start_frame_is_the_only_question_in_a_frame_and_cut_run():
        # One question per run, at the boundary that fires the laser.
        asked = []
        jog = make_dialog(_remind_lid_on_advance=True, _final_reason="complete")
        with PatchLidDialog(True) as p:
            jog._on_advance_clicked()          # Start Frame
        asked += p.calls
        framing = make_dialog(_loop=True, _show_cut_button=True,
                               _cut_btn=FakeWidget(), _state_var=FakeWidget())
        with PatchLidDialog() as p:
            framing._on_cut_clicked()          # Looks Good — Cut!
        asked += p.calls
        assert len(asked) == 1, \
            f"a Frame & Cut run must ask exactly once, got {len(asked)}"
    check("A full Frame & Cut run asks exactly once (on Start Frame)",
          start_frame_is_the_only_question_in_a_frame_and_cut_run)

    def aggressive_stop_does_not_ask():
        # Cut-context Stop is an abort, not a "fires the laser" handoff --
        # it must go straight to the existing Stop-confirmation, not the
        # lid question.
        calls = []

        class Sender:
            def stop(self):
                calls.append('stopped')
        d = make_dialog(_finished=False, _loop=False, _show_cut_button=True,
                         _stop_needs_confirm=False, _sender=Sender())
        with PatchLidDialog() as p:
            d._on_stop_clicked()
        assert not p.calls, f"aggressive stop should not ask about the lid: {p.calls}"
        assert calls == ['stopped']
    check("Aggressive Stop (mid-cut abort) does not ask about the lid",
          aggressive_stop_does_not_ask)

    # ------------------------------------------------------------
    # Wiring
    # ------------------------------------------------------------

    def only_the_advance_path_asks():
        import inspect
        src = inspect.getsource(FalconRunDialog)
        # One definition plus exactly one call site (_on_advance_clicked).
        assert src.count("_remind_lid_closed") == 2, (
            "the lid question should have exactly one call site (Start "
            f"Frame); found {src.count('_remind_lid_closed') - 1}")
        adv_src = inspect.getsource(FalconRunDialog._on_advance_clicked)
        assert "_remind_lid_closed" in adv_src, \
            "the Start Frame path must ask"
        cut_src = inspect.getsource(FalconRunDialog._on_cut_clicked)
        assert "_remind_lid_closed" not in cut_src, \
            "the cut must not ask a second time"
    check("Only the Start Frame path asks", only_the_advance_path_asks)

    def frame_and_cut_opts_into_the_advance_question():
        import inspect
        import main
        sig = inspect.signature(FalconRunDialog.__init__)
        assert 'remind_lid_on_advance' in sig.parameters, \
            "the advance-question flag must exist"
        assert sig.parameters['remind_lid_on_advance'].default is False, \
            "the flag must be opt-in: only laser-firing advances ask"
        src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
        assert "remind_lid_on_advance=True" in src, \
            "Frame & Cut's jog dialog must opt into the lid question"
        assert "Start Frame" in src, "expected the jog dialog in this method"
    check("Frame & Cut opts into the Start Frame question",
          frame_and_cut_opts_into_the_advance_question)

    def dialog_class_has_the_two_answers():
        # The real dialog: a Toplevel with a `result` that the two
        # handlers set, closing either way.
        cls = ui_dialogs.LidClosedDialog
        import tkinter as tk
        assert issubclass(cls, tk.Toplevel)
        assert callable(getattr(cls, "_on_yes", None)) and callable(getattr(cls, "_on_no", None))
        probe = cls.__new__(cls)
        closed = []
        probe.destroy = lambda: closed.append(True)
        probe._on_no()
        assert probe.result is False and closed == [True]
        probe._on_yes()
        assert probe.result is True and len(closed) == 2
    check("LidClosedDialog answers True for Yes and False for No, closing either way",
          dialog_class_has_the_two_answers)

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
