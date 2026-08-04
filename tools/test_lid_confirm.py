"""
Tests for the Frame & Cut cover reminder (FalconRunDialog).

Clicking "Start Frame →" with the laser's cover open drops the Falcon into
Door/Alarm state, which costs a machine power cycle AND an app restart before
SSC sees the laser again. The advance button therefore confirms the cover is
down before handing off to the framing loop.

The reminder is unconditional by design: the jog step's own title invites
opening the lid to push the head by hand, and not every Falcon wires the door
switch through to Grbl, so a clean status is not proof the cover is closed.
When Grbl *does* report the door pin, the wording gets more definite.

These drive the handlers on an instance built via __new__, setting only the
attributes they touch — no Tk, no serial port, so this runs headless anywhere.

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


class FakeSender:
    """Minimal stand-in: a status dict plus a pollable refresh."""

    def __init__(self, status=None, poll_status=None, poll_raises=False):
        self.status = status if status is not None else {'state': 'Idle'}
        self._poll_status = poll_status
        self._poll_raises = poll_raises
        self.poll_count = 0

    def get_status(self, timeout=0.3):
        self.poll_count += 1
        if self._poll_raises:
            raise OSError("serial went away")
        if self._poll_status is not None:
            self.status = self._poll_status
        return self.status


def make_dialog(sender, confirm=True):
    """A FalconRunDialog with only the advance-path attributes set."""
    d = FalconRunDialog.__new__(FalconRunDialog)
    d._sender = sender
    d._confirm_lid_on_advance = confirm
    d._finished = True
    d._destroyed = 0

    def fake_destroy():
        d._destroyed += 1
    d.destroy = fake_destroy
    return d


class PatchAskYesNo:
    """Capture askyesno calls and feed back a canned answer."""

    def __init__(self, answer):
        self.answer = answer
        self.calls = []

    def __enter__(self):
        self._orig = ui_dialogs.messagebox.askyesno

        def fake(title, message, **kwargs):
            self.calls.append((title, message))
            return self.answer
        ui_dialogs.messagebox.askyesno = fake
        return self

    def __exit__(self, *exc):
        ui_dialogs.messagebox.askyesno = self._orig
        return False


def main_test():
    print("Frame & Cut Cover Reminder Tests")
    print("=" * 60)

    def flag_off_never_prompts():
        # Default callers use the advance button as plain "Close" on a
        # finished job — a laser prompt there would be wrong.
        d = make_dialog(FakeSender(), confirm=False)
        with PatchAskYesNo(True) as p:
            d._on_advance_clicked()
        assert not p.calls, f"should not prompt when flag is off: {p.calls}"
        assert d._destroyed == 1, "dialog should close straight through"
    check("Advance without the flag closes with no prompt", flag_off_never_prompts)

    def flag_on_prompts_and_proceeds():
        d = make_dialog(FakeSender())
        with PatchAskYesNo(True) as p:
            d._on_advance_clicked()
        assert len(p.calls) == 1, f"expected one prompt, got {len(p.calls)}"
        assert d._destroyed == 1, "Yes should let the advance through"
    check("Advance with the flag prompts, then proceeds on Yes", flag_on_prompts_and_proceeds)

    def no_keeps_dialog_open():
        d = make_dialog(FakeSender())
        with PatchAskYesNo(False) as p:
            d._on_advance_clicked()
        assert len(p.calls) == 1
        assert d._destroyed == 0, \
            "No must keep the dialog open so the cover can be closed"
    check("Answering No keeps the jog dialog open", no_keeps_dialog_open)

    def reminder_wording_when_door_unknown():
        d = make_dialog(FakeSender(status={'state': 'Idle'}))
        with PatchAskYesNo(True) as p:
            d._confirm_lid_before_advance()
        title, message = p.calls[0]
        assert title == "Cover Closed?", f"unexpected title {title!r}"
        assert "Is the cover closed?" in message, message
    check("Undetectable door state asks the plain reminder", reminder_wording_when_door_unknown)

    def definite_wording_when_state_door():
        d = make_dialog(FakeSender(status={'state': 'Door'}))
        with PatchAskYesNo(True) as p:
            d._confirm_lid_before_advance()
        title, message = p.calls[0]
        assert title == "Cover Is Open", f"unexpected title {title!r}"
        assert "reports the cover is open" in message, message
    check("Grbl 'Door' state gets the definite warning", definite_wording_when_state_door)

    def definite_wording_when_door_pin():
        # Grbl 1.1 Pn: field — 'D' means the door input is active.
        d = make_dialog(FakeSender(status={'state': 'Idle', 'pins': 'D'}))
        with PatchAskYesNo(True) as p:
            d._confirm_lid_before_advance()
        assert p.calls[0][0] == "Cover Is Open", p.calls[0][0]
    check("Door pin in the Pn: field gets the definite warning", definite_wording_when_door_pin)

    def polls_before_deciding():
        # The cached status never sees the door in jog-only mode (nothing
        # streams), so a fresh poll has to happen or the check is blind.
        sender = FakeSender(status={'state': 'Idle'},
                            poll_status={'state': 'Door'})
        d = make_dialog(sender)
        with PatchAskYesNo(True) as p:
            d._confirm_lid_before_advance()
        assert sender.poll_count == 1, \
            f"expected a fresh status poll, got {sender.poll_count}"
        assert p.calls[0][0] == "Cover Is Open", \
            "fresh poll result should drive the wording, not the stale cache"
    check("Status is re-polled before deciding the wording", polls_before_deciding)

    def poll_failure_still_prompts():
        # A dead port must not swallow the reminder.
        d = make_dialog(FakeSender(poll_raises=True))
        with PatchAskYesNo(True) as p:
            d._confirm_lid_before_advance()
        assert len(p.calls) == 1, "a failed poll must still prompt"
        assert p.calls[0][0] == "Cover Closed?", p.calls[0][0]
    check("A failed status poll still shows the reminder", poll_failure_still_prompts)

    def stop_button_cannot_bypass():
        # _on_done rewires the stop button to the advance handler; if that
        # rewire ever fails, the stop path must not skip the check.
        d = make_dialog(FakeSender())
        d._finished = True
        with PatchAskYesNo(False) as p:
            d._on_stop_clicked()
        assert len(p.calls) == 1, "stop-when-finished must run the cover check"
        assert d._destroyed == 0, "No must block this path too"
    check("Stop button's finished path cannot bypass the check", stop_button_cannot_bypass)

    def frame_and_cut_opts_in():
        # Pin the wiring: the Frame & Cut jog dialog must request it.
        import inspect
        import main
        src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
        assert "confirm_lid_on_advance=True" in src, \
            "Frame & Cut jog dialog no longer opts into the cover check"
        assert "Start Frame" in src, "expected the jog dialog in this method"
    check("Frame & Cut opts into the cover check", frame_and_cut_opts_in)

    print("=" * 60)
    passed = sum(1 for r in results if r)
    print(f"Summary: {passed}/{len(results)} passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main_test())
