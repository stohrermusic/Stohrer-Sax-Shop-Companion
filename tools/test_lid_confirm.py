"""
Tests for the Frame & Cut lid reminder (FalconRunDialog).

Clicking "Start Frame →" with the laser's lid open drops the Falcon into
Door/Alarm state, which costs a machine power cycle AND an app restart before
SSC sees the laser again. The advance button shows a plain "Is the lid
closed?" reminder first -- a single-OK popup, not a yes/no gate. Clicking the
only button is what starts framing.

An earlier version tried to detect the door via Grbl status (Door state or a
Pn: 'D' pin) and only asked when it couldn't tell. Live-hardware testing on
2026-08-05 found the Falcon2 Pro 40W (ESP32-S3 controller) never sends a Pn:
field at all, lid open or closed -- so detection was unverifiable and added a
pointless synchronous status poll. Matt asked for the simple version: same
reminder, every time, OK starts it.

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


def make_dialog(confirm=True):
    """A FalconRunDialog with only the advance-path attributes set."""
    d = FalconRunDialog.__new__(FalconRunDialog)
    d._confirm_lid_on_advance = confirm
    d._finished = True
    d._destroyed = 0

    def fake_destroy():
        d._destroyed += 1
    d.destroy = fake_destroy
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
    print("Frame & Cut Lid Reminder Tests")
    print("=" * 60)

    def flag_off_never_prompts():
        # Default callers use the advance button as plain "Close" on a
        # finished job — a laser prompt there would be wrong.
        d = make_dialog(confirm=False)
        with PatchShowInfo() as p:
            d._on_advance_clicked()
        assert not p.calls, f"should not prompt when flag is off: {p.calls}"
        assert d._destroyed == 1, "dialog should close straight through"
    check("Advance without the flag closes with no prompt", flag_off_never_prompts)

    def flag_on_shows_reminder_then_proceeds():
        d = make_dialog()
        with PatchShowInfo() as p:
            d._on_advance_clicked()
        assert len(p.calls) == 1, f"expected one reminder, got {len(p.calls)}"
        assert d._destroyed == 1, \
            "OK is the only button -- dismissing it must always proceed"
    check("Advance with the flag shows the reminder, then always proceeds",
          flag_on_shows_reminder_then_proceeds)

    def reminder_wording_is_constant():
        d = make_dialog()
        with PatchShowInfo() as p:
            d._confirm_lid_before_advance()
        title, message = p.calls[0]
        assert title == "Lid Closed?", f"unexpected title {title!r}"
        assert "Is the laser's lid closed?" in message, message
    check("Reminder uses the plain 'Lid Closed?' wording", reminder_wording_is_constant)

    def reminder_is_not_a_yes_no_gate():
        # Regression guard for the design Matt explicitly asked to move
        # away from: a decision popup that can be answered "No."
        d = make_dialog()
        with FailOnAskYesNo(), PatchShowInfo():
            d._confirm_lid_before_advance()
    check("Reminder never calls askyesno (single OK button only)",
          reminder_is_not_a_yes_no_gate)

    def no_status_poll_happens(monkeypatch_free=True):
        # Detection is gone entirely -- confirming the popup must not
        # touch self._sender at all, let alone poll it.
        d = make_dialog()

        class ExplodingSender:
            def get_status(self, *a, **kw):
                raise AssertionError("reminder must not poll sender status")
        d._sender = ExplodingSender()
        with PatchShowInfo():
            d._confirm_lid_before_advance()
    check("Confirming the reminder never touches the sender", no_status_poll_happens)

    def stop_button_finished_path_still_reminds():
        # _on_done rewires the stop button to the advance handler; if that
        # rewire ever fails, the stop path must still show the reminder.
        d = make_dialog()
        d._finished = True
        with PatchShowInfo() as p:
            d._on_stop_clicked()
        assert len(p.calls) == 1, "stop-when-finished must still show the reminder"
        assert d._destroyed == 1
    check("Stop button's finished path also shows the reminder",
          stop_button_finished_path_still_reminds)

    def frame_and_cut_opts_in():
        # Pin the wiring: the Frame & Cut jog dialog must request it.
        import inspect
        import main
        src = inspect.getsource(main.PadSVGGeneratorApp.on_frame_and_cut)
        assert "confirm_lid_on_advance=True" in src, \
            "Frame & Cut jog dialog no longer opts into the lid reminder"
        assert "Start Frame" in src, "expected the jog dialog in this method"
    check("Frame & Cut opts into the lid reminder", frame_and_cut_opts_in)

    print("=" * 60)
    passed = sum(1 for r in results if r)
    print(f"Summary: {passed}/{len(results)} passed")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main_test())
