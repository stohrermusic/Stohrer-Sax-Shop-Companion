"""No suite may put a native box on the screen.

A real tkinter.messagebox waits for a click: on a developer's machine it
sits on their screen until they answer it, and on a CI runner it waits
until the job times out. Both happened in one go on 2026-11-09: a planted
bug made the Tonehole Crescent's Frame & Cut ask "Engraving Not Set", the
suite driving it had stood in for askyesno but not askokcancel, the real
box came up, and once it was clicked the suite carried on and PASSED.

So a suite that drives handlers or dialogs does this before the first one
runs:

    import box_guard
    guard = box_guard.install()

install() replaces ALL EIGHT tkinter.messagebox functions (showinfo,
showwarning, showerror, askokcancel, askyesno, askyesnocancel, askquestion,
askretrycancel), the function underneath them, the native file dialogs,
simpledialog's three prompts, and at the bottom tkinter.commondialog's
Dialog.show, which every native box and file dialog finally calls, so none
can open by any route. A suite may put its own stand-ins on top afterwards;
it can never take one away.

Anything that tests the guard itself must check that a function IS the
guard's stand-in before calling it. The first version of those checks
called first, and a deliberately broken guard put two real dialogs on the
screen for seven minutes each.

The show* stand-ins record the box in guard.shown and return. A QUESTION
must be expected by the check that causes it:

    guard.expect("Engraving Not Set", answer=True)

names the box by its title and says what to answer. A question nobody
expected is answered no (or cancel), a file dialog or prompt nobody stood
in for is cancelled, and both are remembered; so is an expectation that was
never used. settle() raises for any of them and forgets them, and each
suite's check() calls settle() after the check's own function, so the check
that caused the surprise FAILS instead of answering it and moving on.

Not a test suite (run_tests.py runs test_*.py only).
"""
import tkinter.commondialog as _cd
import tkinter.filedialog as _fd
import tkinter.messagebox as _mb
import tkinter.simpledialog as _sd

SHOWS = ("showinfo", "showwarning", "showerror")
ASKS = ("askokcancel", "askyesno", "askyesnocancel", "askquestion", "askretrycancel")
ALL = SHOWS + ASKS
# What a question answers when no check expected it: no, or cancel.
DECLINE = {"askokcancel": False, "askyesno": False, "askyesnocancel": None,
           "askquestion": "no", "askretrycancel": False}
# Native dialogs that are not message boxes but wait just the same, and
# what "cancelled" looks like for each.
OTHER_DIALOGS = (
    (_fd, "askopenfilename", ""), (_fd, "askopenfilenames", ()), (_fd, "asksaveasfilename", ""),
    (_fd, "askdirectory", ""), (_fd, "askopenfile", None), (_fd, "askopenfiles", []),
    (_fd, "asksaveasfile", None),
    (_sd, "askstring", None), (_sd, "askinteger", None), (_sd, "askfloat", None),
)


class Guard:
    def __init__(self):
        self.installed = False
        self.shown = []         # (kind, title, message) of every box, questions included, in order
        self.asked = []         # (kind, title, message, answer) of every question
        self.unexpected = []    # what no check expected, as text
        self._expected = []     # [title, answer, kind or None]

    # ---- the stand-ins -----------------------------------------------------
    def _show(self, kind):
        def stand_in(title=None, message=None, **options):
            self.shown.append((kind, str(title), str(message)))
            return "ok"
        stand_in.__name__ = kind
        return stand_in

    def _ask(self, kind):
        def stand_in(title=None, message=None, **options):
            title_s, message_s = str(title), str(message)
            self.shown.append((kind, title_s, message_s))
            for i, (want_title, answer, want_kind) in enumerate(self._expected):
                if want_title == title_s and want_kind in (None, kind):
                    del self._expected[i]
                    break
            else:
                answer = DECLINE[kind]
                self.unexpected.append(f"{kind} {title_s!r}: {message_s[:120]!r}")
                self.asked.append((kind, title_s, message_s, answer))
                return answer
            if kind == "askquestion" and isinstance(answer, bool):
                answer = "yes" if answer else "no"
            self.asked.append((kind, title_s, message_s, answer))
            return answer
        stand_in.__name__ = kind
        return stand_in

    def _other(self, name, cancelled):
        def stand_in(*args, **options):
            self.unexpected.append(f"{name}({options.get('title') or (args[0] if args else '')!r}): "
                                   "a native dialog no stand-in was given for")
            return cancelled
        stand_in.__name__ = name
        return stand_in

    def _native(self, title=None, message=None, _icon=None, _type=None, **options):
        """Underneath all eight: reached only by a box built some other way,
        or by a function a suite put back. Never opens anything."""
        self.unexpected.append(f"a NATIVE message box was about to open, {str(title)!r}: {str(message)[:120]!r}")
        return {"ok": "ok", "yesno": "no"}.get(str(_type), "cancel")

    def _common_dialog(self):
        """The bottom layer: tkinter.commondialog.Dialog.show is what every
        native message box, file dialog and colour chooser finally calls.
        With this in place none of them can open by any route."""
        guard = self

        def show(dialog, **options):
            kind = type(dialog).__name__
            guard.unexpected.append(f"a NATIVE {kind} dialog was about to open ({dialog.options.get('title')!r})")
            return (None, None) if kind == "Chooser" else ""
        return show

    # ---- for the suites ----------------------------------------------------
    def install(self):
        for kind in SHOWS:
            setattr(_mb, kind, self._show(kind))
        for kind in ASKS:
            setattr(_mb, kind, self._ask(kind))
        _mb._show = self._native
        _cd.Dialog.show = self._common_dialog()
        for module, name, cancelled in OTHER_DIALOGS:
            setattr(module, name, self._other(name, cancelled))
        self.installed = True
        return self

    def expect(self, title, answer=True, times=1, kind=None):
        """The check is about to cause this question (by its title); answer it so."""
        for _unused in range(times):
            self._expected.append([str(title), answer, kind])

    def reset(self):
        """Forget what is pending, and say what it was: a check that failed
        for its own reason prints this after its message, since a question
        nobody expected is very often why."""
        surprises = list(self.unexpected)
        unused = [f"{kind or 'a question'} {title!r}" for title, _answer, kind in self._expected]
        del self.unexpected[:], self._expected[:]
        problems = []
        if surprises:
            problems.append("not expected by this check: " + "; ".join(surprises))
        if unused:
            problems.append("expected but never asked: " + "; ".join(unused))
        return " | ".join(problems)

    def settle(self):
        """Raise if a question was asked that no check expected, or one was
        expected and never asked. Either way start clean for the next check."""
        problems = self.reset()
        if problems:
            raise AssertionError(problems)


GUARD = Guard()


def install():
    return GUARD.install()


def settle():
    if GUARD.installed:
        GUARD.settle()


def reset():
    """For a check that failed: forget what is pending and return it as
    text to print after the failure (empty when there was nothing)."""
    left = GUARD.reset()
    return f" [{left}]" if left else ""
