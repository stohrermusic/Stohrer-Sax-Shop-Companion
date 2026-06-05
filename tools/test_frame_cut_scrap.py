"""
Tests for scrap-mode Frame & Cut wiring (main.py).

Frame & Cut used to refuse scrap mode outright ("isn't wired up for scrap
mode yet"). It now mirrors the file-export flow one scrap per click via the
shared helper _scrap_begin_partial(), and commits the scrap (decrement +
continue dialog) only after a cut completes, via _frame_cut_scrap_advance().

These exercise the session bookkeeping directly — no Falcon, no real
dialogs. GUI entry points (filedialog / messagebox / the continue dialog /
the remaining-pads window) are monkeypatched so the logic runs headless.

Constructs the full app in a withdrawn Tk root, so it needs a display
(fine on Windows / macOS / CI Windows; headless Linux self-skips).

Run:
    python tools/test_frame_cut_scrap.py
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import main  # noqa: E402

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


def main_test():
    print("Frame & Cut Scrap Mode Tests")
    print("=" * 60)

    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    from main import PadSVGGeneratorApp
    a = PadSVGGeneratorApp(root)

    # --- Neutralize GUI side-effects so the helpers run headless. The
    # session bookkeeping is what we're testing, not the windowing. ---
    continue_calls = []
    info_msgs = []
    warn_msgs = []
    error_msgs = []
    askdir_calls = []

    a._show_scrap_continue_dialog = lambda *args: continue_calls.append(args)
    a._open_remaining_pads_window = lambda: None
    a._update_remaining_pads_window = lambda: None
    # _update_scrap_status_display touches real widgets and is safe, but
    # stub it too so a layout change can't make these logic tests flaky.
    a._update_scrap_status_display = lambda: None

    main.messagebox.showinfo = lambda *a_, **k: info_msgs.append((a_, k))
    main.messagebox.showwarning = lambda *a_, **k: warn_msgs.append((a_, k))
    main.messagebox.showerror = lambda *a_, **k: error_msgs.append((a_, k))
    main.messagebox.askyesno = lambda *a_, **k: False  # never opt into optimize

    def set_askdir(value):
        askdir_calls.clear()

        def _fake(**k):
            askdir_calls.append(k)
            return value
        main.filedialog.askdirectory = _fake

    def reset_session():
        a.scrap_session = {
            'active': False, 'original_pads': [], 'remaining_pads': [],
            'scrap_count': 0, 'material': None, 'save_dir': '',
            'hole_dia': 0, 'optimize': None,
        }

    def clear_msgs():
        continue_calls.clear()
        info_msgs.clear()
        warn_msgs.clear()
        error_msgs.clear()

    # ----------------------------------------------------------------
    # Frame & Cut starts a session WITHOUT prompting for a save folder.
    # ----------------------------------------------------------------
    def t_frame_cut_no_save_dir():
        reset_session()
        clear_msgs()
        set_askdir("/should/not/be/used")
        pads = [{'size': 16.0, 'qty': 4}]
        result = a._scrap_begin_partial(
            pads, 1.5, 'felt', 80, 80, None, ask_save_dir=False)
        assert result is not None, "expected a placement, got None"
        placed, remaining = result
        assert a.scrap_session['active'] is True, "session should be active"
        assert a.scrap_session['save_dir'] == '', \
            f"save_dir should stay empty, got {a.scrap_session['save_dir']!r}"
        assert askdir_calls == [], "Frame & Cut must not prompt for a folder"
        assert a.scrap_session['hole_dia'] == 1.5, "hole_dia not recorded"
        assert len(placed) >= 1, "at least one pad should place on 80x80"
    check("Frame & Cut starts session without folder prompt", t_frame_cut_no_save_dir)

    # ----------------------------------------------------------------
    # A completed cut decrements the session; a second click continues it.
    # ----------------------------------------------------------------
    def t_continue_and_decrement():
        reset_session()
        clear_msgs()
        set_askdir(None)
        # Many pads on a small scrap so not all fit -> remaining non-empty.
        pads = [{'size': 16.0, 'qty': 60}]
        r1 = a._scrap_begin_partial(
            pads, 1.0, 'felt', 55, 55, None, ask_save_dir=False)
        assert r1 is not None
        placed1, remaining1 = r1
        n1 = len(placed1)
        assert n1 >= 1, "first scrap should place some pads"
        assert n1 < 60, "small scrap shouldn't swallow all 60 pads"

        a._frame_cut_scrap_advance(n1, remaining1)
        assert a.scrap_session['scrap_count'] == 1, "scrap_count should be 1"
        left1 = sum(p['qty'] for p in a.scrap_session['remaining_pads'])
        assert left1 == 60 - n1, f"expected {60 - n1} left, got {left1}"
        assert len(continue_calls) == 1, "continue dialog should show (pads remain)"

        # Second click continues the SAME session, nesting the remainder,
        # not the original 60 (passing the original pads again must be
        # ignored in favor of the session's remaining_pads).
        clear_msgs()
        r2 = a._scrap_begin_partial(
            pads, 1.0, 'felt', 55, 55, None, ask_save_dir=False)
        assert r2 is not None
        placed2, remaining2 = r2
        n2 = len(placed2)
        assert n2 <= left1, "second scrap can't place more than what's left"
        left2 = sum(p['qty'] for p in remaining2)
        assert left2 == left1 - n2, \
            f"continuation math wrong: {left1} - {n2} != {left2}"
    check("Completed cut decrements; next click continues session", t_continue_and_decrement)

    # ----------------------------------------------------------------
    # When the last pad is placed, advance announces completion and does
    # NOT show the continue dialog.
    # ----------------------------------------------------------------
    def t_completion_message():
        reset_session()
        clear_msgs()
        set_askdir(None)
        pads = [{'size': 16.0, 'qty': 2}]  # both fit on 80x80
        r = a._scrap_begin_partial(
            pads, 1.0, 'felt', 80, 80, None, ask_save_dir=False)
        assert r is not None
        placed, remaining = r
        assert sum(p['qty'] for p in remaining) == 0, "both pads should fit"
        a._frame_cut_scrap_advance(len(placed), remaining)
        assert len(continue_calls) == 0, "no continue dialog when done"
        assert len(info_msgs) == 1, "should show one completion message"
    check("Session-complete shows completion, not continue dialog", t_completion_message)

    # ----------------------------------------------------------------
    # Material mismatch on a continuing session aborts with an error.
    # ----------------------------------------------------------------
    def t_material_mismatch():
        reset_session()
        clear_msgs()
        set_askdir(None)
        pads = [{'size': 16.0, 'qty': 4}]
        a._scrap_begin_partial(pads, 1.0, 'felt', 80, 80, None, ask_save_dir=False)
        # Now ask for a different material on the active session.
        out = a._scrap_begin_partial(pads, 1.0, 'card', 80, 80, None, ask_save_dir=False)
        assert out is None, "material switch mid-session must abort"
        assert len(error_msgs) == 1, "should warn about material mismatch"
    check("Material mismatch mid-session aborts", t_material_mismatch)

    # ----------------------------------------------------------------
    # File-export path still prompts for a folder on a NEW session.
    # ----------------------------------------------------------------
    def t_file_mode_prompts():
        reset_session()
        clear_msgs()
        set_askdir("/tmp/scrap_out")
        pads = [{'size': 16.0, 'qty': 4}]
        r = a._scrap_begin_partial(pads, 1.0, 'felt', 80, 80, None, ask_save_dir=True)
        assert r is not None
        assert len(askdir_calls) == 1, "file export should prompt once"
        assert a.scrap_session['save_dir'] == "/tmp/scrap_out", "save_dir not stored"
    check("File export prompts for folder on new session", t_file_mode_prompts)

    # ----------------------------------------------------------------
    # Cancelling the folder prompt on a new file-export session aborts
    # without starting a session.
    # ----------------------------------------------------------------
    def t_file_mode_cancel_folder():
        reset_session()
        clear_msgs()
        set_askdir("")  # askdirectory returns "" on cancel
        pads = [{'size': 16.0, 'qty': 4}]
        r = a._scrap_begin_partial(pads, 1.0, 'felt', 80, 80, None, ask_save_dir=True)
        assert r is None, "cancelled folder prompt must abort"
        assert a.scrap_session['active'] is False, "no session on cancel"
    check("Cancelling folder prompt aborts cleanly", t_file_mode_cancel_folder)

    # ----------------------------------------------------------------
    # Mixed workflow: a session started by Frame & Cut (no save_dir) that
    # is later continued by file export prompts for a folder once.
    # ----------------------------------------------------------------
    def t_frame_cut_then_file_export():
        reset_session()
        clear_msgs()
        set_askdir(None)
        pads = [{'size': 16.0, 'qty': 60}]
        # Start via Frame & Cut -> empty save_dir.
        a._scrap_begin_partial(pads, 1.0, 'felt', 55, 55, None, ask_save_dir=False)
        assert a.scrap_session['save_dir'] == ''
        # Continue via file export -> should prompt + store the folder.
        set_askdir("/tmp/late_folder")
        r = a._scrap_begin_partial(pads, 1.0, 'felt', 55, 55, None, ask_save_dir=True)
        assert r is not None
        assert len(askdir_calls) == 1, "should prompt once for the folder"
        assert a.scrap_session['save_dir'] == "/tmp/late_folder", \
            "late save_dir not stored on session"
    check("Frame&Cut session then file export prompts for folder", t_frame_cut_then_file_export)

    try:
        root.destroy()
    except Exception:
        pass

    passed = sum(results)
    total = len(results)
    print("=" * 60)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main_test())
