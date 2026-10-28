"""
Tests for the "On this scrap" state of the Scrap Mode Progress window
(main.py).

Until 2026-10-26 the window knew two states, Remaining and Done, and moved
only when a cut finished or a file was written. From the moment a nest was
previewed until the laser stopped it still read "0 / 60 pads complete"
with every pad in Remaining -- which looks exactly like pads being dropped
(Matt, in the shop, 2026-10-26, after a scrap that skipped three sizes).

The session now records what the current scrap's nest placed in
'pending'. The window shows it in a middle column and takes it off
Remaining; the main-pane status says "N on this scrap". A commit
(_scrap_record_result, reached from the SVG, G-code and Frame & Cut
paths) moves it to Done; a backed-out preview, a cancelled dialog, a
stopped cut or an error puts it back (_scrap_abandon_pending, called from
each path's finally).

Drives the bookkeeping through _scrap_begin_partial() and the commit /
abandon helpers with the REAL progress window (withdrawn), then the real
Generate button with the preview stubbed. Needs a display; headless Linux
self-skips.

Run:
    python tools/test_scrap_pending.py
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
    print("Scrap Mode 'On this scrap' Tests")
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

    continue_calls = []
    a._show_scrap_continue_dialog = lambda *args: continue_calls.append(args)
    main.messagebox.showinfo = lambda *a_, **k: None
    main.messagebox.showwarning = lambda *a_, **k: None
    main.messagebox.showerror = lambda *a_, **k: None
    main.messagebox.askyesno = lambda *a_, **k: False

    # Keep the real progress window (its three columns are what's under
    # test) but never let it show.
    orig_open = a._open_remaining_pads_window

    def open_hidden():
        orig_open()
        a.scrap_remaining_window.withdraw()
    a._open_remaining_pads_window = open_hidden
    a.scrap_mode_var.set(True)

    SMALL = 55    # a 3x3 of 16mm felt discs: 60 of them take several scraps

    def reset_session():
        a._clear_scrap_session()
        continue_calls.clear()

    def begin(pads, w, h):
        r = a._scrap_begin_partial(pads, 1.0, 'felt', w, h, None, ask_save_dir=False)
        assert r is not None, "expected a placement, got None"
        return r

    def col(lb):
        return list(lb.get(0, tk.END))

    def remaining_col():
        return col(a.scrap_remaining_listbox)

    def pending_col():
        return col(a.scrap_pending_listbox)

    def done_col():
        return col(a.scrap_done_listbox)

    def header():
        return a.scrap_window_header.cget("text")

    def status():
        return a.scrap_status_var.get()

    def file_commit(remaining):
        # What the SVG / G-code paths do after writing the file.
        a._scrap_record_result(remaining)
        a.scrap_session['scrap_count'] += 1
        a._update_scrap_status_display()
        a._update_remaining_pads_window()

    # ----------------------------------------------------------------
    def t_nest_moves_pads_to_the_middle_column():
        reset_session()
        placed, _rem = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n = len(placed)
        assert 0 < n < 60, f"small scrap should take a few, took {n}"
        assert a.scrap_session['pending'] == {16.0: n}, a.scrap_session['pending']
        assert a.scrap_session['scrap_count'] == 0, "nothing is committed yet"
        assert remaining_col() == [f" {60 - n} x 16"], remaining_col()
        assert pending_col() == [f" {n} x 16"], pending_col()
        assert done_col() == [" (none)"], done_col()
        assert header() == f"{n} on scrap #1, 0 / 60 cut", header()
        assert status() == f"{60 - n} left, {n} on this scrap (0 scraps)", status()
    check("A nest moves its pads from Remaining to 'On this scrap' before any cut",
          t_nest_moves_pads_to_the_middle_column)

    # ----------------------------------------------------------------
    def t_abandon_puts_them_back():
        reset_session()
        begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        a._scrap_abandon_pending()
        assert a.scrap_session['pending'] == {}
        assert remaining_col() == [" 60 x 16"], remaining_col()
        assert pending_col() == [" (none)"], pending_col()
        assert header() == "0 / 60 pads complete", header()
        assert status() == "60 left (0 scraps)", status()
        # Abandoning twice (every path's finally runs after a commit too)
        # changes nothing.
        a._scrap_abandon_pending()
        assert remaining_col() == [" 60 x 16"]
    check("Backing out (preview Adjust, stopped cut, error) puts them back in Remaining",
          t_abandon_puts_them_back)

    # ----------------------------------------------------------------
    def t_commit_moves_them_to_done():
        reset_session()
        placed, rem = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n = len(placed)
        file_commit(rem)
        assert a.scrap_session['pending'] == {}
        assert done_col() == [f" {n} x 16"], done_col()
        assert remaining_col() == [f" {60 - n} x 16"], remaining_col()
        assert pending_col() == [" (none)"], pending_col()
        assert header() == f"{n} / 60 pads complete", header()
        assert status() == f"{60 - n} left (1 scraps)", status()
        assert "1 scrap(s) used" in a.scrap_window_footer.cget("text")
        # The next nest is scrap #2 and Remaining keeps shrinking.
        placed2, _rem2 = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        assert header() == f"{len(placed2)} on scrap #2, {n} / 60 cut", header()
        assert remaining_col() == [f" {60 - n - len(placed2)} x 16"], remaining_col()
    check("A commit moves them to Done, and the next nest is 'on scrap #2'",
          t_commit_moves_them_to_done)

    # ----------------------------------------------------------------
    def t_mixed_sizes_only_the_placed_ones_move():
        # Matt's case: a big disc takes the wide part, the mid size skips
        # this scrap, small ones fill the strip left over. The skipped
        # size must stay in Remaining in full while the placed ones show
        # in the middle. Pinned to known sizing and bias settings: the
        # app loads the real config, and a felt 60 (59.25 mm) placed
        # flush-west on 74x74 leaves a 12.75 mm strip that takes 9.25 mm
        # discs and not 29.25 mm ones, whatever else is set.
        reset_session()
        saved = {k: a.settings.get(k) for k in ('sizing_range_mode', 'felt_offset', 'edge_bias')}
        a.settings.update(sizing_range_mode='universal', felt_offset=0.75, edge_bias='w')
        try:
            pads = [{'size': 60.0, 'qty': 1}, {'size': 30.0, 'qty': 4}, {'size': 10.0, 'qty': 6}]
            placed, _rem = begin(pads, 74, 74)
            got = a._count_by_size(placed)
            assert got == {60.0: 1, 10.0: 6}, got
            rem_lines = remaining_col()
            assert rem_lines == [" 4 x 30"], rem_lines
            assert pending_col() == [" 1 x 60", " 6 x 10"], pending_col()
            assert header() == "7 on scrap #1, 0 / 11 cut", header()
        finally:
            a.settings.update(saved)
    check("Mixed sizes: skipped sizes stay in Remaining, placed ones move",
          t_mixed_sizes_only_the_placed_ones_move)

    # ----------------------------------------------------------------
    def t_frame_cut_commit_clears_pending():
        reset_session()
        placed, rem = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        a._frame_cut_scrap_advance(len(placed), rem)
        assert a.scrap_session['pending'] == {}
        assert a.scrap_session['done'] == {16.0: len(placed)}
        assert len(continue_calls) == 1
        assert pending_col() == [" (none)"]
    check("Frame & Cut's completed-cut commit clears the pending column",
          t_frame_cut_commit_clears_pending)

    # ----------------------------------------------------------------
    def t_old_session_dict_tolerated():
        # A session built before 'done' / 'pending' existed.
        a.scrap_session = {
            'active': False, 'original_pads': [], 'remaining_pads': [],
            'scrap_count': 0, 'material': None, 'save_dir': '', 'hole_dia': 0,
        }
        a._scrap_abandon_pending()
        assert a._count_pending_pads() == 0
        placed, _rem = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        assert a.scrap_session['pending'] == {16.0: len(placed)}
    check("An old session dict without 'pending' still works",
          t_old_session_dict_tolerated)

    # ================================================================
    # Through the real Generate button with the preview up: Adjust must
    # leave the session untouched (and the window back at 'remaining'),
    # Save must commit.
    # ================================================================
    import shutil
    import tempfile
    tmp = tempfile.mkdtemp(prefix="ssc_pending_")
    a._record_job = lambda *args, **k: None
    main.save_settings = lambda s: None          # never touch the real config
    main.filedialog.askdirectory = lambda **k: tmp
    a.settings['units'] = 'mm'
    a.settings['show_engraving_warning'] = False
    a.custom_polygon = None
    a.preview_var.set(True)

    preview_results = []
    seen_during_preview = []

    class FakePreview:
        """Stands in for NestingPreviewWindow: records what the progress
        window said while the preview was 'open', then answers."""
        def __init__(self, *args, **kwargs):
            seen_during_preview.append((header(), remaining_col(), pending_col(), status()))
            self.result = preview_results.pop(0)
    main.NestingPreviewWindow = FakePreview

    def set_box(text):
        a.pad_entry.delete("1.0", tk.END)
        a.pad_entry.insert("1.0", text)

    def set_sheet(w, h):
        a.width_entry.delete(0, tk.END)
        a.width_entry.insert(0, str(w))
        a.height_entry.delete(0, tk.END)
        a.height_entry.insert(0, str(h))

    def select_only(material):
        for m, v in a.material_vars.items():
            v.set(m == material)

    def t_generate_adjust_then_save():
        reset_session()
        seen_during_preview.clear()
        for f in os.listdir(tmp):
            os.remove(os.path.join(tmp, f))
        select_only('felt')
        a.filename_entry.delete(0, tk.END)
        a.filename_entry.insert(0, "pend")
        set_box("16.0 x 60")
        set_sheet(SMALL, SMALL)

        preview_results.append("cancel")       # Adjust
        a.on_generate_gcode()
        assert len(seen_during_preview) == 1
        hdr, rem, pend, st = seen_during_preview[0]
        n = int(hdr.split(" on scrap")[0])
        assert 0 < n < 60, hdr
        assert hdr == f"{n} on scrap #1, 0 / 60 cut", hdr
        assert rem == [f" {60 - n} x 16"] and pend == [f" {n} x 16"], (rem, pend)
        assert "on this scrap" in st, st
        # Adjust: nothing consumed, window back to the pre-nest state.
        assert a.scrap_session['active']
        assert a.scrap_session['scrap_count'] == 0
        assert a.scrap_session['pending'] == {}
        assert remaining_col() == [" 60 x 16"], remaining_col()
        assert pending_col() == [" (none)"], pending_col()
        assert os.listdir(tmp) == [], os.listdir(tmp)

        preview_results.append("save")
        a.on_generate_gcode()
        assert a.scrap_session['scrap_count'] == 1
        assert a.scrap_session['pending'] == {}
        assert a.scrap_session['done'] == {16.0: n}, a.scrap_session['done']
        assert done_col() == [f" {n} x 16"], done_col()
        assert pending_col() == [" (none)"], pending_col()
        assert os.listdir(tmp) == ["pend_felt_scrap1.gcode"], os.listdir(tmp)
    check("Generate: Adjust on the preview puts the pads back; Save commits them",
          t_generate_adjust_then_save)

    shutil.rmtree(tmp, ignore_errors=True)

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
