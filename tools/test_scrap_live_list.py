"""
Tests for the live pad list in scrap mode (main.py).

A scrap session used to copy the pad list when it started and ignore the
box until the session ended. A user who added sizes mid-session to fill a
half-empty last scrap got a scrap without them and no message (reported
2026-09-07). Now the box is live: the session keeps only a per-size tally
of what is already cut ('done'), and each scrap nests "what the box says
now" minus that tally.

Drives the session bookkeeping directly through _scrap_begin_partial()
(the shared front half of the file-export and Frame & Cut paths) and the
two commit helpers, with dialogs stubbed. Constructs the full app in a
withdrawn Tk root, so it needs a display (Windows / macOS / CI Windows);
headless Linux self-skips.

Run:
    python tools/test_scrap_live_list.py
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
    print("Scrap Mode Live Pad List Tests")
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
    info_msgs = []
    a._show_scrap_continue_dialog = lambda *args: continue_calls.append(args)
    a._open_remaining_pads_window = lambda: None
    a._update_remaining_pads_window = lambda: None
    a._update_scrap_status_display = lambda: None
    main.messagebox.showinfo = lambda *a_, **k: info_msgs.append((a_, k))
    main.messagebox.showwarning = lambda *a_, **k: None
    main.messagebox.showerror = lambda *a_, **k: None
    main.messagebox.askyesno = lambda *a_, **k: False

    def reset_session():
        # Deliberately the OLD dict shape (no 'done'): the helpers must
        # tolerate a session built before the tally existed.
        a.scrap_session = {
            'active': False, 'original_pads': [], 'remaining_pads': [],
            'scrap_count': 0, 'material': None, 'save_dir': '',
            'hole_dia': 0,
        }
        continue_calls.clear()
        info_msgs.clear()

    BIG = 200    # everything in these tests fits on one 200x200 scrap
    SMALL = 55   # a 3x3 of 16mm felt discs, so 60 of them take several scraps

    def begin(pads, w, h):
        r = a._scrap_begin_partial(pads, 1.0, 'felt', w, h, None, ask_save_dir=False)
        assert r is not None, "expected a placement, got None"
        return r

    def by_size(placed):
        out = {}
        for sz, _x, _y, _r in placed:
            out[sz] = out.get(sz, 0) + 1
        return out

    def file_commit(remaining):
        # What the SVG / G-code paths do after writing the file.
        a._scrap_record_result(remaining)
        a.scrap_session['scrap_count'] += 1

    # ----------------------------------------------------------------
    def t_unchanged_list_still_decrements():
        reset_session()
        pads = [{'size': 16.0, 'qty': 60}]
        placed1, rem1 = begin(pads, SMALL, SMALL)
        n1 = len(placed1)
        assert 0 < n1 < 60, f"small scrap should take a few, took {n1}"
        file_commit(rem1)
        assert a.scrap_session['done'] == {16.0: n1}, a.scrap_session['done']
        placed2, rem2 = begin(pads, SMALL, SMALL)
        assert len(placed2) == n1, "same scrap size should fit the same count"
        assert sum(p['qty'] for p in rem2) == 60 - 2 * n1
    check("Unchanged list: remaining decrements scrap by scrap as before",
          t_unchanged_list_still_decrements)

    # ----------------------------------------------------------------
    def t_added_size_is_cut_next_scrap():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n1 = len(placed1)
        file_commit(rem1)
        # The user adds a line to the box before the next scrap.
        placed2, rem2 = begin([{'size': 16.0, 'qty': 60}, {'size': 11.0, 'qty': 4}], BIG, BIG)
        got = by_size(placed2)
        assert got.get(11.0) == 4, f"added size not cut: {got}"
        assert got.get(16.0) == 60 - n1, f"done credit lost: {got}"
        assert rem2 == [], rem2
    check("A size added mid-session is cut on the next scrap",
          t_added_size_is_cut_next_scrap)

    # ----------------------------------------------------------------
    def t_raised_qty_keeps_done_credit():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n1 = len(placed1)
        file_commit(rem1)
        placed2, rem2 = begin([{'size': 16.0, 'qty': 70}], BIG, BIG)
        assert len(placed2) == 70 - n1, f"expected {70 - n1}, got {len(placed2)}"
        assert rem2 == []
    check("Raising a quantity keeps the credit for pads already cut",
          t_raised_qty_keeps_done_credit)

    # ----------------------------------------------------------------
    def t_lowered_below_done_clamps_to_zero():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}, {'size': 20.0, 'qty': 3}], BIG, BIG)
        assert rem1 == [], "everything should fit on the big scrap"
        file_commit(rem1)
        assert a.scrap_session['done'] == {16.0: 60, 20.0: 3}
        # 60 are cut; the user lowers the line to 40.
        nxt = a._scrap_pads_for_this_scrap([{'size': 16.0, 'qty': 40}, {'size': 20.0, 'qty': 3}])
        assert nxt == [], f"nothing should remain, got {nxt}"
        assert a.scrap_session['remaining_pads'] == []
        assert a.scrap_session['done'] == {16.0: 60, 20.0: 3}, "tally must not shrink"
        out = a._scrap_begin_partial([{'size': 16.0, 'qty': 40}], 1.0, 'felt',
                                     BIG, BIG, None, ask_save_dir=False)
        assert out is None and len(info_msgs) == 1, "should report the session complete"
    check("Lowering below the cut count clamps to zero, never negative",
          t_lowered_below_done_clamps_to_zero)

    # ----------------------------------------------------------------
    def t_removed_size_drops_out():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}, {'size': 20.0, 'qty': 30}], SMALL, SMALL)
        cut_20 = by_size(placed1).get(20.0, 0)
        file_commit(rem1)
        placed2, rem2 = begin([{'size': 16.0, 'qty': 60}], BIG, BIG)   # 20.0 line deleted
        assert 20.0 not in by_size(placed2), "deleted size must not be cut"
        assert all(p['size'] != 20.0 for p in rem2)
        assert a.scrap_session['done'].get(20.0, 0) == cut_20, "tally keeps what was cut"
    check("A size removed from the box is no longer cut, but stays in the tally",
          t_removed_size_drops_out)

    # ----------------------------------------------------------------
    def t_duplicate_lines_merge():
        reset_session()
        nxt = a._scrap_pads_for_this_scrap([{'size': 16.0, 'qty': 2}, {'size': 16.0, 'qty': 3}])
        assert nxt == [{'size': 16.0, 'qty': 5}], nxt
        assert a.scrap_session['original_pads'] == [{'size': 16.0, 'qty': 5}]
    check("Duplicate lines for one size merge", t_duplicate_lines_merge)

    # ----------------------------------------------------------------
    def t_status_count_follows_the_box():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n1 = len(placed1)
        file_commit(rem1)
        a._scrap_pads_for_this_scrap([{'size': 16.0, 'qty': 60}, {'size': 11.0, 'qty': 4}])
        assert a._count_remaining_pads() == 60 - n1 + 4
    check("The 'N left' count follows the box", t_status_count_follows_the_box)

    # ----------------------------------------------------------------
    def t_frame_cut_advance_credits_done():
        reset_session()
        placed1, rem1 = begin([{'size': 16.0, 'qty': 60}], SMALL, SMALL)
        n1 = len(placed1)
        a._frame_cut_scrap_advance(n1, rem1)
        assert a.scrap_session['scrap_count'] == 1
        assert a.scrap_session['done'] == {16.0: n1}, a.scrap_session['done']
        assert len(continue_calls) == 1, "continue dialog should show"
        placed2, rem2 = begin([{'size': 16.0, 'qty': 60}, {'size': 11.0, 'qty': 2}], BIG, BIG)
        assert by_size(placed2) == {16.0: 60 - n1, 11.0: 2}, by_size(placed2)
    check("Frame & Cut commit credits the same tally as file export",
          t_frame_cut_advance_credits_done)

    # ----------------------------------------------------------------
    def t_max_first_scrap_only_unchanged():
        # Pre-existing behaviour, pinned so the live list doesn't move it:
        # a 'max' line fills the FIRST scrap only and never enters the tally.
        reset_session()
        pads = [{'size': 16.0, 'qty': 2}, {'size': 11.0, 'qty': 'max'}]
        placed1, rem1 = begin(pads, BIG, BIG)
        got1 = by_size(placed1)
        assert got1.get(16.0) == 2 and got1.get(11.0, 0) > 10, got1
        assert rem1 == []
        file_commit(rem1)
        assert a.scrap_session['done'] == {16.0: 2}, a.scrap_session['done']
        assert a._scrap_pads_for_this_scrap(pads) == [], "max is not carried to scrap 2"
    check("'max' still fills the first scrap only (pre-existing, unchanged)",
          t_max_first_scrap_only_unchanged)

    # ================================================================
    # End to end through the real Generate button path: type in the pad
    # box, Generate with scrap mode on, edit the box, Generate again.
    # Only the file/folder dialogs and message boxes are stubbed.
    # ================================================================
    import shutil
    import tempfile
    tmp = tempfile.mkdtemp(prefix="ssc_scrap_")
    jobs = []
    a._record_job = lambda output, materials, pads, params, **k: jobs.append(
        (output, k.get('placed_count'), k.get('scrap_num')))
    main.save_settings = lambda s: None          # never touch the real config
    main.filedialog.askdirectory = lambda **k: tmp
    a.settings['units'] = 'mm'
    a.settings['show_engraving_warning'] = False
    a.preview_var.set(False)
    a.custom_polygon = None
    a.scrap_mode_var.set(True)

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

    def e2e(generate, ext):
        reset_session()
        jobs.clear()
        for f in os.listdir(tmp):
            os.remove(os.path.join(tmp, f))
        select_only('felt')
        a.filename_entry.delete(0, tk.END)
        a.filename_entry.insert(0, "live")
        set_box("16.0 x 60")
        set_sheet(SMALL, SMALL)
        generate()
        assert a.scrap_session['active'], "first Generate should start the session"
        assert a.scrap_session['scrap_count'] == 1, jobs
        n1 = jobs[-1][1]
        assert 0 < n1 < 60, f"scrap 1 placed {n1}"
        # The user edits the box between scraps: a new size, and a bigger piece.
        set_box("16.0 x 60\n11.0 x 4")
        set_sheet(BIG, BIG)
        generate()
        assert a.scrap_session['scrap_count'] == 2, jobs
        assert jobs[-1][1] == 60 - n1 + 4, \
            f"scrap 2 placed {jobs[-1][1]}, expected {60 - n1 + 4}"
        assert a.scrap_session['remaining_pads'] == []
        assert a.scrap_session['done'] == {16.0: 60, 11.0: 4}, a.scrap_session['done']
        names = sorted(os.listdir(tmp))
        assert names == [f"live_felt_scrap1{ext}", f"live_felt_scrap2{ext}"], names
        assert all(os.path.getsize(os.path.join(tmp, n)) > 0 for n in names)

    check("End to end: Generate G-code, edit the box, Generate again",
          lambda: e2e(a.on_generate_gcode, ".gcode"))
    check("End to end: Generate SVG, edit the box, Generate again",
          lambda: e2e(a.on_generate_svg, ".svg"))
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
