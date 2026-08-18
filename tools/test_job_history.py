"""
Job History tests — storage layer, list/detail formatting, and the
record-then-reload round trip through the Pad Maker form.

Run:
    python tools/test_job_history.py

The storage tests redirect config.JOB_HISTORY_FILE at a temp file, so
they never touch the real job_history.json. The form round-trip needs a
display (fine on Windows / macOS; self-skips on headless Linux).
"""
import json
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from i18n import init_translation  # noqa: E402
init_translation("en")

import config  # noqa: E402

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


def use_temp_history():
    """Point the history file at a fresh temp path. Returns the path."""
    fd, path = tempfile.mkstemp(suffix="_job_history.json")
    os.close(fd)
    os.remove(path)
    config.JOB_HISTORY_FILE = path
    return path


def make_job(**over):
    job = {
        "timestamp": "2026-08-01T14:22:05",
        "output": "gcode",
        "status": "complete",
        "materials": ["felt"],
        "pad_text": "18.0 x 5\n22.5 x 3",
        "pads": [{"size": 18.0, "qty": 5}, {"size": 22.5, "qty": 3}],
        "requested_count": 8,
        "has_max": False,
        "placed_count": 8,
        "base": "testjob",
        "units": "in",
        "sheet_w": "12", "sheet_h": "12",
        "sheet_w_mm": 304.8, "sheet_h_mm": 304.8,
        "card_paper": False,
        "hole_option": "3.0mm",
        "custom_hole": "4.0",
        "scrap_num": None,
        "save_dir": r"C:\out",
        "preset_library": None,
        "preset_name": None,
        "polygon_vertices": 0,
    }
    job.update(over)
    return job


# ==========================================
# Storage layer
# ==========================================

def test_empty_history():
    use_temp_history()
    assert config.load_job_history() == [], "missing file should load as []"


def test_append_round_trip():
    use_temp_history()
    config.append_job_history(make_job(base="first"))
    config.append_job_history(make_job(base="second"))
    jobs = config.load_job_history()
    assert len(jobs) == 2, f"expected 2 jobs, got {len(jobs)}"
    # Newest first — the list is a log read top-down in the UI.
    assert jobs[0]["base"] == "second", f"newest not first: {jobs[0]['base']}"
    assert jobs[1]["base"] == "first"
    assert jobs[0]["pads"][0]["size"] == 18.0, "pad data did not survive JSON"


def test_limit_trim():
    path = use_temp_history()
    for i in range(config.JOB_HISTORY_LIMIT + 25):
        config.append_job_history(make_job(base=f"job{i}"))
    jobs = config.load_job_history()
    assert len(jobs) == config.JOB_HISTORY_LIMIT, \
        f"expected trim to {config.JOB_HISTORY_LIMIT}, got {len(jobs)}"
    # The oldest entries are the ones dropped.
    newest = config.JOB_HISTORY_LIMIT + 24
    assert jobs[0]["base"] == f"job{newest}", f"wrong newest: {jobs[0]['base']}"
    with open(path) as f:
        on_disk = json.load(f)
    assert on_disk["version"] == 1, "version stamp missing"
    assert len(on_disk["jobs"]) == config.JOB_HISTORY_LIMIT


def test_bare_list_tolerated():
    """A hand-edited file holding a bare list still loads."""
    path = use_temp_history()
    with open(path, "w") as f:
        json.dump([make_job(base="bare")], f)
    jobs = config.load_job_history()
    assert len(jobs) == 1 and jobs[0]["base"] == "bare", jobs


def test_corrupt_file_preserved():
    path = use_temp_history()
    with open(path, "w") as f:
        f.write("{ this is not json")
    assert config.load_job_history() == [], "corrupt file should load as []"
    assert os.path.exists(path + ".corrupt.bak"), "no .corrupt.bak written"
    # And a save afterwards still works (defaults took over cleanly).
    assert config.save_job_history([make_job()]), "save after corruption failed"
    assert len(config.load_job_history()) == 1


def test_junk_entries_filtered():
    path = use_temp_history()
    with open(path, "w") as f:
        json.dump({"version": 1, "jobs": [make_job(), "not a dict", 42, None]}, f)
    jobs = config.load_job_history()
    assert len(jobs) == 1, f"non-dict entries not filtered: {jobs}"


def test_save_never_raises():
    """A bad path logs and returns False instead of blowing up a laser run."""
    config.JOB_HISTORY_FILE = os.path.join(
        tempfile.gettempdir(), "no_such_dir_xyz", "job_history.json")
    assert config.save_job_history([make_job()]) is False, \
        "save to an unwritable path should return False"


# ==========================================
# Dialog formatting (no Tk needed)
# ==========================================

def _fmt():
    """A JobHistoryWindow with no widgets — formatting helpers only."""
    from ui_dialogs import JobHistoryWindow
    return JobHistoryWindow.__new__(JobHistoryWindow)


def test_when_formatting():
    w = _fmt()
    assert w._when(make_job()) == "2026-08-01 14:22", w._when(make_job())
    # Garbage timestamps must not crash the list.
    assert isinstance(w._when(make_job(timestamp="nonsense")), str)
    assert isinstance(w._when(make_job(timestamp=None)), str)


def test_output_labels():
    w = _fmt()
    assert w._output_label(make_job(output="svg")) == "SVG"
    assert w._output_label(make_job(output="gcode")) == "G-code"
    assert w._output_label(make_job(output="laser")) == "Laser"
    # A cut that ended early is flagged so missing pads have an explanation.
    stopped = w._output_label(make_job(output="laser", status="cancelled"))
    assert stopped != "Laser" and "Laser" in stopped, stopped


def test_pad_count_prefers_placed():
    w = _fmt()
    # 'max' fills mean requested != placed; placed is the truth.
    assert w._pad_count(make_job(placed_count=31, requested_count=8)) == 31
    # Older entries without placed_count fall back to requested.
    assert w._pad_count(make_job(placed_count=None, requested_count=8)) == 8
    assert w._pad_count(make_job(placed_count=None, requested_count=None)) == 0


def test_row_and_detail_text():
    w = _fmt()
    row = w._row_text(make_job())
    # Material names render through the same display transform as the
    # Materials checkboxes ("felt" -> "Felt", translated when non-English).
    for expect in ("2026-08-01", "G-code", "Felt", "8 pads", "12 x 12 in"):
        assert expect in row, f"{expect!r} missing from row: {row!r}"

    detail = w._detail_text(make_job(
        output="laser", status="complete", scrap_num=2,
        materials=["leather"], polygon_vertices=6, has_max=True,
        preset_name="Tenor Full Set", preset_library="My Presets"))
    for expect in ("Sent to the laser", "Scrap piece #2", "Leather",
                   "max", "Tenor Full Set", "6 points", "18.0 x 5"):
        assert expect in detail, f"{expect!r} missing from detail:\n{detail}"

    # Custom hole is spelled out rather than shown as the literal "Custom".
    custom = w._detail_text(make_job(hole_option="Custom", custom_hole="4.5"))
    assert "4.5mm" in custom, custom


def test_columns_stay_aligned():
    """No value may overflow its column and shove the next one sideways."""
    w = _fmt()
    jobs = [
        make_job(),
        # scrap marker pushes the Pads column widest
        make_job(scrap_num=12, placed_count=311),
        # 3+ materials collapse to a count rather than a chopped join
        make_job(materials=["felt", "card", "leather"]),
        make_job(output="laser", status="cancelled", materials=["exact_size"]),
        # a long sheet label must not shift anything before it
        make_job(sheet_w="1234.5", sheet_h="9876.5", units="mm"),
    ]
    widths = w._column_widths(jobs)
    starts = [w._row_text(j, widths).index(w._sheet_label(j)) for j in jobs]
    assert len(set(starts)) == 1, f"Sheet column misaligned across rows: {starts}"

    # The header lines up with the rows it labels.
    header = w._join(w._headers(), widths)
    assert header.index("Sheet") == starts[0], \
        f"header Sheet at {header.index('Sheet')}, rows at {starts[0]}"

    # Columns must widen for long content instead of truncating it.
    long_mats = w._row_text(make_job(materials=["felt", "card"]))
    assert "Felt+Card" in long_mats, long_mats

    # The collapse must be a whole phrase, not a chopped word.
    three = w._materials_short(make_job(materials=["felt", "card", "leather"]))
    assert three == "3 materials", three
    assert w._materials_short(make_job(materials=["felt"])) == "Felt"


def test_columns_survive_long_translations():
    """Simulated long translated headers must still align (es 'Zapatillas')."""
    w = _fmt()
    jobs = [make_job(), make_job(scrap_num=3)]
    real_headers = w._headers

    def long_headers():
        return ["Cuándo", "Salida", "Material", "Zapatillas", "Hoja"]

    try:
        w._headers = long_headers
        widths = w._column_widths(jobs)
        header = w._join(long_headers(), widths)
        assert "Zapatillas" in header, header
        starts = [w._row_text(j, widths).index(w._sheet_label(j)) for j in jobs]
        assert len(set(starts)) == 1, starts
        assert header.index("Hoja") == starts[0], (header, starts)
    finally:
        w._headers = real_headers


def test_row_survives_sparse_entry():
    """A half-written entry must still render — the list is a log, not a form."""
    w = _fmt()
    row = w._row_text({"output": "svg"})
    assert isinstance(row, str) and "SVG" in row, row
    assert isinstance(w._detail_text({"output": "svg"}), str)


# ==========================================
# Record + reload round trip (needs a display)
# ==========================================

def test_form_round_trip(app, tk):
    """Record a job off the form, wipe the form, load the job back."""
    use_temp_history()

    app.pad_entry.delete("1.0", tk.END)
    app.pad_entry.insert(tk.END, "18.0 x 5\n22.5 x 3")
    app.width_entry.delete(0, tk.END)
    app.width_entry.insert(0, "14")
    app.height_entry.delete(0, tk.END)
    app.height_entry.insert(0, "11")
    app.filename_entry.delete(0, tk.END)
    app.filename_entry.insert(0, "roundtrip")
    app.hole_var.set("3.5mm")
    for m, var in app.material_vars.items():
        var.set(m == "leather")
    app.settings['units'] = 'in'

    pads = app.parse_pad_list(app.pad_entry.get("1.0", tk.END))
    params = {'base': 'roundtrip', 'width_mm': 14 * 25.4,
              'height_mm': 11 * 25.4, 'card_paper_dims': None}
    app._record_job("gcode", ["leather"], pads, params,
                    placed_count=8, save_dir=r"C:\out")

    jobs = config.load_job_history()
    assert len(jobs) == 1, f"job not recorded: {jobs}"
    job = jobs[0]
    assert job["output"] == "gcode" and job["materials"] == ["leather"]
    assert job["placed_count"] == 8 and job["requested_count"] == 8
    assert job["hole_option"] == "3.5mm"
    assert job["base"] == "roundtrip"
    assert abs(job["sheet_w_mm"] - 355.6) < 0.01, job["sheet_w_mm"]

    # Wipe the form, then restore from history.
    app.pad_entry.delete("1.0", tk.END)
    app.width_entry.delete(0, tk.END)
    app.height_entry.delete(0, tk.END)
    app.filename_entry.delete(0, tk.END)
    app.hole_var.set("3.0mm")
    for var in app.material_vars.values():
        var.set(False)

    app._load_job_into_form(job)

    assert app.pad_entry.get("1.0", tk.END).strip() == "18.0 x 5\n22.5 x 3", \
        repr(app.pad_entry.get("1.0", tk.END))
    assert app.width_entry.get() == "14", app.width_entry.get()
    assert app.height_entry.get() == "11", app.height_entry.get()
    assert app.filename_entry.get() == "roundtrip"
    assert app.hole_var.get() == "3.5mm"
    assert app.material_vars['leather'].get() is True
    assert app.material_vars['felt'].get() is False
    # Loading a job is not loading a preset.
    assert app.pad_preset_loaded_name is None


def test_unit_conversion_on_load(app, tk):
    """A job saved in inches loads correctly when the app is now in mm."""
    job = make_job(units="in", sheet_w="12", sheet_h="12",
                   sheet_w_mm=304.8, sheet_h_mm=304.8)
    app.settings['units'] = 'mm'
    app._load_job_into_form(job)
    assert app.width_entry.get() == "304.8", app.width_entry.get()
    assert app.height_entry.get() == "304.8", app.height_entry.get()

    app.settings['units'] = 'cm'
    app._load_job_into_form(job)
    assert app.width_entry.get() == "30.48", app.width_entry.get()
    app.settings['units'] = 'in'


def test_scrap_session_blocks_load(app, tk):
    """A live scrap session owns the pad list — loading over it is refused."""
    import main as main_mod
    app.pad_entry.delete("1.0", tk.END)
    app.pad_entry.insert(tk.END, "KEEP ME")
    app.scrap_session['active'] = True

    warned = []
    orig = main_mod.messagebox.showwarning
    main_mod.messagebox.showwarning = lambda *a, **k: warned.append(a)
    try:
        app._load_job_into_form(make_job())
    finally:
        main_mod.messagebox.showwarning = orig
        app.scrap_session['active'] = False

    assert warned, "no warning shown for load during scrap session"
    assert app.pad_entry.get("1.0", tk.END).strip() == "KEEP ME", \
        "form was modified despite active scrap session"


def _run_generation(app, tk, method_name, out_dir):
    """Drive a real generate path with the file dialog and popups stubbed."""
    import main as main_mod

    def fake_askdirectory(**kwargs):
        return out_dir

    stubs = {
        'askdirectory': main_mod.filedialog.askdirectory,
        'showinfo': main_mod.messagebox.showinfo,
        'showerror': main_mod.messagebox.showerror,
        'showwarning': main_mod.messagebox.showwarning,
    }
    errors = []
    main_mod.filedialog.askdirectory = fake_askdirectory
    main_mod.messagebox.showinfo = lambda *a, **k: None
    main_mod.messagebox.showerror = lambda *a, **k: errors.append(a)
    main_mod.messagebox.showwarning = lambda *a, **k: errors.append(a)
    try:
        getattr(app, method_name)()
    finally:
        main_mod.filedialog.askdirectory = stubs['askdirectory']
        main_mod.messagebox.showinfo = stubs['showinfo']
        main_mod.messagebox.showerror = stubs['showerror']
        main_mod.messagebox.showwarning = stubs['showwarning']
    assert not errors, f"{method_name} reported an error: {errors}"


def _setup_simple_job(app, tk):
    app.scrap_mode_var.set(False)
    app.preview_var.set(False)
    app.custom_polygon = None
    app.pad_entry.delete("1.0", tk.END)
    app.pad_entry.insert(tk.END, "18.0 x 4")
    app.width_entry.delete(0, tk.END)
    app.width_entry.insert(0, "12")
    app.height_entry.delete(0, tk.END)
    app.height_entry.insert(0, "12")
    app.filename_entry.delete(0, tk.END)
    app.filename_entry.insert(0, "e2e")
    app.hole_var.set("3.0mm")
    app.settings['units'] = 'in'
    for m, var in app.material_vars.items():
        var.set(m == "felt")


def test_end_to_end_svg_generation(app, tk):
    """on_generate_svg must actually land an entry in the history."""
    use_temp_history()
    _setup_simple_job(app, tk)
    with tempfile.TemporaryDirectory() as out_dir:
        _run_generation(app, tk, "on_generate_svg", out_dir)
        written = os.listdir(out_dir)
        assert any(f.endswith(".svg") for f in written), f"no SVG written: {written}"

    jobs = config.load_job_history()
    assert len(jobs) == 1, f"expected 1 history entry, got {len(jobs)}"
    job = jobs[0]
    assert job["output"] == "svg", job["output"]
    assert job["materials"] == ["felt"], job["materials"]
    assert job["placed_count"] == 4, job["placed_count"]
    assert job["base"] == "e2e"
    assert job["pad_text"] == "18.0 x 4"
    assert job["status"] == "complete"


def test_end_to_end_gcode_generation(app, tk):
    """on_generate_gcode must log too, and not double-log the SVG run."""
    use_temp_history()
    _setup_simple_job(app, tk)
    with tempfile.TemporaryDirectory() as out_dir:
        _run_generation(app, tk, "on_generate_gcode", out_dir)
        written = os.listdir(out_dir)
        assert any(f.endswith(".gcode") for f in written), f"no G-code: {written}"

    jobs = config.load_job_history()
    assert len(jobs) == 1, f"expected 1 history entry, got {len(jobs)}"
    assert jobs[0]["output"] == "gcode", jobs[0]["output"]
    assert jobs[0]["placed_count"] == 4, jobs[0]["placed_count"]


def test_end_to_end_cancel_logs_nothing(app, tk):
    """Backing out of the save-folder dialog must not record a job."""
    use_temp_history()
    _setup_simple_job(app, tk)
    import main as main_mod
    orig = main_mod.filedialog.askdirectory
    main_mod.filedialog.askdirectory = lambda **k: ""   # user hit Cancel
    try:
        app.on_generate_svg()
    finally:
        main_mod.filedialog.askdirectory = orig
    assert config.load_job_history() == [], "cancelled generation was logged"


def test_record_never_raises(app, tk):
    """History failures must not propagate into a generation path."""
    use_temp_history()
    # params missing every key it normally has.
    app._record_job("svg", ["felt"], [], {})
    # An unwritable path: the write fails, the caller does not see it.
    config.JOB_HISTORY_FILE = os.path.join(
        tempfile.gettempdir(), "no_such_dir_xyz", "job_history.json")
    app._record_job("svg", ["felt"], [{'size': 18.0, 'qty': 2}],
                    {'base': 'x', 'width_mm': 300, 'height_mm': 300})


def main():
    print("Job History Tests")
    print("=" * 60)

    real_history_file = config.JOB_HISTORY_FILE

    print("\nStorage layer:")
    for fn in (test_empty_history, test_append_round_trip, test_limit_trim,
               test_bare_list_tolerated, test_corrupt_file_preserved,
               test_junk_entries_filtered, test_save_never_raises):
        check(fn.__name__, fn)

    print("\nList / detail formatting:")
    for fn in (test_when_formatting, test_output_labels,
               test_pad_count_prefers_placed, test_row_and_detail_text,
               test_columns_stay_aligned, test_columns_survive_long_translations,
               test_row_survives_sparse_entry):
        check(fn.__name__, fn)

    print("\nForm round trip:")
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  SKIP  no display available ({e})")
        root = None

    if root is not None:
        root.withdraw()
        try:
            from main import PadSVGGeneratorApp
            app = PadSVGGeneratorApp(root)
            for fn in (test_form_round_trip, test_unit_conversion_on_load,
                       test_scrap_session_blocks_load, test_record_never_raises,
                       test_end_to_end_svg_generation,
                       test_end_to_end_gcode_generation,
                       test_end_to_end_cancel_logs_nothing):
                check(fn.__name__, lambda f=fn: f(app, tk))
        finally:
            config.JOB_HISTORY_FILE = real_history_file
            try:
                root.destroy()
            except Exception:
                pass

    config.JOB_HISTORY_FILE = real_history_file

    passed = sum(results)
    total = len(results)
    print("=" * 60)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
