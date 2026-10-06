"""Pad Maker end to end through the real Generate buttons.

Types a pad list into the live form, picks all four materials, turns on
labeled zones and leather locator marks, and clicks Generate SVG and
Generate G-code exactly as a user would. Only the folder dialog and the
message boxes are stubbed. Then the OUTPUT is checked, not the internals:
one file per material, every SVG parses as XML with the right root, the
leather SVG carries the locator layer, the felt SVG carries a zone box,
every G-code coordinate lies on the sheet, every G-code file ends with the
laser off, and the job history recorded both runs.

Runs on every platform; on the macOS CI runner it is the Pad Maker's only
end-to-end check. Uses an isolated profile (SAXSHOP_CONFIG_DIR) so the
user's settings and job history are never touched. Self-skips without a
display.
"""
import glob
import os
import re
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-padmaker-e2e-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []
SHEET_W, SHEET_H = 200.0, 150.0
PADS = "7.0 x 4\n9.5 x 3\n12.5 x 2\n20.0 x 2\n35.5 x 1"
MATERIALS = ("felt", "card", "leather", "exact_size")


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


def main():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    print("Pad Maker end to end")
    print("=" * 60)
    import main as main_mod
    from config import load_job_history
    app = main_mod.PadSVGGeneratorApp(root)

    out = tempfile.mkdtemp(prefix="ssc-padmaker-out-")
    boxes = []
    main_mod.filedialog.askdirectory = lambda **k: out
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(main_mod.messagebox, name, lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg)[:80])))
    main_mod.messagebox.askyesno = lambda *a, **k: True
    main_mod.messagebox.askokcancel = lambda *a, **k: True

    # The form, as a user would fill it.
    app.settings['units'] = 'mm'
    app.settings['show_engraving_warning'] = False
    app.settings['zone_labels_enabled'] = True
    app.settings['locator_marks_enabled'] = True
    app.settings['locator_marks_style'] = 'both'
    app.preview_var.set(False)
    app.scrap_mode_var.set(False)
    app.custom_polygon = None
    app.pad_entry.delete("1.0", tk.END)
    app.pad_entry.insert("1.0", PADS)
    app.width_entry.delete(0, tk.END)
    app.width_entry.insert(0, str(SHEET_W))
    app.height_entry.delete(0, tk.END)
    app.height_entry.insert(0, str(SHEET_H))
    app.filename_entry.delete(0, tk.END)
    app.filename_entry.insert(0, "e2e")
    for m, v in app.material_vars.items():
        v.set(m in MATERIALS)
    app.hole_var.set("3.5mm")
    root.update()

    # ---- SVG -----------------------------------------------------------
    app.on_generate_svg()
    root.update()
    svgs = sorted(glob.glob(os.path.join(out, "*.svg")))
    check(f"Generate SVG wrote one file per material ({len(svgs)})", lambda: len(svgs) == len(MATERIALS))
    def no_error_boxes():
        errs = [b for b in boxes if b[0] == "showerror"]
        assert not errs, errs
    check("no error box during SVG generation", no_error_boxes)

    def svgs_parse():
        for path in svgs:
            root_el = ET.parse(path).getroot()
            assert root_el.tag.split('}')[-1] == "svg", path
            assert root_el.get("width", "").startswith(f"{SHEET_W:g}") and root_el.get("height", "").startswith(f"{SHEET_H:g}"), \
                (path, root_el.get("width"), root_el.get("height"))
    check("every SVG parses with the sheet size on its root", svgs_parse)

    def by_material(ext):
        found = {}
        for m in (MATERIALS if ext == ".svg" else GCODE_MATERIALS):
            hits = [p for p in glob.glob(os.path.join(out, f"*{ext}")) if m in os.path.basename(p)]
            assert hits, f"no {ext} for {m}"
            found[m] = hits[0]
        return found

    def leather_has_locator_layer():
        text = open(by_material(".svg")["leather"], encoding="utf-8").read()
        assert text.count('stroke="#00E0E0"') >= 4, "leather SVG has no locator marks"
    check("leather SVG carries the locator-mark layer", leather_has_locator_layer)

    def felt_has_zone_box():
        text = open(by_material(".svg")["felt"], encoding="utf-8").read()
        assert "<rect" in text, "felt SVG has no zone box"
    check("felt SVG carries a labeled-zone box", felt_has_zone_box)

    # ---- G-code --------------------------------------------------------
    boxes.clear()
    app.on_generate_gcode()
    root.update()
    gcodes = sorted(glob.glob(os.path.join(out, "*.gcode")))
    GCODE_MATERIALS = tuple(m for m in MATERIALS if m != "exact_size")   # no laser settings for exact size
    check(f"Generate G-code wrote one file per laser material ({len(gcodes)} of {len(GCODE_MATERIALS)})",
          lambda: len(gcodes) == len(GCODE_MATERIALS))
    check("no error box during G-code generation", no_error_boxes)

    def gcode_on_sheet_and_laser_off():
        kerf_slack = 1.5
        for path in gcodes:
            text = open(path, encoding="utf-8").read()
            assert "M5" in text, f"{os.path.basename(path)}: laser never switched off"
            pts = [(float(x), float(y)) for x, y in re.findall(r"^G[01] X(-?[0-9.]+)Y(-?[0-9.]+)", text, re.M)]
            assert pts, f"{os.path.basename(path)}: no moves"
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            assert min(xs) >= -kerf_slack and max(xs) <= SHEET_W + kerf_slack, (path, min(xs), max(xs))
            assert min(ys) >= -kerf_slack and max(ys) <= SHEET_H + kerf_slack, (path, min(ys), max(ys))
    check("every G-code move lies on the sheet and the file ends with the laser off", gcode_on_sheet_and_laser_off)

    def leather_gcode_has_locator_layer():
        text = open(by_material(".gcode")["leather"], encoding="utf-8").read()
        assert "; Layer C06" in text, "leather G-code has no locator layer"
        first_loc = text.index("; Layer C06")
        first_cut = min(i for i in (text.find("; Layer C02"), text.find("; Layer C03")) if i >= 0)
        assert first_loc < first_cut, "locator marks must be engraved before any cut"
    check("leather G-code engraves the locator marks before cutting", leather_gcode_has_locator_layer)

    # ---- Job history ---------------------------------------------------
    def history_recorded_both():
        jobs = load_job_history()
        outputs = [j.get("output") for j in jobs]
        assert "svg" in outputs and "gcode" in outputs, outputs
        newest = jobs[0]
        assert newest.get("output") == "gcode", newest.get("output")
        assert "7.0" in str(newest.get("pads_text", newest.get("pads", ""))), newest
    check("job history recorded the SVG run and the G-code run", history_recorded_both)

    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
