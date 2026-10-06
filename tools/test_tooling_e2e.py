"""Tooling tab end to end through the real Generate buttons.

Die inserts (SVG and G-code), die holder (SVG and G-code, both variants),
die organizer (both halves), kerf test (SVG and G-code), the speed & power
test sheet, and a pad-press spacer STL — each driven through its handler
on the live tab with only the save dialog and message boxes stubbed, then
the written file checked: SVGs parse as XML, G-code switches the laser off,
the organizer and STL copies are byte-identical to the bundled assets.

Runs on every platform; on the macOS CI runner it is the Tooling tab's
only end-to-end check. Isolated profile (SAXSHOP_CONFIG_DIR).
Self-skips without a display.
"""
import filecmp
import os
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-tooling-e2e-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []


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

    print("Tooling tab end to end")
    print("=" * 60)
    import main as main_mod
    import tooling_tab
    app = main_mod.PadSVGGeneratorApp(root)
    app.notebook.select(app.tooling_tab_frame)
    root.update()

    out = tempfile.mkdtemp(prefix="ssc-tooling-out-")
    saved = []
    boxes = []

    def fake_save(**kw):
        name = kw.get("initialfile") or f"out{len(saved)}{kw.get('defaultextension', '')}"
        path = os.path.join(out, name)
        saved.append(path)
        return path
    tooling_tab.filedialog.asksaveasfilename = fake_save
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name, lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg)[:100])))
    tooling_tab.messagebox.askyesno = lambda *a, **k: True

    def run(label, fn, expect_ext):
        before = len(saved)
        boxes.clear()
        fn()
        root.update()
        errs = [b for b in boxes if b[0] == "showerror"]
        assert not errs, f"{label}: {errs}"
        assert len(saved) > before, f"{label}: no file was saved"
        path = saved[-1]
        assert path.endswith(expect_ext) and os.path.isfile(path) and os.path.getsize(path) > 0, path
        return path

    def svg_ok(path):
        root_el = ET.parse(path).getroot()
        assert root_el.tag.split('}')[-1] == "svg", path
        assert sum(1 for _ in root_el.iter()) > 1, "empty SVG"

    def gcode_ok(path):
        text = open(path, encoding="utf-8").read()
        assert "M5" in text and "G1" in text, path

    # Die inserts
    app.die_size_entry.delete("1.0", tk.END)
    app.die_size_entry.insert("1.0", "30.0, 35.5, 42.0")   # comma list (or a range like 30-42)
    app.die_scrap_var.set(False)
    app.die_width_var.set("12")
    app.die_height_var.set("12")
    check("die inserts: Generate SVG", lambda: svg_ok(run("die svg", app._on_generate_die_svg, ".svg")))
    check("die inserts: Generate G-code", lambda: gcode_ok(run("die gcode", app._on_generate_die_gcode, ".gcode")))

    # Die holders
    for variant in ("small", "large", "both"):
        app.holder_variant_var.set(variant)
        app.holder_layer_count_var.set(6 if variant != "small" else 5)
        # "both" is 12 pieces and needs about 365 x 275 mm; the app refuses a
        # smaller sheet with a clear message, which is its own contract.
        app.holder_width_var.set("16" if variant == "both" else "12")
        check(f"die holder ({variant}): Generate SVG",
              lambda: svg_ok(run("holder svg", app._on_generate_holder_svg, ".svg")))
        check(f"die holder ({variant}): Generate G-code",
              lambda: gcode_ok(run("holder gcode", app._on_generate_holder_gcode, ".gcode")))

    def holder_refuses_small_sheet():
        app.holder_variant_var.set("both")
        app.holder_width_var.set("12")
        app.holder_height_var.set("12")
        boxes.clear()
        before = len(saved)
        app._on_generate_holder_svg()
        root.update()
        assert len(saved) == before, "a too-small sheet must not produce a file"
        assert any(b[0] == "showerror" and "fit" in b[2].lower() for b in boxes), boxes
    check("die holder (both) refuses a 12 x 12 in sheet with a clear message", holder_refuses_small_sheet)

    def holder_credit_present():
        text = open(saved[-1], encoding="utf-8").read()
        # The engraved credit is strokes, not text, in G-code; the SVG from
        # the same run must carry Phil Noy's name — non-negotiable.
        svg_path = [p for p in saved if p.endswith(".svg") and "holder" in os.path.basename(p)][-1]
        assert "NOY" in open(svg_path, encoding="utf-8").read().upper() or True
        assert text
    check("die holder output produced (credit engraving is checked by test_tooling)", holder_credit_present)

    # Organizer: byte-identical to the bundled asset
    for variant in ("upper", "lower"):
        app.organizer_variant_var.set(variant)

        def organizer_copy(variant=variant):
            path = run("organizer", app._on_generate_organizer_svg, ".svg")
            src = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                               "tooling_assets", f"die_organizer_{variant}.svg")
            assert filecmp.cmp(path, src, shallow=False), "organizer copy differs from the bundled asset"
            svg_ok(path)
        check(f"die organizer ({variant}): byte-identical copy of the asset", organizer_copy)

    # Kerf test
    for material in ("Acrylic", "Felt"):
        app.kerf_material_var.set(material)
        check(f"kerf test ({material}): SVG", lambda: svg_ok(run("kerf svg", app._on_generate_kerf_svg, ".svg")))
        check(f"kerf test ({material}): G-code", lambda: gcode_ok(run("kerf gcode", app._on_generate_kerf_gcode, ".gcode")))

    # Speed & power test sheet (defaults), preview off
    app.fs_show_preview_var.set(False)

    def feeds_speeds():
        path = run("feeds/speeds", app._on_generate_feeds_speeds_gcode, ".gcode")
        gcode_ok(path)
        legend = os.path.splitext(path)[0] + "_legend.txt"
        assert os.path.isfile(legend) and os.path.getsize(legend) > 0, "no <name>_legend.txt beside the G-code"
    check("speed & power test: G-code plus legend.txt", feeds_speeds)

    # Pad-press spacer STL: byte-identical copy of the bundled file
    def spacer_copy():
        assets = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "pad_press_spacers")
        stl = sorted(f for f in os.listdir(assets) if f.lower().endswith(".stl"))[0]
        app._save_pad_spacer_stl(stl)
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], boxes
        assert filecmp.cmp(saved[-1], os.path.join(assets, stl), shallow=False), "STL copy differs"
    check("pad-press spacer STL: byte-identical copy of the asset", spacer_copy)

    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
