"""Key Height Library, Serial Lookup and Screw Specs through their real handlers.

  - Key heights: create a library, fill make/model/size and two heights,
    Save Key Height Set (name prompt stubbed), reload it, delete it.
  - Serial lookup: pick a maker, type a serial, the result label shows what
    lookup_serial_year says; a nonsense serial shows the no-match text.
  - Screw specs: type a maker and model, fill two fields and the notes,
    save, see it in the data and on disk, delete it.

Only prompts and message boxes are stubbed. Isolated profile
(SAXSHOP_CONFIG_DIR) so the user's libraries are never touched.
Self-skips without a display.
"""
import json
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-library-e2e-")
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

    print("Library tabs end to end")
    print("=" * 60)
    import main as main_mod
    import library_features as lf
    from config import KEY_PRESET_FILE, SCREW_SPECS_FILE
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()

    boxes = []
    for mod in (lf, main_mod):
        for name in ("showinfo", "showwarning", "showerror"):
            setattr(mod.messagebox, name, lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg)[:100])))
        mod.messagebox.askyesno = lambda *a, **k: True
    lf.simpledialog.askstring = lambda *a, **k: "E2E Set"

    # ---- Key heights -----------------------------------------------------
    app.notebook.select(app.key_tab)
    root.update()

    def key_save():
        app.key_presets["E2E Library"] = {}
        app.update_key_library_dropdown()
        app.key_library_var.set("E2E Library")
        app.on_key_library_selected()
        app.key_field_vars["make"].set("Selmer")
        app.key_field_vars["model"].set("Mark VI")
        app.key_field_vars["size"].set("Tenor")
        heights = list(app.key_height_vars.items())
        assert heights, "no key height fields"
        heights[0][1].set("7.5")
        heights[1][1].set("8.0")
        boxes.clear()
        app.on_save_key_preset()
        assert not [b for b in boxes if b[0] in ("showerror", "showwarning")], boxes
        assert "E2E Set" in app.key_presets["E2E Library"], list(app.key_presets["E2E Library"])
        saved = app.key_presets["E2E Library"]["E2E Set"]
        assert saved.get("make") == "Selmer" and saved.get("size") == "Tenor", saved
        on_disk = json.load(open(KEY_PRESET_FILE, encoding="utf-8"))
        assert "E2E Set" in on_disk.get("E2E Library", {}), "key set not written to the profile"
    check("key heights: save a set into a new library (on disk too)", key_save)

    def key_reload_and_delete():
        for v in app.key_field_vars.values():
            if hasattr(v, "set"):          # the notes field is a Text widget
                v.set("")
        app.key_preset_var.set("E2E Set")
        app.on_load_key_preset("E2E Set")
        assert app.key_field_vars["make"].get() == "Selmer", "reload did not restore the make"
        app.on_delete_key_preset()
        assert "E2E Set" not in app.key_presets.get("E2E Library", {}), "set not deleted"
    check("key heights: reload the set, then delete it", key_reload_and_delete)

    # ---- Serial lookup ---------------------------------------------------
    app.notebook.select(app.serial_tab)
    root.update()

    def serial_lookup():
        makers = list(app.serial_maker_dropdown.cget("values"))
        assert makers, "no makers in the dropdown"
        maker = makers[0]
        app.serial_maker_var.set(maker)
        app.serial_entry_var.set("55000")
        root.update()
        shown = app.serial_result_label.cget("text")
        expected = lf.lookup_serial_year(maker, "55000")
        assert shown == expected, (shown, expected)
        assert shown not in ("", "..."), shown
    check("serial lookup: result label matches lookup_serial_year", serial_lookup)

    def serial_nonsense():
        app.serial_entry_var.set("not-a-serial")
        root.update()
        shown = app.serial_result_label.cget("text")
        assert shown and shown == lf.lookup_serial_year(app.serial_maker_var.get(), "not-a-serial"), shown
    check("serial lookup: nonsense input shows the no-match text, no crash", serial_nonsense)

    # ---- Screw specs -----------------------------------------------------
    app.notebook.select(app.screw_tab)
    root.update()

    def screw_save():
        app.screw_maker_var.set("E2E Maker")
        app.screw_model_var.set("E2E Model")
        fields = list(app.screw_vars.items())
        assert len(fields) >= 2, "no screw spec fields"
        fields[0][1].set("M4x0.7")
        fields[1][1].set("neck")
        app.screw_notes_text.delete("1.0", tk.END)
        app.screw_notes_text.insert("1.0", "end-to-end test")
        boxes.clear()
        app.save_screw_spec()
        assert not [b for b in boxes if b[0] in ("showerror", "showwarning")], boxes
        spec = app.screw_data["E2E Maker"]["E2E Model"]
        assert spec[fields[0][0]] == "M4x0.7" and spec["notes"] == "end-to-end test", spec
        on_disk = json.load(open(SCREW_SPECS_FILE, encoding="utf-8"))
        assert "E2E Model" in on_disk.get("E2E Maker", {}), "screw spec not written to the profile"
    check("screw specs: save a new maker/model (on disk too)", screw_save)

    def screw_delete():
        app.screw_maker_var.set("E2E Maker")
        app.screw_model_var.set("E2E Model")
        app.delete_screw_spec()
        assert "E2E Maker" not in app.screw_data, "maker left behind after deleting its only model"
    check("screw specs: delete removes the model and the empty maker", screw_delete)

    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
