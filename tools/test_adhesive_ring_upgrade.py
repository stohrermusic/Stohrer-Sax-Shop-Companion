"""Adhesive rings: upgrading an existing install, and switching the
feature on and off, through the real dialogs (2026-11-05).

Starts from a config folder exactly as a version before this feature left
it -- app_settings.json, sizing_presets.json and gcode_presets.json with no
trace of adhesive rings, the user's own values in them, and Matt's preset
(presets/Matt mode.json) applied -- and then uses the app the way a user
would: the real PadSVGGeneratorApp, the real Sizing Rules / G-code Settings
/ Layer Colors / Framing Power dialogs opened from their menu handlers,
their real Apply / Load / OK buttons. Only message boxes, the folder
picker and the Save Preset name prompt are answered by the test.

Covers:
  - the old folder loads without an error box; the user's values survive;
    the new keys arrive with their defaults; the new material gets its
    Default G-code preset on disk and the existing presets are untouched
  - the active sizing preset ("Matt mode") and the active G-code presets
    are still detected by the dialogs
  - off: no ring checkbox, Layer Colors row, G-code section or Framing
    Power row
  - on, through Sizing Rules: all four appear; settings and the new preset
    reach disk; the older preset is left as it was and still detected
  - tick the ring, switch the feature off again: all four go away, nothing
    is generated from the stale tick by Generate SVG, Generate G-code or
    Frame & Cut, and the ring's own laser settings are kept for next time
  - `python main.py --selftest` on the untouched old folder and on a
    folder with the feature on
  - the `--tour all` walk as a gate (run_tour) with the feature on and off:
    every stop, no errors, no stray windows, and the ring section and
    preview show in the tour only while the feature is on. What the tour's
    Sizing Rules stop showed is sampled when the tour closes that dialog,
    not on a timer, and the dialog is slowed past a whole tour step on
    purpose, so the gate cannot depend on how fast the machine is (a timer
    probe lost that race on CI's Intel Mac, 2026-11-05)

Needs a display (self-skips without one).
"""
import copy
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
_PROFILE = tempfile.mkdtemp(prefix="ssc-ring-upgrade-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

import config  # noqa: E402
from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

RING = 'adhesive_ring'
RING_KEYS = ("adhesive_ring_enabled", "adhesive_ring_trim_mm")

# The tour gate. run_tour schedules each stop's finish BEFORE it opens the
# stop, so a dialog that takes longer to build than one step is closed the
# moment it is built. That is the normal case on a slow runner, and here it
# is forced: the wrapped Sizing Rules dialog sleeps longer than a step.
TOUR_STEP_MS = 350
SLOW_DIALOG_S = float(os.environ.get("SSC_TEST_SLOW_DIALOG_S", "0.45"))
results = []


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        box_guard.settle()      # a question this check did not expect fails it
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}{box_guard.reset()}")
        traceback.print_exc()
        results.append(False)


def read_json(name, folder=None):
    with open(os.path.join(folder or _PROFILE, name), encoding="utf-8") as f:
        return json.load(f)


def write_old_profile(folder):
    """What a pre-feature version wrote: today's defaults with every trace
    of adhesive rings removed (checked against v2.9's config.py when this
    suite was written: identical), plus a user's own values."""
    with open(os.path.join(REPO, "presets", "Matt mode.json"), encoding="utf-8") as f:
        raw = json.load(f)
    matt = raw if any(k in raw for k in config.SIZING_PRESET_KEYS) else raw[next(iter(raw))]
    matt = {k: v for k, v in matt.items() if k not in RING_KEYS}

    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    for k in RING_KEYS:
        del s[k]
    del s["gcode_settings"][RING]
    del s["layer_colors"]["adhesive_ring_cut"]
    del s["laser_framing_power_by_material"][RING]
    default_preset = {k: copy.deepcopy(s[k]) for k in config.SIZING_PRESET_KEYS if k in s}
    default_gcode = {m: {"Default": copy.deepcopy(s["gcode_settings"][m])} for m in s["gcode_settings"]}

    # The user's own state on top.
    s.update(copy.deepcopy(matt))                        # Matt mode is the applied sizing preset
    s["gcode_settings"]["felt"]["cut_speed"] = 613
    s["gcode_settings"]["felt"]["cut_power"] = 72.5
    s["layer_colors"]["felt_outline"] = "#00A000"
    s["laser_framing_power_by_material"] = {"felt": 12, "card": 10, "leather": 25, "acrylic": 10, "basswood": 15}
    s["seen_preview_tutorial"] = True
    s["show_engraving_warning"] = False
    default_gcode["felt"]["Thick felt"] = copy.deepcopy(s["gcode_settings"]["felt"])
    default_gcode["leather"]["Goat"] = dict(default_gcode["leather"]["Default"], cut_power=88)

    files = {
        "app_settings.json": s,
        "sizing_presets.json": {"Default": default_preset, "Matt mode": matt},
        "gcode_presets.json": default_gcode,
    }
    for name, data in files.items():
        with open(os.path.join(folder, name), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    for name, data in files.items():
        assert "adhesive" not in json.dumps(data), f"{name}: the 'old' profile mentions adhesive rings"
    return files


def selftest(folder):
    env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder)
    p = subprocess.run([sys.executable, "main.py", "--selftest"], cwd=REPO, env=env,
                       capture_output=True, timeout=300)
    return p.returncode, (p.stdout + p.stderr).decode("utf-8", "replace")


def main():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    print("Adhesive rings: upgrade and on/off lifecycle")
    print("=" * 60)
    old = write_old_profile(_PROFILE)
    pristine = tempfile.mkdtemp(prefix="ssc-ring-upgrade-pristine-")
    for name in old:
        shutil.copy2(os.path.join(_PROFILE, name), pristine)

    import main as main_mod
    import ui_dialogs
    import falcon_sender
    from tkinter import messagebox as mb

    guard = box_guard.GUARD     # every question a check causes is named there, with its answer
    boxes = []
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(mb, name, lambda title=None, msg=None, _n=name, **k: boxes.append((_n, str(title), str(msg))))
    out_dir = tempfile.mkdtemp(prefix="ssc-ring-upgrade-out-")
    main_mod.filedialog.askdirectory = lambda **k: out_dir

    app = main_mod.PadSVGGeneratorApp(root)

    # ---- open the real dialogs from their menu handlers, keeping a handle ----
    def capture(cls_name):
        """Swap a dialog class in main for a subclass that records itself."""
        real = getattr(main_mod, cls_name)
        made = []

        class Recording(real):
            def __init__(self, *a, **k):
                self.init_kwargs = k
                super().__init__(*a, **k)
                made.append(self)
        setattr(main_mod, cls_name, Recording)
        return real, made

    def open_dialog(cls_name, opener):
        real, made = capture(cls_name)
        try:
            opener()
        finally:
            setattr(main_mod, cls_name, real)
        assert len(made) == 1, f"{cls_name}: {len(made)} dialogs opened"
        made[0].top.withdraw()
        return made[0]

    def sizing_rules():
        return open_dialog("OptionsWindow", app.open_options_window)

    def gcode_settings():
        return open_dialog("GcodeSettingsWindow", app.open_gcode_settings_window)

    def layer_colors():
        return open_dialog("LayerColorWindow", app.open_color_window)

    def framing_power_rows():
        before = set(root.winfo_children())
        app._on_machine_framing_power()
        dlg = [c for c in root.winfo_children() if c not in before and hasattr(c, "vars_by_mat")][0]
        rows = dict((m, v.get()) for m, v in dlg.vars_by_mat.items())
        dlg.destroy()
        return rows

    def close(w):
        try:
            w.top.grab_release()
            w.top.destroy()
        except tk.TclError:
            pass

    save_names = []

    class NamePrompt:
        """Stands in for the Save Preset name prompt (a modal wait)."""
        def __init__(self, parent, existing_names, default_existing=None, title=None, intro=None):
            self.result = {"name": save_names.pop(0)} if save_names else None

    def ring_ui_state():
        g = gcode_settings()
        lc = layer_colors()
        try:
            return {
                "checkbox": RING in app.material_vars and RING in app.material_checkboxes,
                "gcode": RING in g.vars,
                "layer": "adhesive_ring_cut" in lc.color_vars,
                "framing": RING in framing_power_rows(),
            }
        finally:
            close(g)
            close(lc)

    # ------------------------------------------------------------ upgrade ----
    def old_folder_loads_cleanly():
        assert not [b for b in boxes if b[0] != "showinfo"], f"boxes at startup: {boxes}"
        s = app.settings
        d = config.DEFAULT_SETTINGS
        assert s["adhesive_ring_enabled"] is False and s["adhesive_ring_trim_mm"] == 0.0
        assert s["gcode_settings"][RING] == d["gcode_settings"][RING]
        assert s["layer_colors"]["adhesive_ring_cut"] == d["layer_colors"]["adhesive_ring_cut"]
        assert s["laser_framing_power_by_material"][RING] == 10
        # The user's own values are all still there.
        assert s["gcode_settings"]["felt"]["cut_speed"] == 613 and s["gcode_settings"]["felt"]["cut_power"] == 72.5
        assert s["layer_colors"]["felt_outline"] == "#00A000"
        assert s["laser_framing_power_by_material"]["leather"] == 25
        assert s["laser_framing_power_by_material"]["basswood"] == 15
        for k, v in old["sizing_presets.json"]["Matt mode"].items():
            assert s[k] == v, f"Matt mode value {k} changed on load"
    check("a pre-feature config folder loads: no error, user values kept, new keys at their defaults",
          old_folder_loads_cleanly)

    def new_material_gets_its_default_gcode_preset_on_disk():
        disk = read_json("gcode_presets.json")
        assert disk[RING] == {"Default": config.DEFAULT_SETTINGS["gcode_settings"][RING]}, disk.get(RING)
        for mat, presets in old["gcode_presets.json"].items():
            assert disk[mat] == presets, f"{mat} presets changed by the backfill"
        assert app.gcode_presets == disk
        # Nothing else was rewritten at startup.
        assert read_json("sizing_presets.json") == old["sizing_presets.json"]
        assert read_json("app_settings.json") == old["app_settings.json"]
    check("the new material gets a Default G-code preset on disk; existing presets and files untouched",
          new_material_gets_its_default_gcode_preset_on_disk)

    def active_presets_still_detected():
        w = sizing_rules()
        try:
            assert w.active_preset_name == "Matt mode" and w.preset_combo.get() == "Matt mode", \
                (w.active_preset_name, w.preset_combo.get())
            assert not w._is_dirty()
            assert w.ring_enabled_var.get() is False and not w.lesser_open
            assert "beta" in w.ring_frame.cget("text").lower()
        finally:
            close(w)
        g = gcode_settings()
        try:
            assert [k for k, _label in g.active_materials] == ["felt", "card", "leather"]
            assert g.active_preset_name["felt"] == "Thick felt", g.active_preset_name
            assert g.active_preset_name["card"] == "Default" and g.active_preset_name["leather"] == "Default"
            assert not g._dirty_materials()
        finally:
            close(g)
    check("Sizing Rules still names 'Matt mode' and G-code Settings still names the active presets",
          active_presets_still_detected)

    def nothing_ring_shaped_while_off():
        state = ring_ui_state()
        assert state == {"checkbox": False, "gcode": False, "layer": False, "framing": False}, state
        assert list(app.material_vars) == ["felt", "card", "leather", "exact_size"]
    check("feature off: no ring checkbox, Layer Colors row, G-code section or Framing Power row",
          nothing_ring_shaped_while_off)

    # ---------------------------------------------------------- switch on ----
    def switch_on_through_sizing_rules():
        keep = ui_dialogs.SaveSizingPresetDialog
        ui_dialogs.SaveSizingPresetDialog = NamePrompt
        save_names[:] = ["Matt rings"]
        w = sizing_rules()
        try:
            w.ring_cb.invoke()                       # the real checkbox
            assert w.ring_enabled_var.get() is True
            assert str(w.ring_clearance_spin.cget("state")) == "normal"
            w.ring_trim_var.set(0.3)
            assert w._is_dirty()
            guard.expect("Save changes as a preset?", answer=True)
            w.on_apply()                             # asks to save as a preset first -> "Matt rings"
        finally:
            ui_dialogs.SaveSizingPresetDialog = keep
            close(w)
        assert app.settings["adhesive_ring_enabled"] is True and app.settings["adhesive_ring_trim_mm"] == 0.3
        disk = read_json("app_settings.json")
        assert disk["adhesive_ring_enabled"] is True and disk["adhesive_ring_trim_mm"] == 0.3
        presets = read_json("sizing_presets.json")
        assert presets["Matt rings"]["adhesive_ring_enabled"] is True
        assert presets["Matt rings"]["adhesive_ring_trim_mm"] == 0.3
        assert presets["Matt mode"] == old["sizing_presets.json"]["Matt mode"], "the older preset was rewritten"
        assert presets["Default"] == old["sizing_presets.json"]["Default"]
        state = ring_ui_state()
        assert state == {"checkbox": True, "gcode": True, "layer": True, "framing": True}, state
        assert app.material_vars[RING].get() is False, "the new material arrives unticked"
    check("switched on in Sizing Rules: saved as a preset, on disk, and all four pieces of UI appear",
          switch_on_through_sizing_rules)

    def dialogs_after_switching_on():
        w = sizing_rules()
        try:
            assert w.active_preset_name == "Matt rings" and w.lesser_open and not w._is_dirty()
        finally:
            close(w)
        g = gcode_settings()
        try:
            assert g.active_preset_name[RING] == "Default" and g.active_preset_name["felt"] == "Thick felt"
            assert set(g.vars[RING]) == {"cut", "kerf_width", "air_assist_cut"}
            # The user tunes the tape and applies.
            g.vars[RING]["cut"]["power"].set(38.0)
            g.vars[RING]["cut"]["speed"].set(1350)
            guard.expect("Unsaved changes", answer=False)     # "apply anyway", no preset
            g._on_apply_clicked()
        finally:
            close(g)
        disk = read_json("app_settings.json")["gcode_settings"]
        assert disk[RING]["cut_power"] == 38.0 and disk[RING]["cut_speed"] == 1350
        assert set(disk[RING]) == set(config.GCODE_CUT_ONLY_KEYS)
        assert disk["felt"] == old["app_settings.json"]["gcode_settings"]["felt"], "felt block changed"
        rows = framing_power_rows()
        assert rows[RING] == "1.0" and rows["leather"] == "2.5"
    check("after switching on: presets detected, the tape's G-code settings edit and save on their own",
          dialogs_after_switching_on)

    # ---- generate with the ring ticked, then switch off with the tick stale ----
    def fill_form():
        app.preview_var.set(False)
        app.scrap_mode_var.set(False)
        app.custom_polygon = None
        app.custom_polygon_outline = None
        app.pad_entry.delete("1.0", tk.END)
        app.pad_entry.insert("1.0", "7.0 x 3\n12.5 x 2\n20.0 x 2")
        app.width_entry.delete(0, tk.END)
        app.width_entry.insert(0, "120" if app.settings["units"] == "mm" else "5")
        app.height_entry.delete(0, tk.END)
        app.height_entry.insert(0, "90" if app.settings["units"] == "mm" else "4")
        app.filename_entry.delete(0, tk.END)
        app.filename_entry.insert(0, "up")
        root.update()

    def outputs():
        return sorted(os.path.basename(p) for p in glob.glob(os.path.join(out_dir, "*")))

    def ring_generates_while_on():
        fill_form()
        for m, v in app.material_vars.items():
            v.set(m == RING)
        boxes.clear()
        app.on_generate_svg()
        app.on_generate_gcode()
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], boxes
        assert outputs() == ["up_adhesive_ring.gcode", "up_adhesive_ring.svg"], outputs()
        text = open(os.path.join(out_dir, "up_adhesive_ring.gcode"), encoding="utf-8").read()
        assert "; Cut @ 1350 mm/min, 38.0% power" in text, "the tuned tape settings were not used"
    check("while on: Generate SVG and Generate G-code write the ring files at the tuned settings",
          ring_generates_while_on)

    def switch_off_with_the_ring_still_ticked():
        assert app.material_vars[RING].get() is True, "fixture: the ring should still be ticked"
        for f in glob.glob(os.path.join(out_dir, "*")):
            os.remove(f)
        w = sizing_rules()
        try:
            w.preset_combo.set("Matt mode")          # the pre-feature preset
            w.on_load_sizing_preset()
            assert w.ring_enabled_var.get() is False and not w._is_dirty()
            w.on_apply()
        finally:
            close(w)
        assert app.settings["adhesive_ring_enabled"] is False
        assert read_json("app_settings.json")["adhesive_ring_enabled"] is False
        state = ring_ui_state()
        assert state == {"checkbox": False, "gcode": False, "layer": False, "framing": False}, state
        assert RING not in app.material_vars

        # Nothing is ticked now (the only tick was the ring's): no output at all.
        fill_form()
        boxes.clear()
        app.on_generate_svg()
        app.on_generate_gcode()
        root.update()
        assert outputs() == [], f"files written from a stale ring tick: {outputs()}"
        assert [b[1] for b in boxes] == ["No Materials Selected", "No Materials Selected"], boxes

        # With felt ticked, felt and only felt comes out.
        app.material_vars["felt"].set(True)
        boxes.clear()
        app.on_generate_svg()
        app.on_generate_gcode()
        root.update()
        assert not [b for b in boxes if b[0] == "showerror"], boxes
        assert outputs() == ["up_felt.gcode", "up_felt.svg"], outputs()

        # Frame & Cut finds no ring either.
        app.material_vars["felt"].set(False)
        runs = []

        class Run:
            def __init__(self, parent, sender, lines, **kw):
                runs.append(lines)
                self._final_reason = "complete"

        class Sender:
            def __init__(self, port=None):
                pass

            def connect(self):
                pass

            def unlock(self):
                pass

            def disconnect(self):
                pass
        keep = (ui_dialogs.FalconRunDialog, falcon_sender.FalconSender)
        ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = Run, Sender
        app.falcon_port = "TEST"
        boxes.clear()
        try:
            app.on_frame_and_cut()
        finally:
            ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = keep
            app.falcon_port = None
        assert not runs and [b[1] for b in boxes] == ["Pick One Material"], (runs, boxes)

        # The tape's own laser settings, framing power and layer color wait for next time.
        disk = read_json("app_settings.json")
        assert disk["gcode_settings"][RING]["cut_power"] == 38.0
        assert disk["laser_framing_power_by_material"][RING] == 10
        assert "adhesive_ring_cut" in disk["layer_colors"]
    check("switched off with the ring still ticked: all four pieces of UI go, and Generate and "
          "Frame & Cut produce nothing from the stale tick", switch_off_with_the_ring_still_ticked)

    def hidden_rows_keep_their_values():
        # Saving Layer Colors and Framing Power while the ring rows are hidden must not drop them.
        app.settings["laser_framing_power_by_material"][RING] = 21
        app.settings["layer_colors"]["adhesive_ring_cut"] = "#00A0FF"
        lc = layer_colors()
        lc.save_colors()
        before = set(root.winfo_children())
        app._on_machine_framing_power()
        dlg = [c for c in root.winfo_children() if c not in before and hasattr(c, "vars_by_mat")][0]
        ok = [b for fr in dlg.winfo_children() for b in fr.winfo_children()
              if isinstance(b, tk.Button) and b.cget("text") == "OK"]
        ok[0].invoke()
        disk = read_json("app_settings.json")
        assert disk["laser_framing_power_by_material"][RING] == 21, disk["laser_framing_power_by_material"]
        assert disk["layer_colors"]["adhesive_ring_cut"] == "#00A0FF"
    check("saving Layer Colors and Framing Power while the feature is off keeps the ring's hidden values",
          hidden_rows_keep_their_values)

    def switch_back_on_restores_everything():
        w = sizing_rules()
        try:
            w.preset_combo.set("Matt rings")
            w.on_load_sizing_preset()
            w.on_apply()
        finally:
            close(w)
        assert app.settings["adhesive_ring_enabled"] is True and app.settings["adhesive_ring_trim_mm"] == 0.3
        assert app.material_vars[RING].get() is False, "comes back unticked, not with the old tick"
        g = gcode_settings()
        try:
            assert g.vars[RING]["cut"]["power"].get() == 38.0
        finally:
            close(g)
        assert framing_power_rows()[RING] == "2.1"
    check("switched back on: the trim, the tape's laser settings and framing power are as they were left",
          switch_back_on_restores_everything)

    # ----------------------------------------------------------- selftest ----
    def selftest_on_the_old_folder():
        rc, out = selftest(pristine)
        assert rc == 0 and "selftest ok" in out, f"rc={rc}\n{out[-800:]}"
        assert read_json("gcode_presets.json", pristine)[RING]["Default"] == \
            config.DEFAULT_SETTINGS["gcode_settings"][RING]
        assert read_json("app_settings.json", pristine) == old["app_settings.json"], "selftest saved settings"
    check("python main.py --selftest passes on the untouched pre-feature folder", selftest_on_the_old_folder)

    def selftest_with_the_feature_on():
        folder = tempfile.mkdtemp(prefix="ssc-ring-upgrade-on-")
        for name in ("app_settings.json", "sizing_presets.json", "gcode_presets.json"):
            shutil.copy2(os.path.join(_PROFILE, name), folder)
        assert read_json("app_settings.json", folder)["adhesive_ring_enabled"] is True
        rc, out = selftest(folder)
        assert rc == 0 and "selftest ok" in out, f"rc={rc}\n{out[-800:]}"
    check("python main.py --selftest passes with adhesive rings switched on", selftest_with_the_feature_on)

    # --------------------------------------------------------------- tour ----
    expected = ["pad-maker", "sizing-rules-with-preview", "layer-colors", "gcode-settings", "nesting-preview",
                "polygon-draw", "job-history", "feature-set", "user-guide", "about", "welcome",
                "quick-start", "key-heights",
                "key-layout", "serial-lookup", "screw-specs", "tooling", "tuner-no-mic", "tuner",
                "analyzer-chart", "analyzer-compare", "analyzer-preset-library", "analyzer-preset-form",
                "analyzer-harmonics"]
    if not main_mod.audio_supported():
        expected = [s for s in expected if s not in main_mod.AUDIO_TOUR_STOPS]

    def tour(enabled):
        app.settings["adhesive_ring_enabled"] = enabled
        app.update_ui_from_settings()
        app._toner_mic_checked = True
        seen = {"gcode": None, "layer": None, "preview_ring": None, "lesser_open": None, "sampled": 0}
        real_opts, real_gc, real_lc = main_mod.OptionsWindow, main_mod.GcodeSettingsWindow, main_mod.LayerColorWindow

        class Opts(real_opts):
            def __init__(self, *a, **k):
                super().__init__(*a, **k)
                # Slower than a whole tour step, on purpose: the stop's finish
                # is already due when this returns, exactly as on a slow CI
                # runner. Nothing below may depend on the dialog staying open
                # for any length of time.
                time.sleep(SLOW_DIALOG_S)
                # Sample what the stop showed at the moment the tour closes
                # the dialog (it calls top.destroy()): by then the tour has
                # opened the lesser-used section and the preview, however
                # long that took. No timer, so nothing to race.
                really_destroy = self.top.destroy

                def sample_then_destroy():
                    try:
                        seen["sampled"] += 1
                        seen["lesser_open"] = self.lesser_open
                        pv = self.preview_window
                        if pv is not None and pv.winfo_exists():
                            pv._render()
                            seen["preview_ring"] = len(pv.canvas.find_withtag("ring"))
                    finally:
                        really_destroy()
                self.top.destroy = sample_then_destroy

        class Gc(real_gc):
            def __init__(self, *a, **k):
                super().__init__(*a, **k)
                seen["gcode"] = [m for m, _label in self.active_materials]

        class Lc(real_lc):
            def __init__(self, *a, **k):
                super().__init__(*a, **k)
                seen["layer"] = "adhesive_ring_cut" in self.color_vars
        main_mod.OptionsWindow, main_mod.GcodeSettingsWindow, main_mod.LayerColorWindow = Opts, Gc, Lc
        try:
            log, errors = main_mod.run_tour(root, app, shots_dir=None, step_ms=TOUR_STEP_MS, on_done=root.quit)
            guard = root.after(180000, root.quit)
            root.mainloop()
            root.after_cancel(guard)
        finally:
            main_mod.OptionsWindow, main_mod.GcodeSettingsWindow, main_mod.LayerColorWindow = real_opts, real_gc, real_lc
        stray = [w for w in root.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_exists()]
        for w in stray:
            w.destroy()
        return log, errors, stray, seen

    def tour_gate(enabled):
        def fn():
            log, errors, stray, seen = tour(enabled)
            assert log == expected, f"stops visited: {log}"
            assert not errors, "\n".join(errors)
            assert not stray, f"windows left open: {stray}"
            want_gc = ["felt", "card", "leather"] + ([RING] if enabled else [])
            assert seen["gcode"] == want_gc, seen["gcode"]
            assert seen["layer"] is enabled, seen["layer"]
            assert SLOW_DIALOG_S * 1000 > TOUR_STEP_MS, "the dialog must be slower than a tour step for this gate"
            assert seen["sampled"] == 1, f"the Sizing Rules stop was sampled {seen['sampled']} times, wanted once"
            assert seen["lesser_open"] is True, "the tour opens the lesser-used section"
            if enabled:
                assert seen["preview_ring"] and seen["preview_ring"] >= 1, f"no ring in the tour's preview: {seen}"
            else:
                assert seen["preview_ring"] == 0, f"ring drawn in the tour's preview while off: {seen}"
        return fn
    check(f"the tour gate with adhesive rings ON: all {len(expected)} stops, no errors, ring shown",
          tour_gate(True))
    check(f"the tour gate with adhesive rings OFF: all {len(expected)} stops, no errors, nothing ring-shaped",
          tour_gate(False))

    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == '__main__':
    sys.exit(main())
