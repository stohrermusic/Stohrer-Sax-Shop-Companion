"""Cork as a laser material, and "engraving not set".

Matt, 2026-11-09: "tonehole crescents are usually cork. we do not have that
as a material yet". Cork is a full G-code material for the Tooling tab: in
its settings dialog, the Kerf Test, the Speed & Power Test, Framing Power
and the Tonehole Crescent (its default). It is not a Pad Maker material.

Its cut default is Matt's own, tested on 3/32 in cork. Its engraving is 0
and 0 on purpose ("don't want a guess in there"), and 0 means NOT SET: a
generator must never write an F0 move or engrave at S0, so the Tooling
generators that engrave with the chosen material leave the engraving out,
still write the cuts, and the tab says so. This suite pins:

  - the defaults (cut 1000 mm/min, 30 %, 1 pass, air on, kerf 0, engraving
    0/0, framing power 10 and clamped like every material's)
  - a settings folder saved before this change gains cork (settings block,
    framing power, a Default G-code preset on disk) and loses nothing; a
    v2.9 folder gains adhesive rings and cork (run in a child process: a
    profile is fixed per process)
  - where cork is offered (the four Tooling places) and where it is not
    (the Pad Maker's materials and its G-code Settings)
  - the Tooling settings dialog shows cork with its note, and takes 0/0
  - Kerf Test and Speed & Power Test with cork: cuts all written, no
    engraving layer, no F0, the notice in the Done box and asked before a
    Frame & Cut run; real values bring the engraving back; other materials
    are untouched
  - "Cork" is translated in all four catalogs

Self-skips the GUI part without a display. Isolated profile.
"""
import copy
import gettext
import json
import os
import re
import subprocess
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
CHILD = sys.argv[2] if len(sys.argv) > 2 and sys.argv[1] == "--child" else None
_PROFILE = os.environ["SAXSHOP_CONFIG_DIR"] if CHILD else tempfile.mkdtemp(prefix="ssc-cork-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import config  # noqa: E402

results = []
RING = "adhesive_ring"
RING_KEYS = ("adhesive_ring_enabled", "adhesive_ring_trim_mm")
NEW_KEYS = ("welcomed_version",)


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


def write_old_profile(folder, with_rings):
    """A settings folder as an older version left it: today's defaults with
    every trace of cork removed (and of adhesive rings too for a v2.9
    folder), plus a user's own values and presets. Returns what was written."""
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    del s["gcode_settings"]["cork"]
    del s["laser_framing_power_by_material"]["cork"]
    for k in NEW_KEYS:
        s.pop(k, None)
    if not with_rings:
        for k in RING_KEYS:
            del s[k]
        del s["gcode_settings"][RING]
        del s["layer_colors"]["adhesive_ring_cut"]
        del s["laser_framing_power_by_material"][RING]
    s["gcode_settings"]["felt"]["cut_speed"] = 613
    s["gcode_settings"]["acrylic"]["kerf_width"] = 0.17
    s["laser_framing_power_by_material"]["leather"] = 25
    presets = {m: {"Default": copy.deepcopy(block), "Mine": dict(copy.deepcopy(block), cut_speed=block["cut_speed"] + 7)}
               for m, block in s["gcode_settings"].items()}
    with open(os.path.join(folder, os.path.basename(config.SETTINGS_FILE)), "w", encoding="utf-8") as f:
        json.dump(s, f)
    with open(os.path.join(folder, os.path.basename(config.GCODE_PRESET_FILE)), "w", encoding="utf-8") as f:
        json.dump(presets, f)
    return s, presets


def upgrade_report(app, wrote_settings, wrote_presets):
    """What the app made of an older folder, as plain data."""
    on_disk = json.load(open(config.GCODE_PRESET_FILE, encoding="utf-8"))
    return {
        "settings_materials": sorted(app.settings["gcode_settings"]),
        "cork_block_is_default": app.settings["gcode_settings"]["cork"] == config.DEFAULT_SETTINGS["gcode_settings"]["cork"],
        "ring_block_is_default": app.settings["gcode_settings"].get(RING) == config.DEFAULT_SETTINGS["gcode_settings"][RING],
        "preset_materials_on_disk": sorted(on_disk),
        "cork_default_preset": on_disk.get("cork", {}).get("Default"),
        "ring_default_preset": on_disk.get(RING, {}).get("Default"),
        "old_presets_untouched": all(on_disk.get(m) == wrote_presets[m] for m in wrote_presets),
        "user_values_kept": (app.settings["gcode_settings"]["felt"]["cut_speed"],
                             app.settings["gcode_settings"]["acrylic"]["kerf_width"],
                             app.settings["laser_framing_power_by_material"]["leather"]),
        "framing_cork": app.settings["laser_framing_power_by_material"].get("cork"),
        "framing_ring": app.settings["laser_framing_power_by_material"].get(RING),
        "old_blocks_untouched": all(app.settings["gcode_settings"][m] == wrote_settings["gcode_settings"][m]
                                    for m in wrote_settings["gcode_settings"]),
        "welcomed_version": app.settings.get("welcomed_version"),
    }


def child(scenario):
    """One older folder, in its own process: build the app, report as JSON."""
    import tkinter as tk
    wrote = write_old_profile(_PROFILE, with_rings=(scenario != "v2.9"))
    import main as main_mod
    root = tk.Tk()
    root.withdraw()
    app = main_mod.PadSVGGeneratorApp(root)
    print("REPORT " + json.dumps(upgrade_report(app, *wrote)))
    root.destroy()


if CHILD:
    child(CHILD)
    sys.exit(0)


# ------------------------------------------------------------------ headless

def defaults():
    c = config.DEFAULT_SETTINGS["gcode_settings"]["cork"]
    assert config.GCODE_PRESET_MATERIALS == ("felt", "card", "leather", "acrylic", "basswood", RING, "cork")
    assert config.TOOLING_GCODE_MATERIALS == ("acrylic", "basswood", "cork")
    assert "cork" not in config.GCODE_CUT_ONLY_MATERIALS and set(c) == set(config.GCODE_PRESET_KEYS), \
        "cork is a full material: every G-code preset key, nothing else"
    assert config.gcode_preset_keys_for("cork") == config.GCODE_PRESET_KEYS
    assert (c["cut_speed"], c["cut_power"], c["cut_passes"], c["air_assist_cut"]) == (1000, 30, 1, True)
    assert (c["hole_speed"], c["hole_power"], c["hole_passes"]) == (1000, 30, 1), "holes cut as the outline does"
    assert c["kerf_width"] == 0.0, "the kerf is not measured: 0 reproduces the crescent Matt cut"
    assert c["engraving_mode"] == "line" and c["engraving_passes"] == 1 and c["filled_engraving_passes"] == 1
    for key in ("engraving_speed", "engraving_power", "filled_engraving_speed", "filled_engraving_power"):
        assert c[key] == 0, f"{key} must ship as 0 (not set): no guess in there"
    assert config.DEFAULT_SETTINGS["laser_framing_power_by_material"]["cork"] == 10
    presets = config.settings_to_gcode_presets(copy.deepcopy(config.DEFAULT_SETTINGS))
    assert presets["cork"]["Default"] == c


check("defaults: cork is a full material, cut 1000 mm/min at 30 %, kerf 0, engraving 0/0, framing power 10", defaults)


def framing_power_is_clamped_for_cork_too():
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    assert config.get_framing_power_s("cork", s) == 10
    s["laser_framing_power_by_material"]["cork"] = 17
    assert config.get_framing_power_s("cork", s) == 17
    s["laser_framing_power_by_material"]["cork"] = 600
    assert config.get_framing_power_s("cork", s) == config.FRAMING_POWER_S_MAX
    del s["laser_framing_power_by_material"]["cork"]
    assert config.get_framing_power_s("cork", s) == s["laser_framing_power_s"], "a config without the entry falls back"


check("framing power for cork: its own value, clamped to the 3 % ceiling like every material", framing_power_is_clamped_for_cork_too)


def zero_means_not_set():
    ok = config.engraving_is_set
    assert ok(3000, 10) and ok("1200", "8") and ok(0.5, 0.1)
    for speed, power in ((0, 0), (3000, 0), (0, 10), (-5, 10), (3000, -1), (None, 10), ("", 10), ("fast", 10)):
        assert not ok(speed, power), (speed, power)


check("engraving_is_set: a speed or a power of 0 (or nothing usable) means the engraving is not set", zero_means_not_set)


def cork_is_translated():
    for lang in ("es", "de", "fr", "it"):
        tr = gettext.translation("saxshop", os.path.join(REPO, "locale"), languages=[lang])
        assert tr.gettext("Cork") not in ("Cork", ""), f"'Cork' is not translated in {lang}"


check("'Cork' is translated in all four catalogs", cork_is_translated)


def a_v29_folder_gains_rings_and_cork():
    folder = tempfile.mkdtemp(prefix="ssc-cork-v29-")
    env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
    p = subprocess.run([sys.executable, os.path.abspath(__file__), "--child", "v2.9"],
                       capture_output=True, text=True, env=env, timeout=300, cwd=REPO)
    if "no display" in (p.stderr or "") or "couldn't connect to display" in (p.stderr or ""):
        print("  (skipped: no display for the child process)")
        return
    line = [ln for ln in p.stdout.splitlines() if ln.startswith("REPORT ")]
    assert p.returncode == 0 and line, (p.returncode, p.stdout[-400:], p.stderr[-800:])
    r = json.loads(line[0][7:])
    assert r["settings_materials"] == sorted(config.GCODE_PRESET_MATERIALS), r["settings_materials"]
    assert r["preset_materials_on_disk"] == sorted(config.GCODE_PRESET_MATERIALS), r["preset_materials_on_disk"]
    assert r["cork_block_is_default"] and r["ring_block_is_default"]
    assert r["cork_default_preset"] == config.DEFAULT_SETTINGS["gcode_settings"]["cork"]
    assert r["ring_default_preset"] == config.DEFAULT_SETTINGS["gcode_settings"][RING]
    assert r["old_presets_untouched"] and r["old_blocks_untouched"], r
    assert r["user_values_kept"] == [613, 0.17, 25], r["user_values_kept"]
    assert (r["framing_cork"], r["framing_ring"]) == (10, 10)
    assert r["welcomed_version"] == "", "an updated install has not been welcomed yet"


check("upgrade: a v2.9 folder gains adhesive rings and cork (settings, framing power, Default presets), loses nothing",
      a_v29_folder_gains_rings_and_cork)


# ----------------------------------------------------------------------- GUI

def gui():
    import tkinter as tk
    from tkinter import ttk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"  (GUI checks skipped: no display: {e})")
        return
    root.withdraw()
    wrote = write_old_profile(_PROFILE, with_rings=True)        # the folder as the version before this one left it
    import gcode_engine as ge
    import main as main_mod
    import tooling_tab
    import ui_dialogs
    import falcon_sender
    app = main_mod.PadSVGGeneratorApp(root)
    out = tempfile.mkdtemp(prefix="ssc-cork-out-")
    saved, boxes = [], []
    guard = box_guard.GUARD     # a check names each question it causes there, with its answer

    def fake_save(**kw):
        path = os.path.join(out, kw.get("initialfile"))
        saved.append(path)
        return path
    tooling_tab.filedialog.asksaveasfilename = fake_save
    for name in ("showinfo", "showwarning", "showerror"):
        setattr(tooling_tab.messagebox, name,
                lambda title, msg, _n=name, **k: boxes.append((_n, str(title), str(msg))))

    app._show_kerf_instructions = lambda name: None
    note = ui_dialogs.gcode_file_note()

    def the_previous_versions_folder_gains_cork():
        r = upgrade_report(app, *wrote)
        assert r["settings_materials"] == sorted(config.GCODE_PRESET_MATERIALS)
        assert r["cork_block_is_default"] and r["framing_cork"] == 10
        assert r["preset_materials_on_disk"] == sorted(config.GCODE_PRESET_MATERIALS), \
            "the new material's Default preset is saved at once (first-launch backfill)"
        assert r["cork_default_preset"] == config.DEFAULT_SETTINGS["gcode_settings"]["cork"]
        assert r["old_presets_untouched"] and r["old_blocks_untouched"], "something of the user's was changed"
        assert r["user_values_kept"] == (613, 0.17, 25), r["user_values_kept"]
        assert app.gcode_presets["cork"] == {"Default": config.DEFAULT_SETTINGS["gcode_settings"]["cork"]}
    check("upgrade: the previous version's folder gains cork (settings, framing power, Default preset) and loses nothing",
          the_previous_versions_folder_gains_cork)

    def combo_values(section, var):
        frame = app._tooling_sections[section]
        found = []

        def walk(w):
            for c in w.winfo_children():
                if isinstance(c, ttk.Combobox) and str(c.cget("textvariable")) == str(var):
                    found.append(list(c.cget("values")))
                walk(c)
        walk(frame)
        assert len(found) == 1, f"{section}: {len(found)} material dropdowns"
        return found[0]

    def offered_on_tooling_not_on_the_pad_maker():
        assert combo_values("kerf_test", app.kerf_material_var) == ["Felt", "Card", "Leather", "Acrylic", "Cork"]
        assert combo_values("feeds_speeds_tester", app.fs_material_var) == \
            ["Felt", "Card", "Leather", "Acrylic", "Basswood", "Cork"]
        crescent = combo_values("tonehole_crescent", app.crescent_material_var)
        assert crescent[0] == "Cork" and app.crescent_material_var.get() == "Cork", "cork is the crescent's default"
        seen = {}
        keep_t, keep_m = tooling_tab.GcodeSettingsWindow, main_mod.GcodeSettingsWindow

        class Spy:
            def __init__(self, parent, settings, cb, materials=None, **kw):
                seen[kw.get("show_tooling_engraving", False)] = [k for k, _label in materials]
        tooling_tab.GcodeSettingsWindow = main_mod.GcodeSettingsWindow = Spy
        try:
            app._open_tooling_gcode_settings()
            app.settings["adhesive_ring_enabled"] = True
            app.open_gcode_settings_window()
        finally:
            tooling_tab.GcodeSettingsWindow, main_mod.GcodeSettingsWindow = keep_t, keep_m
            app.settings["adhesive_ring_enabled"] = False
        assert seen[True] == ["acrylic", "basswood", "cork"], f"Tooling > Options > Settings: {seen[True]}"
        assert "cork" not in seen[False] and seen[False][:3] == ["felt", "card", "leather"], \
            f"the Pad Maker's G-code Settings: {seen[False]}"
        assert "cork" not in app.material_vars, "cork is not a Pad Maker material"
        app._on_machine_framing_power()
        dlg = [c for c in root.winfo_children() if hasattr(c, "vars_by_mat")][0]
        try:
            assert "cork" in dlg.vars_by_mat and dlg.vars_by_mat["cork"].get() == "1.0"
        finally:
            dlg.grab_release()
            dlg.destroy()
    check("cork is offered in the Kerf Test, Speed & Power, Crescent, Tooling settings and Framing Power; not on the Pad Maker",
          offered_on_tooling_not_on_the_pad_maker)

    def the_settings_dialog_shows_cork_and_takes_zero():
        s = copy.deepcopy(app.settings)
        applied = []
        w = ui_dialogs.GcodeSettingsWindow(root, s, lambda st: applied.append(True),
                                           materials=[("cork", "Cork")], show_tooling_engraving=True)
        w.top.withdraw()
        try:
            v = w.vars["cork"]
            assert (v["engraving"]["speed"].get(), v["engraving"]["power"].get()) == (0, 0.0)
            assert (v["filled_engraving"]["speed"].get(), v["filled_engraving"]["power"].get()) == (0, 0.0)
            assert (v["cut"]["speed"].get(), v["cut"]["power"].get()) == (1000, 30.0)
            texts = []

            def walk(x):
                for c in x.winfo_children():
                    if isinstance(c, tk.Label):
                        texts.append(str(c.cget("text")))
                    walk(c)
            walk(w.top)
            notes = [t for t in texts if "3/32 in (2.4 mm) sheet cork" in t]
            assert len(notes) == 1, "the cork section should say what its defaults are worth"
            n = notes[0]
            assert "cut was tested" in n and "Kerf is 0" in n and "engraving is not set" in n
            assert "Kerf Test" in n and "Speed & Power Test" in n
            del boxes[:]
            ui_boxes = []
            keep = ui_dialogs.messagebox.showerror
            ui_dialogs.messagebox.showerror = lambda *a, **k: ui_boxes.append(a)
            try:
                w._on_save()
            finally:
                ui_dialogs.messagebox.showerror = keep
            assert not ui_boxes, f"0 and 0 must be accepted: {ui_boxes}"
            assert s["gcode_settings"]["cork"]["engraving_speed"] == 0 and s["gcode_settings"]["cork"]["engraving_power"] == 0
            assert s["gcode_settings"]["cork"]["cut_speed"] == 1000
        finally:
            if w.top.winfo_exists():
                w.top.destroy()
    check("Tooling settings dialog: cork shows its tested cut, 0/0 engraving and its note, and applies with 0/0",
          the_settings_dialog_shows_cork_and_takes_zero)

    def layers_in(path):
        return re.findall(r"^; Layer (\w+)", open(path).read(), flags=re.M)

    def no_zero_feed(path):
        bad = [ln for ln in open(path).read().splitlines() if re.search(r"F0(?![\d.])", ln)]
        assert not bad, f"a zero feed was written: {bad[:3]}"

    def click(fn):
        del boxes[:]
        before = len(saved)
        fn()
        root.update()
        return saved[before:], list(boxes)

    def set_cork_engraving(speed, power):
        app.settings["gcode_settings"]["cork"].update(engraving_speed=speed, engraving_power=power)

    def kerf_test_on_cork():
        app.kerf_material_var.set("Cork")
        app.kerf_use_existing_var.set(True)
        try:
            new, shown = click(app._on_generate_kerf_gcode)
            assert os.path.basename(new[0]) == "kerf_test_cork.gcode" and not [b for b in shown if b[0] == "showerror"], shown
            assert layers_in(new[0]) == ["C02"], "cork has no engraving settings: the three circles, no labels"
            text = open(new[0]).read()
            assert len(re.findall(r"^G0 X", text, flags=re.M)) == 3 and "F1000" in text and "S300" in text
            no_zero_feed(new[0])
            done = [b for b in shown if b[0] == "showinfo"]
            assert len(done) == 1, shown
            msg = done[0][2]
            assert msg.startswith(f"Saved to: {new[0]}") and msg.endswith(note)
            assert "Cork has no engraving speed and power set (Tooling > Options > Settings)" in msg and \
                "will not be engraved" in msg, msg
            # Real values bring the labels back, and the notice goes.
            set_cork_engraving(3000, 10)
            new, shown = click(app._on_generate_kerf_gcode)
            assert layers_in(new[0]) == ["C03", "C02"] and "F3000" in open(new[0]).read()
            assert "will not be engraved" not in [b for b in shown if b[0] == "showinfo"][0][2]
        finally:
            set_cork_engraving(0, 0)
    check("Kerf Test on cork: the circles are cut, no labels, no F0, and the Done box says so; real values bring them back",
          kerf_test_on_cork)

    def other_materials_are_untouched():
        app.kerf_material_var.set("Felt")
        new, shown = click(app._on_generate_kerf_gcode)
        mat = app.settings["gcode_settings"]["felt"]
        want = os.path.join(out, "want_felt.gcode")
        filled = mat.get("engraving_mode", "line") == "filled"
        ge.generate_kerf_test_gcode("Felt", want, app.settings, mat["cut_speed"], mat["cut_power"],
                                    mat["filled_engraving_speed" if filled else "engraving_speed"],
                                    mat["filled_engraving_power" if filled else "engraving_power"],
                                    engraving_mode=mat.get("engraving_mode", "line"),
                                    filled_line_spacing=mat.get("filled_line_spacing", 0.15))
        assert open(new[0]).read() == open(want).read()
        assert layers_in(new[0]) == ["C03", "C02"], "felt keeps its engraved labels"
        assert "will not be engraved" not in [b for b in shown if b[0] == "showinfo"][0][2]
    check("other materials are untouched: the felt kerf test still engraves its labels, and nothing is said",
          other_materials_are_untouched)

    def speed_and_power_on_cork():
        app.fs_show_preview_var.set(False)
        app.fs_material_var.set("Cork")
        app.fs_engraving_var.set(True)
        app.fs_filename_var.set("sp")
        keep = app.fs_eng_speed_var.get(), app.fs_eng_power_var.get()
        try:
            app._apply_feeds_speeds_material_defaults()
            assert (app.fs_eng_speed_var.get(), app.fs_eng_power_var.get()) == ("0", "0"), \
                "cork pre-fills the engraving fields with its 0 and 0"
            assert app.fs_sweep_vars["speed"]["value"].get() == "1000" and app.fs_sweep_vars["power"]["value"].get() == "30"
            new, shown = click(app._on_generate_feeds_speeds_gcode)
            assert not [b for b in shown if b[0] == "showerror"], f"0 engraving must not block the test: {shown}"
            path = [p for p in new if p.endswith(".gcode")][0]
            names = layers_in(path)
            assert "C00" not in names and len(names) >= 4, f"no ID layer, every disc still cut: {names}"
            no_zero_feed(path)
            msg = [b for b in shown if b[0] == "showinfo"][0][2]
            assert "disc IDs will not be engraved" in msg and msg.endswith(note), msg
            n_discs = len(names)
            # Values typed into the Engraving fields bring the IDs back.
            app.fs_eng_speed_var.set("2500")
            app.fs_eng_power_var.set("9")
            new, shown = click(app._on_generate_feeds_speeds_gcode)
            path = [p for p in new if p.endswith(".gcode")][0]
            assert layers_in(path)[0] == "C00" and len(layers_in(path)) == n_discs + 1
            assert "will not be engraved" not in [b for b in shown if b[0] == "showinfo"][0][2]
            # IDs unticked: nothing was skipped, so nothing is said.
            app.fs_eng_speed_var.set("0")
            app.fs_eng_power_var.set("0")
            app.fs_engraving_var.set(False)
            new, shown = click(app._on_generate_feeds_speeds_gcode)
            assert "will not be engraved" not in [b for b in shown if b[0] == "showinfo"][0][2]
            app.fs_engraving_var.set(True)
            # Negative is still a typo.
            app.fs_eng_speed_var.set("-5")
            new, shown = click(app._on_generate_feeds_speeds_gcode)
            assert new == [] and [b for b in shown if b[0] == "showerror"], "a negative engraving speed is refused"
        finally:
            app.fs_eng_speed_var.set(keep[0])
            app.fs_eng_power_var.set(keep[1])
            app.fs_material_var.set("Felt")
    check("Speed & Power on cork: 0/0 pre-fills and does not block; discs cut, no IDs, no F0, and the Done box says so",
          speed_and_power_on_cork)

    def frame_and_cut_asks_before_it_runs():
        runs = []

        class Run:
            def __init__(self, parent, sender, lines, **kw):
                runs.append(list(lines))
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
        keep = ui_dialogs.FalconRunDialog, falcon_sender.FalconSender
        ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = Run, Sender
        app.falcon_port = "TEST"
        app.live_camera_var.set(False)
        app.kerf_material_var.set("Cork")
        app.fs_material_var.set("Cork")
        app.fs_show_preview_var.set(False)
        keep_fields = app.fs_eng_speed_var.get(), app.fs_eng_power_var.get()
        app.fs_eng_speed_var.set("0")
        app.fs_eng_power_var.set("0")
        app.crescent_material_var.set("Cork")
        app.crescent_list.clear()
        app.crescent_list.extend([{"tonehole_d": 43.0, "percent": 20.0, "curve": 75.0, "qty": 1},
                                  {"tonehole_d": 41.0, "percent": 10.0, "curve": 75.0, "qty": 1}])
        try:
            # (name, handler, the notice's words, the layer its engraving would have been on)
            for label, fn, words, eng_layer in (
                    ("kerf", app._on_frame_cut_kerf, "title and sizes will not be engraved", "C03"),
                    ("speed & power", app._on_frame_cut_feeds_speeds, "disc IDs will not be engraved", "C00"),
                    ("crescent list", app._on_frame_cut_crescent, "will not be numbered", "C03")):
                for ok in (False, True):
                    del runs[:], guard.asked[:], boxes[:]
                    before = len(saved)
                    guard.expect("Engraving Not Set", answer=ok, kind="askokcancel")
                    fn()
                    root.update()
                    asked = [(q[1], q[2]) for q in guard.asked]
                    assert len(asked) == 1 and asked[0][0] == "Engraving Not Set" and words in asked[0][1], (label, asked)
                    assert not [b for b in boxes if b[0] == "showerror"], boxes
                    if ok:
                        assert len(runs) == 3, f"{label}: OK goes on to jog, frame and cut"
                        assert not [ln for ln in runs[2] if re.search(r"F0(?![\d.])", ln)]
                        assert f"; Layer {eng_layer}" not in runs[2], f"{label}: the engraving layer was streamed"
                        assert [ln for ln in runs[2] if ln.startswith("; Layer ")], f"{label}: no cut layer streamed"
                    else:
                        assert runs == [], f"{label}: Cancel must run nothing"
                        assert len(saved) == before, f"{label}: Cancel came after the legend was asked for"
            # With real values nothing is asked.
            set_cork_engraving(3000, 10)
            del runs[:], guard.asked[:]
            app._on_frame_cut_kerf()
            root.update()
            assert guard.asked == [] and len(runs) == 3 and "; Layer C03" in runs[2]
        finally:
            set_cork_engraving(0, 0)
            ui_dialogs.FalconRunDialog, falcon_sender.FalconSender = keep
            app.falcon_port = None
            app.crescent_list.clear()
            app.fs_eng_speed_var.set(keep_fields[0])
            app.fs_eng_power_var.set(keep_fields[1])
            app.fs_material_var.set("Felt")
    check("Frame & Cut says what will not be engraved and asks before it runs; Cancel runs nothing",
          frame_and_cut_asks_before_it_runs)

    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
