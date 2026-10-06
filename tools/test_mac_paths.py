"""The macOS code paths, and the dialog tour as a gate.

Nobody on the project owns a Mac, so the macOS CI runner is where these
run for real; on Windows and Linux the same checks pin the non-Mac side.

  - IS_MACOS agrees across modules with sys.platform
  - the quit path: WM_DELETE_WINDOW everywhere, plus ::tk::mac::Quit on
    darwin (Cmd-Q used to skip on_exit and never save settings)
  - theming: system colors on darwin (no cream palette anywhere in a
    dialog), the cream theme elsewhere
  - every visible tab label fits in the window (Aqua clips, it doesn't
    grow; seven tabs at 640 px read "Pad Mak / Toner (bet" on the first
    Mac screenshot) — measured with the notebook's own font
  - the GPU renderer is never active on darwin
  - the default config folder lands under Library/Application Support
  - a hidden tab can be re-added and selected
  - the full --tour walk constructs and tears down every tab and dialog
    on the live app without raising (screenshots off)

Uses an isolated profile (SAXSHOP_CONFIG_DIR) with the toner unlocked, so
all seven tabs exist and nothing touches the user's settings.
Self-skips without a display.
"""
import json
import os
import sys
import tempfile
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-mac-test-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
with open(os.path.join(_PROFILE, "app_settings.json"), "w", encoding="utf-8") as _f:
    json.dump({"toner_unlocked": True, "visible_tabs": {"Toner": True}}, _f)

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []
IS_MAC = sys.platform == "darwin"
CREAM_PALETTE = {"#F0EAD6", "#FFFDD0", "#E0F7FA", "#E8F5E9"}


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
    from tkinter import ttk
    import tkinter.font as tkfont
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0

    print(f"macOS paths and dialog tour  (platform: {sys.platform})")
    print("=" * 60)
    import config
    import main as main_mod
    import ui_dialogs
    import tuner_tab

    check("IS_MACOS agrees across main / ui_dialogs / sys.platform",
          lambda: main_mod.IS_MACOS == ui_dialogs.IS_MACOS == IS_MAC)

    def default_config_dir():
        saved = os.environ.pop("SAXSHOP_CONFIG_DIR")
        try:
            d = config.get_config_dir().replace("\\", "/")
        finally:
            os.environ["SAXSHOP_CONFIG_DIR"] = saved
        if IS_MAC:
            assert d.endswith("Library/Application Support/StohrerSaxShopCompanion"), d
        elif sys.platform == "win32":
            assert d.endswith("StohrerSaxShopCompanion") and "AppData" in d or os.environ.get("APPDATA", "") in d, d
        else:
            assert d.endswith(".config/StohrerSaxShopCompanion") or "StohrerSaxShopCompanion" in d, d
    check("default config folder is the platform's", default_config_dir)

    root.geometry("+20+40")
    app = main_mod.PadSVGGeneratorApp(root)
    root.update()           # runs the after_idle tab fit
    root.update_idletasks()

    check("WM_DELETE_WINDOW routes to on_exit",
          lambda: bool(root.protocol("WM_DELETE_WINDOW")))
    if IS_MAC:
        def quit_command_registered():
            # tk.call can hand back a Tcl object rather than a str; compare text.
            found = str(root.tk.call("info", "commands", "::tk::mac::Quit"))
            assert "::tk::mac::Quit" in found, f"info commands returned {found!r}"
        check("::tk::mac::Quit is registered (Cmd-Q saves settings)", quit_command_registered)
        check("DIALOG_BG is the system window color",
              lambda: ui_dialogs.DIALOG_BG == "systemWindowBackgroundColor")
        check("GPU renderer is never active on darwin",
              lambda: not tuner_tab._HAS_GPU_RENDERER and not app._tuner_use_gpu)
    else:
        check("::tk::mac::Quit is not registered off-Mac",
              lambda: str(root.tk.call("info", "commands", "::tk::mac::Quit")) == "")
        check("cream theme applied to the root window",
              lambda: root.cget("bg").upper() == "#FFFDD0")

    def tab_labels_fit():
        style = ttk.Style(root)
        spec = style.lookup("TNotebook.Tab", "font") or "TkDefaultFont"
        font = tkfont.nametofont(spec) if spec in tkfont.names(root) else tkfont.Font(root, font=spec)
        visible = [t for t in app.notebook.tabs() if app.notebook.tab(t, "state") != "hidden"]
        assert len(visible) == 7, f"{len(visible)} visible tabs"
        need = sum(font.measure(app.notebook.tab(t, "text")) for t in visible)
        need += main_mod.TAB_LABEL_PADDING_PX * len(visible) + 40
        width = root.winfo_width()
        assert width >= need, f"window {width} px, labels need {need} px"
    check("all seven tab labels fit in the window", tab_labels_fit)

    def sizing_rules_fits_its_content():
        # Aqua clipped the preset bar at the dialog's fixed 500 px (Mac tour
        # screenshot, 2026-10-06). The content's widest row must fit the canvas.
        opts = ui_dialogs.OptionsWindow(root, app, app.settings, lambda: None, lambda: None,
                                        sizing_presets={}, sizing_presets_save_callback=lambda: None)
        root.update()                # runs the after_idle fit (and <Map>)
        opts.top.update()            # lets the new geometry lay out
        need = opts.scrollable_frame.winfo_reqwidth()
        have = opts.canvas.winfo_width()
        opts.top.destroy()
        assert have >= need, f"Sizing Rules content needs {need} px, canvas is {have} px"
    check("Sizing Rules dialog is wide enough for its content", sizing_rules_fits_its_content)

    def dialogs_use_platform_colors():
        dialogs = [
            ui_dialogs.OptionsWindow(root, app, app.settings, lambda: None, lambda: None,
                                     sizing_presets={}, sizing_presets_save_callback=lambda: None),
            ui_dialogs.LayerColorWindow(root, app.settings, lambda: None),
            ui_dialogs.KeyLayoutWindow(root, app.settings, lambda: None, lambda: None),
            ui_dialogs.GcodeSettingsWindow(root, app.settings, lambda s: None),
        ]
        offenders = []

        def walk(w):
            try:
                bg = str(w.cget("bg")).upper()
            except tk.TclError:
                bg = ""
            if IS_MAC and bg in CREAM_PALETTE:
                offenders.append(f"{type(w).__name__}:{bg}")
            for c in w.winfo_children():
                walk(c)
        for d in dialogs:
            d.top.withdraw()
            walk(d.top)
            d.top.destroy()
        assert not offenders, f"cream colours on darwin: {offenders[:6]}"
    check("settings dialogs use the platform's colours" if IS_MAC else "settings dialogs construct", dialogs_use_platform_colors)

    def hidden_tab_readd():
        app.notebook.hide(app.toner_tab_frame)
        assert app.notebook.tab(app.toner_tab_frame, "state") == "hidden"
        app.notebook.add(app.toner_tab_frame)
        app.notebook.select(app.toner_tab_frame)
        root.update_idletasks()
        assert app.notebook.select() == str(app.toner_tab_frame)
        app.notebook.select(app.pad_tab)
    check("a hidden tab can be re-added and selected", hidden_tab_readd)

    # ---- the full tour as a gate (no screenshots) --------------------
    app._toner_mic_checked = True
    log, errors = main_mod.run_tour(root, app, shots_dir=None, step_ms=350, on_done=root.quit)
    root.after(120000, root.quit)   # safety net
    root.mainloop()
    expected = ["pad-maker", "sizing-rules-with-preview", "layer-colors", "gcode-settings", "nesting-preview",
                "polygon-draw", "job-history", "feature-set", "user-guide", "about", "key-heights",
                "key-layout", "serial-lookup", "screw-specs", "tooling", "tuner-no-mic", "tuner", "toner"]
    check(f"tour visited every stop ({len(log)}/{len(expected)})", lambda: log == expected)
    def tour_raised_nothing():
        assert not errors, "\n".join(errors)
    check("tour raised nothing", tour_raised_nothing)
    check("no stray Toplevel left open after the tour",
          lambda: not [w for w in root.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_exists()])
    check("tuner and toner stopped after the tour",
          lambda: not app._tuner_running and not app._toner_engine.is_running)

    try:
        root.destroy()
    except tk.TclError:
        pass
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
