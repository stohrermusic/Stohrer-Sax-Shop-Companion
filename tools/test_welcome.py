"""The Welcome window, the Quick Start Guide and the User Guide's search.

Matt, 2026-11-09: "each install shows a quick welcome, a version number,
and thank you for installing, be sure to search the user guide with any
questions, and says click here for the quick start guide", and of the
guide: "whatever guide we write should be SUPER concise". A first version
of the window was styled and animated; he took that back the same day:
"let's have the welcome screen be simple. if I have a detailed ui request
I will give it."

Pinned here:

  - the Welcome window shows once for each version: on a fresh settings
    folder, once more when the stored version differs, and not again on
    the next launch; the version is recorded when it is shown
  - it NEVER shows by itself except in a real launch: building the app
    does not show it, --selftest does not show it or record it, and the
    only call is in main.py's __main__ block
  - it is a plain dialog in the app's ordinary look and holds the app
    name, the version, the thank-you, two links and Close, and nothing
    else: no picture, no canvas, no animation, no timer, no grab
  - the links open the User Guide with the cursor in its search box, and
    the Quick Start Guide
  - the Quick Start Guide: its lines, the Frame & Cut line only on a build
    that can drive a laser, no scrolling, not modal
  - the User Guide search: finds a known word, counts the matches, wraps
    round, says when there are none, ignores case, and switches a one-tab
    guide to the full guide
  - Help has Welcome and Quick Start Guide above User Guide on every tab
  - in German (the longest language) both windows fit, measured in a
    child process

Self-skips without a display. Isolated profile (SAXSHOP_CONFIG_DIR).
"""
import json
import os
import subprocess
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
CHILD = sys.argv[2] if len(sys.argv) > 2 and sys.argv[1] == "--child" else None
_PROFILE = os.environ["SAXSHOP_CONFIG_DIR"] if CHILD else tempfile.mkdtemp(prefix="ssc-welcome-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE
if CHILD:
    with open(os.path.join(_PROFILE, "app_settings.json"), "w") as _f:
        json.dump({"language": CHILD}, _f)

from i18n import init_translation  # noqa: E402
init_translation(CHILD or "en")

import box_guard  # noqa: E402
box_guard.install()     # no native box from here on, and no question answered unasked (box_guard.py)

import config  # noqa: E402

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


def child(lang):
    """Both windows built in `lang`; their sizes, as JSON."""
    import tkinter as tk
    import tkinter.font as tkfont
    import main as main_mod
    import ui_dialogs
    root = tk.Tk()
    root.withdraw()
    app = main_mod.PadSVGGeneratorApp(root)
    win = app.open_welcome()
    win.update_idletasks()
    labels = [win.name_label, win.version_label, win.thanks_label, win.guide_link, win.quick_start_link]
    links = [str(w.cget("text")) for w in (win.guide_link, win.quick_start_link)]
    widest = max(tkfont.Font(font=lab.cget("font")).measure(str(lab.cget("text"))) for lab in labels)
    quick = ui_dialogs.QuickStartWindow(root, show_frame_cut=True)
    quick.update_idletasks()
    print("REPORT " + json.dumps({
        "language_shown": str(win.close_button.cget("text")),
        "welcome": [win.winfo_reqwidth(), win.winfo_reqheight()],
        "widest_line": widest, "wrap": win.WRAP, "links": links,
        "quick": [quick.winfo_reqwidth(), quick.winfo_reqheight()],
        "quick_lines": len(quick.line_labels),
        "screen": [root.winfo_screenwidth(), root.winfo_screenheight()],
    }))
    root.destroy()


if CHILD:
    child(CHILD)
    sys.exit(0)


def gui():
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return
    root.withdraw()
    import inspect
    import main as main_mod
    import ui_dialogs
    app = main_mod.PadSVGGeneratorApp(root)
    version = config.APP_VERSION
    callback_errors = []
    root.report_callback_exception = lambda *a: callback_errors.append(a)

    def welcomes():
        return [w for w in root.winfo_children() if isinstance(w, ui_dialogs.WelcomeWindow) and w.winfo_exists()]

    def settings_on_disk():
        if not os.path.exists(config.SETTINGS_FILE):
            return {}
        with open(config.SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)

    def building_the_app_shows_nothing():
        assert config.DEFAULT_SETTINGS["welcomed_version"] == "", "a new install has not been welcomed"
        assert app.settings["welcomed_version"] == ""
        assert welcomes() == [], "building the app opened a Welcome window"
        assert settings_on_disk().get("welcomed_version", "") == ""
        src = inspect.getsource(main_mod)
        before_main, in_main = src.split("if __name__ == '__main__':")
        calls = before_main.count("maybe_show_welcome(")
        assert calls == 1, f"maybe_show_welcome should only be defined outside the launch block, found {calls} uses"
        assert in_main.count("app.maybe_show_welcome") == 1 and "if not tour:" in in_main.split("app.maybe_show_welcome")[0][-200:], \
            "the one call belongs in the real launch path, and not under --tour"
        init_src = inspect.getsource(main_mod.PadSVGGeneratorApp.__init__)
        assert "welcome" not in init_src.lower(), "the constructor must not show or schedule the Welcome window"
        assert "maybe_show_welcome" not in inspect.getsource(main_mod._selftest)
    check("building the app never shows the Welcome window; the only call is in the real launch path, not under --tour",
          building_the_app_shows_nothing)

    def once_per_version():
        win = app.maybe_show_welcome()
        assert isinstance(win, ui_dialogs.WelcomeWindow) and welcomes() == [win], "a fresh folder is welcomed"
        assert app.settings["welcomed_version"] == version
        assert settings_on_disk().get("welcomed_version") == version, "recorded at once, not when it is closed"
        assert app.maybe_show_welcome() is None, "not again"
        win.destroy()
        assert app.maybe_show_welcome() is None and welcomes() == [], "not on the next launch either"
        # An update: the stored version is some other one.
        older = "1.0" if version != "1.0" else "0.9"
        app.settings["welcomed_version"] = older
        win = app.maybe_show_welcome()
        assert isinstance(win, ui_dialogs.WelcomeWindow) and app.settings["welcomed_version"] == version
        assert app.maybe_show_welcome() is None
        win.destroy()
        # Help > Welcome shows it any time, one at a time, and records nothing new.
        a = app.open_welcome()
        assert app.open_welcome() is a and len(welcomes()) == 1, "a second request reuses the open window"
        a.destroy()
    check("shown once on a fresh folder, once more when the stored version differs, never again after that",
          once_per_version)

    def pending_ids():
        return set(str(x) for x in root.tk.splitlist(root.tk.call("after", "info")))

    def what_the_window_holds():
        before = pending_ids()
        win = ui_dialogs.WelcomeWindow(root, "9.87")
        try:
            children = win.winfo_children()
            labels = [w for w in children if isinstance(w, tk.Label)]
            assert [str(w.cget("text")) for w in labels] == [
                "Stohrer Sax Shop Companion", "Version 9.87", "Thank you for installing.",
                "Questions? Search the User Guide.", "Click here for the quick start guide"], \
                [str(w.cget("text")) for w in labels]
            buttons = [w for w in children if isinstance(w, tk.Button)]
            assert [str(b.cget("text")) for b in buttons] == ["Close"]
            assert len(children) == len(labels) + 1, \
                f"the window holds five lines and Close, nothing else: {[type(w).__name__ for w in children]}"
            # The app's ordinary dialog look, nothing of its own.
            assert str(win.cget("bg")) == str(ui_dialogs.DIALOG_BG)
            for w in labels:
                assert str(w.cget("bg")) == str(ui_dialogs.DIALOG_BG)
            assert str(win.guide_link.cget("fg")).lower() == "#0066cc" and "underline" in str(win.guide_link.cget("font"))
            assert str(win.guide_link.cget("cursor")) == "hand2" and str(win.quick_start_link.cget("cursor")) == "hand2"
            # Nothing moves and nothing waits: no canvas, no timer, no grab.
            everything = []

            def walk(w):
                for c in w.winfo_children():
                    everything.append(c)
                    walk(c)
            walk(win)
            assert not [w for w in everything if isinstance(w, (tk.Canvas, tk.Scrollbar, tk.Text))]
            root.update()
            assert pending_ids() <= before, f"the Welcome window scheduled something: {pending_ids() - before}"
            assert win.grab_current() is not win, "the Welcome window must not be modal"
            src = inspect.getsource(ui_dialogs.WelcomeWindow)
            code = "\n".join(ln for ln in src.splitlines() if not ln.strip().startswith(("#", '"')))
            for word in ("after(", "grab_set", "focus_force", "Canvas", "create_", "wait_window"):
                assert word not in code.split('"""')[-1], f"a plain dialog has no {word}"
            for gone in ("advance", "skip_animation", "lit", "revealed", "pending", "canvas"):
                assert not hasattr(win, gone), f"left over from the animated version: {gone}"
            assert tuple(win.resizable()) in ((False, False), (0, 0))
        finally:
            win.destroy()
        assert not callback_errors, callback_errors
    check("the window is a plain dialog: name, version, thank-you, two links and Close; no picture, no timer, no grab",
          what_the_window_holds)

    def the_two_links():
        opened = []
        win = ui_dialogs.WelcomeWindow(root, version,
                                       on_search_guide=lambda: opened.append("guide"),
                                       on_quick_start=lambda: opened.append("quick"))
        try:
            assert win.guide_link.bind("<Button-1>") and win.quick_start_link.bind("<Button-1>")
            # Click each link itself (the event goes to the label's own
            # binding), so a link wired to the other window is caught.
            # An event only reaches a mapped window, and a transient of this
            # test's withdrawn root is not mapped: untie it and map it far
            # off screen for the two clicks.
            win.tk.call("wm", "transient", win._w, "")
            win.geometry("+-4000+-4000")
            win.deiconify()
            win.update()
            win.guide_link.event_generate("<Button-1>", x=2, y=2)
            root.update()
            assert opened == ["guide"], f"the User Guide link opened {opened}"
            win.quick_start_link.event_generate("<Button-1>", x=2, y=2)
            root.update()
            assert opened == ["guide", "quick"], f"the quick start link opened {opened[1:]}"
        finally:
            win.destroy()
        # The app's own two targets.
        focused = []
        real = ui_dialogs.UserGuideWindow.focus_search
        ui_dialogs.UserGuideWindow.focus_search = lambda self: focused.append(self) or real(self)
        try:
            guide = app.open_user_guide_search()
        finally:
            ui_dialogs.UserGuideWindow.focus_search = real
        try:
            assert isinstance(guide, ui_dialogs.UserGuideWindow) and guide.opened_at is None, "the whole guide"
            assert focused == [guide], "the cursor goes to the search box"
        finally:
            guide.destroy()
        quick = app.open_quick_start()
        try:
            assert isinstance(quick, ui_dialogs.QuickStartWindow)
        finally:
            quick.destroy()
        win = app.open_welcome()
        try:
            assert win._on_search_guide == app.open_user_guide_search and win._on_quick_start == app.open_quick_start
        finally:
            win.destroy()
    check("the links open the full User Guide with the cursor in its search box, and the Quick Start Guide",
          the_two_links)

    def the_quick_start_guide():
        full = [t for _k, t in ui_dialogs.QuickStartWindow.lines(True)]
        short = [t for _k, t in ui_dialogs.QuickStartWindow.lines(False)]
        assert full == [
            "Three ways to cut",
            "SD card: Generate G-code, save it to the card, run it on the laser.",
            "LightBurn: Generate SVG, open it in LightBurn, set speed and power for each color.",
            "Frame & Cut: laser on USB, jog to the corner, frame, cut. Switch it on in File > Feature Set.",
            "Do not mix them",
            "G-code is not for LightBurn. SVG is not for the SD card.",
            "Before you cut",
            "Speeds and powers are in Options > G-code Settings (Pad Maker) and Options > Settings (Tooling).",
            "Test on the material you will really cut.",
            "Die inserts and holders use the Acrylic settings.",
            "The rest is in Help > User Guide.",
        ], full
        assert short == [t for t in full if not t.startswith("Frame & Cut")], \
            "a build that cannot drive a laser does not get the Frame & Cut line"
        assert [k for k, _t in ui_dialogs.QuickStartWindow.lines(True)].count("h") == 3
        win = ui_dialogs.QuickStartWindow(root, show_frame_cut=False)
        try:
            assert [str(lab.cget("text")) for lab in win.line_labels] == short
            everything = []

            def walk(w):
                for c in w.winfo_children():
                    everything.append(c)
                    walk(c)
            walk(win)
            assert not [w for w in everything if isinstance(w, (tk.Scrollbar, tk.Text, tk.Canvas))], \
                "plain labels: nothing to scroll, no pictures"
            assert [str(w.cget("text")) for w in everything if isinstance(w, tk.Button)] == ["Close"]
            assert win.grab_current() is not win, "not modal"
        finally:
            win.destroy()
        keep = main_mod.machine_supported
        try:
            for supported in (True, False):
                main_mod.machine_supported = lambda s=supported: s
                q = app.open_quick_start()
                has = any(str(lab.cget("text")).startswith("Frame & Cut") for lab in q.line_labels)
                q.destroy()
                assert has is supported
        finally:
            main_mod.machine_supported = keep
    check("Quick Start Guide: its eleven lines (ten without Frame & Cut on a build that cannot drive a laser), no scrolling",
          the_quick_start_guide)

    def help_menu_on_every_tab():
        for menu in (app.pad_menu, app.key_menu, app.screw_menu, app.serial_menu,
                     app.tooling_menu, app.tuner_menu, app.toner_menu):
            helps = [i for i in range(menu.index("end") + 1)
                     if menu.type(i) == "cascade" and menu.entrycget(i, "label") == "Help"]
            assert len(helps) == 1
            sub = menu.nametowidget(menu.entrycget(helps[0], "menu"))
            labels = [sub.entrycget(i, "label") for i in range(sub.index("end") + 1) if sub.type(i) == "command"]
            assert labels[:3] == ["Welcome", "Quick Start Guide", "User Guide..."], labels
    check("Help on every tab: Welcome and Quick Start Guide, above User Guide", help_menu_on_every_tab)

    def searching_the_user_guide():
        # Opened at the Tuner's section: the whole guide is there all the
        # same (2026-11-09, the guide upgrade: no one-tab guide any more,
        # no "Show Full Guide" button), so a search finds the Tooling text.
        guide = ui_dialogs.UserGuideWindow(root, section="tuner")
        try:
            text = lambda: guide.text.get("1.0", "end")      # noqa: E731
            assert guide.opened_at == "tuner" and not hasattr(guide, "_show_all_btn")
            assert guide.text.index("@0,0").split(".")[0] == guide.text.index("sec_tuner").split(".")[0], \
                "opened at the Tuner section: its heading at the top"
            guide._search_var.set("crescent")
            total = text().lower().count("crescent")
            assert total >= 3, f"fixture: the full guide should mention crescents several times ({total})"
            assert len(guide._search_hits) == total
            assert guide._search_status_var.get() == f"1 of {total}"
            assert len(guide.text.tag_ranges("search_hit")) == 2 * total, "every match is highlighted"
            cur = guide.text.tag_ranges("search_current")
            assert len(cur) == 2 and str(cur[0]) == guide.text.index(guide._search_hits[0][0])
            assert guide.text.get(*cur).lower() == "crescent"
            # Next, and round again.
            seen = [str(guide.text.tag_ranges("search_current")[0])]
            for n in range(2, total + 1):
                guide.search_next()
                assert guide._search_status_var.get() == f"{n} of {total}"
                seen.append(str(guide.text.tag_ranges("search_current")[0]))
            assert len(set(seen)) == total, "Next should visit each match once"
            guide.search_next()
            assert guide._search_status_var.get() == f"1 of {total}", "after the last match it wraps to the first"
            assert str(guide.text.tag_ranges("search_current")[0]) == seen[0]
            # Case does not matter.
            guide._search_var.set("CRESCENT")
            assert len(guide._search_hits) == total
            # Nothing found.
            guide._search_var.set("zzqqxx")
            assert guide._search_status_var.get() == "No matches" and guide._search_hits == []
            assert not guide.text.tag_ranges("search_hit") and not guide.text.tag_ranges("search_current")
            guide.search_next()
            assert guide._search_status_var.get() == "No matches"
            # Cleared.
            guide._search_var.set("")
            assert guide._search_status_var.get() == "" and not guide.text.tag_ranges("search_hit")
            # A phrase, and the things people will look for.
            for phrase in ("Frame & Cut", "kerf", "G-code or SVG?"):
                guide._search_var.set(phrase)
                assert len(guide._search_hits) == text().lower().count(phrase.lower()) >= 1, phrase
            key = "<Command-f>" if sys.platform == "darwin" else "<Control-f>"
            assert guide.bind(key), f"{key} puts the cursor in the search box"
            assert guide.search_entry.bind("<Return>"), "Enter goes to the next match"
            assert str(guide._search_next_btn.cget("text")) == "Next"
            assert str(guide.text.cget("state")) == "disabled", "searching must not make the guide editable"
        finally:
            guide.destroy()
        # The full guide opened as such searches in place.
        guide = ui_dialogs.UserGuideWindow(root)
        try:
            guide._search_var.set("Welcome")
            assert guide.opened_at is None and guide._search_status_var.get() in ("No matches",) or guide._search_hits
        finally:
            guide.destroy()
    check("User Guide search: finds a word, counts and highlights every match, wraps, ignores case, covers the "
          "whole guide whatever section it opened at", searching_the_user_guide)

    def selftest_shows_and_records_nothing():
        folder = tempfile.mkdtemp(prefix="ssc-welcome-selftest-")
        env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
        p = subprocess.run([sys.executable, os.path.join(REPO, "main.py"), "--selftest"],
                           capture_output=True, text=True, env=env, timeout=300, cwd=REPO)
        assert p.returncode == 0 and "selftest ok" in (p.stdout + p.stderr), (p.returncode, p.stdout[-300:], p.stderr[-600:])
        path = os.path.join(folder, os.path.basename(config.SETTINGS_FILE))
        stored = json.load(open(path, encoding="utf-8")).get("welcomed_version", "") if os.path.exists(path) else ""
        assert stored == "", "--selftest recorded a welcome: the next real launch would not show it"
    check("--selftest shows no Welcome window and records none, so the next real launch still welcomes",
          selftest_shows_and_records_nothing)

    def both_windows_fit_in_german():
        folder = tempfile.mkdtemp(prefix="ssc-welcome-de-")
        env = dict(os.environ, SAXSHOP_CONFIG_DIR=folder, PYTHONIOENCODING="utf-8")
        p = subprocess.run([sys.executable, os.path.abspath(__file__), "--child", "de"],
                           capture_output=True, text=True, env=env, timeout=300, cwd=REPO, encoding="utf-8")
        line = [ln for ln in (p.stdout or "").splitlines() if ln.startswith("REPORT ")]
        assert p.returncode == 0 and line, (p.returncode, (p.stdout or "")[-300:], (p.stderr or "")[-800:])
        r = json.loads(line[0][7:])
        assert r["language_shown"] == "Schließen", f"fixture: the child should be in German, got {r['language_shown']!r}"
        assert r["links"] == ["Fragen? Durchsuchen Sie das Benutzerhandbuch.", "Hier klicken für den Schnellstart"], r["links"]
        w, h = r["welcome"]
        assert w <= 480 and h <= 300, f"the German Welcome window is {w} x {h}: it should be small"
        assert w <= r["screen"][0] and h <= r["screen"][1]
        assert r["widest_line"] <= r["wrap"] + 120, "a welcome line is far wider than its wrap width"
        qw, qh = r["quick"]
        assert r["quick_lines"] == 11 and qw <= 560 and qh <= 620, f"the German Quick Start Guide is {qw} x {qh}"
        assert qh <= r["screen"][1] - 80, "the German Quick Start Guide would need scrolling on this screen"
    check("in German both windows fit: the Welcome window stays small and the guide needs no scrolling",
          both_windows_fit_in_german)

    try:
        root.destroy()
    except tk.TclError:
        pass


gui()
print("=" * 60)
print(f"{sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
