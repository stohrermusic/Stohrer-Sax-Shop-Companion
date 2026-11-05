"""No test may touch the user's real settings folder (2026-11-05).

What happened: the app gives every G-code material a Default preset at
startup and saves at once when one is missing. For years that wrote
nothing for an existing user. The day a new material was added (adhesive
rings), every suite that built the app WITHOUT `SAXSHOP_CONFIG_DIR` started
writing that material's Default into the developer's real
%APPDATA%\\StohrerSaxShopCompanion\\gcode_presets.json -- with the values of
the moment, which were early placeholders. Found by an independent test
pass the same day. The same hunt found test_v161_compat overwriting the
real app_settings.json with old-version fixtures and restoring it from
memory afterwards: harmless when it finishes, a lost settings file when it
is killed half way.

Two guards, both checked here:

  1. Every suite that can reach state on disk sets `SAXSHOP_CONFIG_DIR` to a
     temp folder at module level, before any project module is imported
     (config's file paths are module constants, fixed at import). "Can
     reach" is decided from the suite's source: it imports the app or a
     module that owns saved state or builds windows, or it names one of
     config's file paths, writers or migrating loaders.
  2. tools/run_tests.py hands every suite its own temp `SAXSHOP_CONFIG_DIR`
     when the environment has none, so a full run (and CI) is safe even
     from a suite that slips past the scan.

And the proof, in a subprocess: the app built with the OS config location
pointed at a scratch "real" folder and no override writes exactly what the
first-launch bootstrap is designed to write there -- so the scan above is
guarding something real -- while the same construction under an override
leaves that folder byte for byte alone.

Headless except the last check (needs a display; self-skips).
"""
import ast
import hashlib
import json
import os
import site
import subprocess
import sys
import tempfile
import traceback

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(REPO, "tools")
sys.path.insert(0, REPO)
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-config-isolation-")

import config  # noqa: E402

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


# Every .py in the repo root is a project module; importing any of them
# imports config (directly or through the engines), which fixes the paths.
PROJECT_MODULES = {f[:-3] for f in os.listdir(REPO) if f.endswith(".py")}

# Importing one of these means the suite can build the app, a tab or a
# dialog, or load/save a library: all of them can reach the config folder.
STATEFUL_MODULES = {"main", "ui_dialogs", "tooling_tab", "toner_tab", "tuner_tab", "library_features"}

# Names that read-with-migration, write, or locate files in the config folder.
STATE_NAMES = {
    # config
    "save_settings", "save_presets", "load_presets", "append_job_history", "save_job_history",
    "setup_logging", "import_config_files", "ensure_config_dir", "migrate_config_files",
    "SETTINGS_FILE", "PAD_PRESET_FILE", "KEY_PRESET_FILE", "SCREW_SPECS_FILE", "TONER_DATA_FILE",
    "SIZING_PRESET_FILE", "GCODE_PRESET_FILE", "JOB_HISTORY_FILE",
    # toner_engine preset storage
    "save_tone_presets", "load_tone_presets",
    # camera_capture calibration file
    "save_calibration", "default_calibration_path",
}


def analyse(path):
    """(reasons it needs isolation, line of the override or None, line of the
    first module-level project import or None)."""
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename=path)
    reasons = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                if a.name.split(".")[0] in STATEFUL_MODULES:
                    reasons.add(f"imports {a.name}")
        elif isinstance(node, ast.ImportFrom):
            mod = (node.module or "").split(".")[0]
            if mod in STATEFUL_MODULES:
                reasons.add(f"imports {mod}")
            for a in node.names:
                if a.name in STATE_NAMES:
                    reasons.add(f"uses {a.name}")
        elif isinstance(node, ast.Attribute) and node.attr in STATE_NAMES:
            reasons.add(f"uses {node.attr}")
        elif isinstance(node, ast.Name) and node.id == "PadSVGGeneratorApp":
            reasons.add("builds the app")

    override_line = None
    first_import = None
    for node in tree.body:                       # module level only
        if override_line is None and _sets_override(node):
            override_line = node.lineno
        if first_import is None:
            names = []
            if isinstance(node, ast.Import):
                names = [a.name.split(".")[0] for a in node.names]
            elif isinstance(node, ast.ImportFrom):
                names = [(node.module or "").split(".")[0]]
            if any(n in PROJECT_MODULES for n in names):
                first_import = node.lineno
    return reasons, override_line, first_import


def _sets_override(node):
    """os.environ["SAXSHOP_CONFIG_DIR"] = ...  or  os.environ.setdefault("SAXSHOP_CONFIG_DIR", ...)."""
    def is_environ(n):
        return isinstance(n, ast.Attribute) and n.attr == "environ"
    if isinstance(node, ast.Assign):
        for t in node.targets:
            if (isinstance(t, ast.Subscript) and is_environ(t.value)
                    and isinstance(t.slice, ast.Constant) and t.slice.value == "SAXSHOP_CONFIG_DIR"):
                return True
    if isinstance(node, ast.Expr) and isinstance(node.value, ast.Call):
        fn = node.value.func
        if (isinstance(fn, ast.Attribute) and fn.attr == "setdefault" and is_environ(fn.value)
                and node.value.args and isinstance(node.value.args[0], ast.Constant)
                and node.value.args[0].value == "SAXSHOP_CONFIG_DIR"):
            return True
    return False


def suites():
    return sorted(f for f in os.listdir(TOOLS) if f.startswith("test_") and f.endswith(".py"))


def every_stateful_suite_isolates_itself():
    offenders, late, checked = [], [], 0
    for fname in suites():
        reasons, override_line, first_import = analyse(os.path.join(TOOLS, fname))
        if not reasons:
            continue
        checked += 1
        if override_line is None:
            offenders.append(f"{fname} ({', '.join(sorted(reasons))[:70]})")
        elif first_import is not None and first_import < override_line:
            late.append(f"{fname} (project import on line {first_import}, override on line {override_line})")
    assert checked >= 30, f"the scan only recognised {checked} stateful suites; its patterns have gone stale"
    assert not offenders, ("suites that can reach the config folder without SAXSHOP_CONFIG_DIR:\n    "
                           + "\n    ".join(offenders))
    assert not late, ("SAXSHOP_CONFIG_DIR is set AFTER a project import (too late, the paths are fixed):\n    "
                      + "\n    ".join(late))


def the_scan_recognises_the_patterns_that_bit():
    # The scan is only worth something if it would have flagged the suites
    # that actually wrote to the real folder. Feed it their old shapes.
    samples = {
        "builds the app, nested import (the six that wrote gcode_presets.json)": (
            "import os, sys\nsys.path.insert(0, '..')\n\n"
            "def main_test():\n    from main import PadSVGGeneratorApp\n    PadSVGGeneratorApp(None)\n"),
        "writes the settings path directly (test_v161_compat)": (
            "from config import load_settings, SETTINGS_FILE\nopen(SETTINGS_FILE, 'w')\n"),
        "saves through an alias": "import config as c\nc.save_settings({})\n",
        "override set too late": (
            "import os\nimport config\nos.environ['SAXSHOP_CONFIG_DIR'] = 'x'\nconfig.save_settings({})\n"),
    }
    with tempfile.TemporaryDirectory() as d:
        for label, src in samples.items():
            path = os.path.join(d, "sample.py")
            with open(path, "w", encoding="utf-8") as f:
                f.write(src)
            reasons, override_line, first_import = analyse(path)
            bad = bool(reasons) and (override_line is None
                                     or (first_import is not None and first_import < override_line))
            assert bad, f"the scan would have passed: {label}"
        good = ("import os, sys, tempfile\nos.environ['SAXSHOP_CONFIG_DIR'] = tempfile.mkdtemp()\n"
                "import main\nmain.PadSVGGeneratorApp(None)\n")
        path = os.path.join(d, "good.py")
        with open(path, "w", encoding="utf-8") as f:
            f.write(good)
        reasons, override_line, first_import = analyse(path)
        assert reasons and override_line is not None and override_line < first_import
        pure = "import svg_engine\nsvg_engine.get_disc_diameter(10, 'felt', {})\n"
        path = os.path.join(d, "pure.py")
        with open(path, "w", encoding="utf-8") as f:
            f.write(pure)
        assert not analyse(path)[0], "a pure engine suite should not be asked to isolate"


def the_runner_isolates_every_suite():
    import importlib.util
    spec = importlib.util.spec_from_file_location("run_tests", os.path.join(TOOLS, "run_tests.py"))
    rt = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(rt)
    base = {"PATH": "x", "APPDATA": "C:/real"}
    env, folder = rt.suite_env(base, "test_example")
    try:
        assert env["SAXSHOP_CONFIG_DIR"] == folder and os.path.isdir(folder), env.get("SAXSHOP_CONFIG_DIR")
        assert os.path.realpath(folder).startswith(os.path.realpath(tempfile.gettempdir()))
        assert "SAXSHOP_CONFIG_DIR" not in base, "the caller's environment must not be modified"
        env2, folder2 = rt.suite_env(base, "test_other")
        assert folder2 != folder, "each suite gets its own folder"
        os.rmdir(folder2)
    finally:
        os.rmdir(folder)
    # A developer who points the whole run at a prepared profile keeps it.
    env3, folder3 = rt.suite_env({"SAXSHOP_CONFIG_DIR": "D:/mine"}, "test_example")
    assert env3["SAXSHOP_CONFIG_DIR"] == "D:/mine" and folder3 is None


# ---- the proof: what the app writes to an unprotected "real" folder ------

def _snapshot(folder):
    out = {}
    for base, _dirs, files in os.walk(folder):
        for name in files:
            p = os.path.join(base, name)
            with open(p, "rb") as f:
                out[os.path.relpath(p, folder)] = (hashlib.sha256(f.read()).hexdigest(), os.path.getmtime(p))
    return out


def _redirected(env_key, root_dir, **extra):
    """This process's environment with the OS config location pointed at
    root_dir. On Windows the per-user site-packages also lives under
    %APPDATA%, so keep it importable."""
    env = {k: v for k, v in os.environ.items() if k != "SAXSHOP_CONFIG_DIR"}
    env["PYTHONUSERBASE"] = site.getuserbase()
    env[env_key] = root_dir
    env.update(extra)
    return env


def _fake_real_home(root_dir):
    """An existing user's folder from before today's newest material,
    reachable through the OS location on every platform."""
    if sys.platform == "win32":
        env_key, folder = "APPDATA", os.path.join(root_dir, config.APP_NAME)
    elif sys.platform == "darwin":
        env_key, folder = "HOME", os.path.join(root_dir, "Library", "Application Support", config.APP_NAME)
    else:
        env_key, folder = "XDG_CONFIG_HOME", os.path.join(root_dir, config.APP_NAME)
    os.makedirs(folder)
    s = json.loads(json.dumps(config.DEFAULT_SETTINGS))
    newest = config.GCODE_PRESET_MATERIALS[-1]
    presets = {m: {"Mine": s["gcode_settings"][m]} for m in config.GCODE_PRESET_MATERIALS if m != newest}
    with open(os.path.join(folder, "app_settings.json"), "w") as f:
        json.dump(s, f)
    with open(os.path.join(folder, "gcode_presets.json"), "w") as f:
        json.dump(presets, f)
    with open(os.path.join(folder, "sizing_presets.json"), "w") as f:
        json.dump({"Default": config.settings_to_sizing_preset(s)}, f)
    return env_key, folder, newest


_BUILD_APP = (
    "import sys, tkinter as tk\n"
    "sys.path.insert(0, r'%s')\n"
    "from tkinter import messagebox as mb\n"
    "for n in ('showinfo', 'showwarning', 'showerror'):\n"
    "    setattr(mb, n, lambda *a, **k: None)\n"
    "import main\n"
    "root = tk.Tk(); root.withdraw()\n"
    "main.PadSVGGeneratorApp(root)\n"
    "root.destroy()\n" % REPO.replace("\\", "\\\\")
)


def app_without_override_writes_only_the_bootstrap_and_with_it_writes_nothing():
    import tkinter as tk
    try:
        tk.Tk().destroy()
    except tk.TclError as e:
        print(f"        (skipped: no display: {e})")
        return
    with tempfile.TemporaryDirectory() as d:
        env_key, folder, newest = _fake_real_home(d)
        before = _snapshot(folder)

        # 1) Under an override the "real" folder is not touched at all.
        with tempfile.TemporaryDirectory() as override:
            env = _redirected(env_key, d, SAXSHOP_CONFIG_DIR=override)
            r = subprocess.run([sys.executable, "-c", _BUILD_APP], env=env, capture_output=True, text=True,
                               timeout=300)
            assert r.returncode == 0, r.stderr[-800:]
        assert _snapshot(folder) == before, "the app wrote to the OS config folder despite SAXSHOP_CONFIG_DIR"

        # 2) Without one, the designed first-launch bootstrap writes there:
        #    this is exactly what an un-isolated suite does to a developer.
        env = _redirected(env_key, d)
        r = subprocess.run([sys.executable, "-c", _BUILD_APP], env=env, capture_output=True, text=True,
                           timeout=300)
        assert r.returncode == 0, r.stderr[-800:]
        after = _snapshot(folder)
        # Files that were there before and are different now. (First launch
        # may also create files a folder lacks, such as the screw specs
        # library; those are new, not changed.)
        changed = sorted(k for k in before if before[k][0] != after.get(k, ("",))[0])
        assert changed == ["gcode_presets.json"], f"existing files the bootstrap changed: {changed}"
        with open(os.path.join(folder, "gcode_presets.json")) as f:
            now = json.load(f)
        assert newest in now and "Default" in now[newest], "the bootstrap should add the missing material"


def main():
    print("Config isolation: no test touches the real settings folder")
    print("=" * 60)
    check("every suite that can reach the config folder sets SAXSHOP_CONFIG_DIR before importing the app",
          every_stateful_suite_isolates_itself)
    check("the scan flags the shapes that wrote to the real folder, and passes a clean and a pure suite",
          the_scan_recognises_the_patterns_that_bit)
    check("tools/run_tests.py gives each suite its own temp SAXSHOP_CONFIG_DIR (and respects one already set)",
          the_runner_isolates_every_suite)
    check("the app writes nothing to the OS folder under an override, and only the bootstrap without one",
          app_without_override_writes_only_the_bootstrap_and_with_it_writes_nothing)
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
