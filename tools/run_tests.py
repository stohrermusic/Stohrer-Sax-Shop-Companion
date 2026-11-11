"""Run every test suite in tools/ and exit non-zero if any fails.

Each tools/test_*.py is a standalone script (not pytest), so each runs in
its own interpreter from the repo root. One line per suite, then the
failures' output in full. CI runs this on Windows, macOS and Linux (under
xvfb); locally it's the "run everything before a release" command.

    python tools/run_tests.py            # all suites
    python tools/run_tests.py zone svg   # only suites whose name contains a word
    python tools/run_tests.py --allow-missing numpy --allow-missing sounddevice
        # a suite that dies on "No module named 'numpy'" counts as skipped, not
        # failed — for the Intel Mac build, which ships without the audio stack
    python tools/run_tests.py --skip nesting_parity --skip polygon_parity
        # leave out suites by name (reported as skipped, never silently)

Skips test_descriptor_validity unless its WAV corpus is on this machine.

Every suite runs with its own temporary SAXSHOP_CONFIG_DIR (unless the
environment already sets one), so a full run can never read or write the
real settings folder of whoever runs it, whatever a suite forgets. Suites
that can reach the config folder also set the override themselves, for
direct runs; tools/test_config_isolation.py checks both.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(ROOT, "tools")
WAV_CORPUS = r"C:\sax shop companion\recordings"
# Per suite. The slowest suite takes about 135 s locally and 150 s on CI's
# Intel Mac; a suite past this is hung, not slow. SSC_SUITE_TIMEOUT_S
# overrides it (a slow machine, a profiling run). Until 2026-11-10 this was
# 900 s, and a hung suite cost a CI job fifteen minutes before it was named.
TIMEOUT_S = int(os.environ.get("SSC_SUITE_TIMEOUT_S", "300"))


def _popen_kwargs():
    """Start each suite in its own process group, so a hung one can be
    killed with every child it started (a --selftest child, a tour)."""
    if os.name == "nt":
        return {"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP}
    return {"start_new_session": True}


def _kill_tree(proc):
    """Kill the suite and everything it started. subprocess.run(timeout=)
    kills only the direct child and then waits for its pipes to close,
    which a grandchild holding stdout keeps open for ever: the Linux CI
    job of 2026-11-10 sat 33 minutes past a hung suite that way."""
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                           capture_output=True, timeout=30)
        else:
            import signal
            os.killpg(proc.pid, signal.SIGKILL)
    except Exception:  # noqa: BLE001  (the process may be gone already)
        pass
    try:
        proc.wait(timeout=15)
    except subprocess.TimeoutExpired:
        proc.kill()


SITECUSTOMIZE = '''\
# Written by tools/run_tests.py for each suite's interpreter (run_suite): arms
# faulthandler so a suite that hangs writes every thread's stack into its own
# output shortly before the runner kills it. Three intermittent CI hangs on
# 2026-11-10 left no trace of where they were; this is what makes the next
# one diagnosable.
import faulthandler
import os
import sys
try:
    faulthandler.enable()
    _s = float(os.environ.get("SSC_FAULTHANDLER_S", "0") or 0)
    if _s > 0:
        faulthandler.dump_traceback_later(_s, repeat=False, exit=False)
except Exception:
    pass
# Python imports only the first sitecustomize on sys.path: hand over to any
# other one further along (a developer's backstop on PYTHONPATH).
try:
    import importlib.util
    _here = os.path.dirname(os.path.abspath(__file__))
    for _p in sys.path:
        _cand = os.path.join(_p or os.getcwd(), "sitecustomize.py")
        if os.path.isfile(_cand) and os.path.dirname(os.path.abspath(_cand)) != _here:
            _spec = importlib.util.spec_from_file_location("sitecustomize_next", _cand)
            _spec.loader.exec_module(importlib.util.module_from_spec(_spec))
            break
except Exception:
    pass
'''
_FAULTHANDLER_SITE = None


def _faulthandler_site():
    """The folder holding the sitecustomize above, made once per run."""
    global _FAULTHANDLER_SITE
    if _FAULTHANDLER_SITE is None:
        d = tempfile.mkdtemp(prefix="ssc-faulthandler-")
        with open(os.path.join(d, "sitecustomize.py"), "w", encoding="utf-8") as f:
            f.write(SITECUSTOMIZE)
        _FAULTHANDLER_SITE = d
    return _FAULTHANDLER_SITE


def dump_at(timeout_s):
    """When the suite's own faulthandler dump fires: 15 s before the kill,
    or at 60 % of a short timeout."""
    return timeout_s - 15 if timeout_s > 30 else timeout_s * 0.6


def run_suite(cmd, cwd, env, timeout_s=None):
    """Run one suite to completion or the timeout. Returns (exit code, its
    output). The output goes to a file, not a pipe, so a timed-out suite's
    output so far is read back after it is killed, and no pipe can hold
    the runner once the process tree is gone. The suite's interpreter
    arms faulthandler (SITECUSTOMIZE on PYTHONPATH, SSC_FAULTHANDLER_S), so
    a hung suite's output ends with every thread's stack."""
    timeout_s = TIMEOUT_S if timeout_s is None else timeout_s
    env = dict(env)
    env["SSC_FAULTHANDLER_S"] = f"{dump_at(timeout_s):.1f}"
    site = _faulthandler_site()
    env["PYTHONPATH"] = site + (os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
    with tempfile.TemporaryFile() as log:
        proc = subprocess.Popen(cmd, cwd=cwd, env=env, stdout=log, stderr=subprocess.STDOUT,
                                **_popen_kwargs())
        try:
            code = proc.wait(timeout=timeout_s)
            timed_out = False
        except subprocess.TimeoutExpired:
            _kill_tree(proc)
            code, timed_out = 124, True
        log.seek(0)
        out = log.read().decode("utf-8", "replace")
    if timed_out:
        out = f"TIMEOUT after {timeout_s}s; killed. The suite's output so far:\n" + out
    return code, out


def summary_line(output):
    """The suite's own verdict line, or its last non-empty line."""
    lines = [ln.strip() for ln in output.splitlines() if ln.strip()]
    for ln in reversed(lines):
        low = ln.lower()
        if "passed" in low or "fail" in low or "skip" in low:
            return ln
    return lines[-1] if lines else ""


def suite_env(base_env, name):
    """(environment for one suite, its temp config folder or None).

    A fresh folder per suite, so one suite's leftovers can't steer the
    next. An override already in the environment is left alone: a developer
    pointing the whole run at a prepared profile means it.
    """
    env = dict(base_env)
    if env.get("SAXSHOP_CONFIG_DIR"):
        return env, None
    folder = tempfile.mkdtemp(prefix=f"ssc-run-{name.replace('test_', '', 1)}-")
    env["SAXSHOP_CONFIG_DIR"] = folder
    return env, folder


def main(argv):
    # On Windows, outside CI, the whole run moves to a private desktop first (Matt,
    # 2026-11-11: "I can't use my pc for a while when they are running"): this process
    # re-runs itself there and every suite inherits it. --visible keeps it here.
    import hidden_desktop
    if hidden_desktop.wanted(argv):
        return hidden_desktop.rerun_hidden([sys.executable] + list(argv))
    print(hidden_desktop.where(argv), flush=True)
    allow_missing = []
    skip_names = []
    args = [a for a in argv[1:] if a != "--visible"]
    while "--allow-missing" in args:
        i = args.index("--allow-missing")
        allow_missing.append(args[i + 1])
        del args[i:i + 2]
    while "--skip" in args:
        i = args.index("--skip")
        skip_names.append(args[i + 1])
        del args[i:i + 2]
    only = [a for a in args if not a.startswith("-")]
    suites = sorted(f for f in os.listdir(TOOLS) if f.startswith("test_") and f.endswith(".py"))
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    failed, skipped, ran = [], [], 0
    t_all = time.time()
    for fname in suites:
        name = fname[:-3]
        if only and not any(o in name for o in only):
            continue
        if name == "test_descriptor_validity" and not os.path.isdir(WAV_CORPUS):
            skipped.append(name)
            print(f"skip {name:34s}        WAV corpus not on this machine")
            continue
        if any(sk in name for sk in skip_names):
            skipped.append(name)
            print(f"skip {name:34s}        excluded by --skip")
            continue
        t0 = time.time()
        run_env, config_folder = suite_env(env, name)
        # Named before it starts, so a hung job's log says where it is.
        print(f"..   {name}", flush=True)
        try:
            code, out = run_suite([sys.executable, os.path.join(TOOLS, fname)], ROOT, run_env)
        finally:
            if config_folder:
                shutil.rmtree(config_folder, ignore_errors=True)
        if code == 124:
            # Say it now, with the tail of what the suite printed (its own
            # faulthandler dump of every thread is the end of it): the
            # summary at the end may never be reached if the job is killed.
            print(f"TIMEOUT {name}: killed after {TIMEOUT_S}s; its last lines:", flush=True)
            print("\n".join("    " + ln for ln in out.splitlines()[-80:]), flush=True)
        ran += 1
        missing = next((m for m in allow_missing if f"No module named '{m}'" in out), None)
        if code != 0 and missing:
            skipped.append(name)
            print(f"skip {name:34s} {time.time() - t0:5.1f}s  needs {missing} (not installed here, by design)", flush=True)
            continue
        status = "ok  " if code == 0 else "FAIL"
        print(f"{status} {name:34s} {time.time() - t0:5.1f}s  {summary_line(out)[:72]}", flush=True)
        if code != 0:
            failed.append((name, out))
    for name, out in failed:
        print("\n" + "=" * 70 + f"\nFAILED: {name} — output:\n" + "=" * 70)
        print("\n".join(out.splitlines()[-60:]))
    print("\n" + "-" * 70)
    print(f"{ran} suites run in {time.time() - t_all:.0f}s, {len(failed)} failed, {len(skipped)} skipped"
          + (f" ({', '.join(skipped)})" if skipped else ""))
    if failed:
        print("FAILED: " + ", ".join(n for n, _ in failed))
        return 1
    print("ALL SUITES PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
