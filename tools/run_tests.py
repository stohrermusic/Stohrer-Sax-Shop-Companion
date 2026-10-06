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

Skips test_descriptor_validity unless its WAV corpus is on this machine.
"""
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOOLS = os.path.join(ROOT, "tools")
WAV_CORPUS = r"C:\sax shop companion\recordings"
TIMEOUT_S = 900


def summary_line(output):
    """The suite's own verdict line, or its last non-empty line."""
    lines = [ln.strip() for ln in output.splitlines() if ln.strip()]
    for ln in reversed(lines):
        low = ln.lower()
        if "passed" in low or "fail" in low or "skip" in low:
            return ln
    return lines[-1] if lines else ""


def main(argv):
    allow_missing = []
    args = list(argv[1:])
    while "--allow-missing" in args:
        i = args.index("--allow-missing")
        allow_missing.append(args[i + 1])
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
        t0 = time.time()
        try:
            r = subprocess.run([sys.executable, os.path.join(TOOLS, fname)], cwd=ROOT, env=env,
                               capture_output=True, text=True, encoding="utf-8", errors="replace",
                               timeout=TIMEOUT_S)
            code, out = r.returncode, (r.stdout or "") + (r.stderr or "")
        except subprocess.TimeoutExpired as e:
            code, out = 124, f"TIMEOUT after {TIMEOUT_S}s\n" + str(e.stdout or "")
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
