"""Run the test suites the way CI runs them: Python 3.11 and the hashed lock, not this machine's
interpreter and whatever it has installed.

Why (2026-11-11): five Analyzer commits in a row passed every suite here (Python 3.14, an unlocked numpy 2.4.2)
and failed on all four CI platforms (Python 3.11, numpy 2.4.6 from locks/full.txt): a float one
last place apart from a golden literal, and nothing local to show it. This script builds a virtual
environment with the interpreter CI uses, installs `locks/full.txt` with `--require-hashes` exactly
as the workflow does, and runs `tools/run_tests.py` in it (filters pass through).

    python tools/run_tests_as_ci.py                 # every suite, Python 3.11
    python tools/run_tests_as_ci.py toner_data      # the suites whose name holds "toner_data"
    python tools/run_tests_as_ci.py --python 3.12   # another interpreter
    python tools/run_tests_as_ci.py --fresh         # rebuild the environment first

What it needs: a Python 3.11 on this machine. It looks, in order, for the Windows launcher's
`py -3.11`, a working `python3.11` on PATH, and uv's managed interpreters (`uv python find 3.11`);
with `--install` and uv available it fetches one (`uv python install 3.11`). To get uv:
`python -m pip install --user uv` (puts uv.exe in the user Scripts folder, which this script also
checks). The environment lives in `<repo>/.venv/ci-py3.11` (git ignores `.venv`) and is reused;
the lock is reinstalled only when `locks/full.txt` changed since. The suites' own rules apply
unchanged: run_tests.py gives each suite its own SAXSHOP_CONFIG_DIR.

The first lines it prints name the interpreter the run used: put them in the report.
"""
import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import sysconfig

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCK = os.path.join(REPO, "locks", "full.txt")
PROBE = ["-c", "import sys, venv; print(sys.executable)"]


def find_uv():
    cands = [shutil.which("uv")]
    if os.name == "nt":
        cands.append(os.path.join(sysconfig.get_path("scripts", "nt_user"), "uv.exe"))
    cands.append(os.path.join(sysconfig.get_path("scripts"), "uv.exe" if os.name == "nt" else "uv"))
    for cand in cands:
        if cand and os.path.isfile(cand):
            return cand
    return None


def works(cmd):
    """The interpreter `cmd` (a list) runs and has venv: its sys.executable, else None."""
    try:
        r = subprocess.run(cmd + PROBE, capture_output=True, text=True, timeout=60)
    except (OSError, subprocess.SubprocessError):
        return None
    out = r.stdout.strip()
    return out if r.returncode == 0 and out and os.path.isfile(out) else None


def find_python(version, install):
    """A python of `version` ("3.11"): the launcher, PATH, uv's managed ones, or a fresh uv install."""
    if os.name == "nt" and shutil.which("py"):
        found = works(["py", f"-{version}"])
        if found:
            return found
    on_path = shutil.which(f"python{version}")
    if on_path:
        found = works([on_path])
        if found:
            return found
    uv = find_uv()
    if uv:
        for attempt in range(2):
            r = subprocess.run([uv, "python", "find", version], capture_output=True, text=True)
            cand = r.stdout.strip()
            if r.returncode == 0 and cand and os.path.isfile(cand):
                return cand
            if attempt == 0 and install:
                subprocess.run([uv, "python", "install", version], check=False)
            else:
                break
    return None


def make_venv(base, venv):
    r = subprocess.run([base, "-m", "venv", venv], capture_output=True, text=True)
    if r.returncode == 0:
        return "pip"
    uv = find_uv()
    if uv:
        # a uv-managed interpreter may refuse `-m venv` through its shim; uv itself makes the venv
        r2 = subprocess.run([uv, "venv", "--python", base, venv], capture_output=True, text=True)
        if r2.returncode == 0:
            return "uv"
        print(r2.stderr.strip(), file=sys.stderr)
    print(r.stderr.strip(), file=sys.stderr)
    raise SystemExit(f"could not create {venv} from {base}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0], formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--python", default="3.11", help="interpreter version CI uses (default 3.11)")
    ap.add_argument("--fresh", action="store_true", help="rebuild the environment")
    ap.add_argument("--install", action="store_true", help="fetch the interpreter with uv if none is found")
    ap.add_argument("filters", nargs="*", help="suite name filters, as for run_tests.py")
    args = ap.parse_args()

    base = find_python(args.python, args.install)
    if not base:
        print(f"no Python {args.python} found: install one (py launcher, PATH, or `python -m pip install --user uv` "
              f"then rerun with --install)", file=sys.stderr)
        return 2
    venv = os.path.join(REPO, ".venv", f"ci-py{args.python}")
    vpy = os.path.join(venv, "Scripts" if os.name == "nt" else "bin", "python.exe" if os.name == "nt" else "python")
    if args.fresh and os.path.isdir(venv):
        shutil.rmtree(venv)
    installer = None
    if not os.path.isfile(vpy):
        print(f"creating {venv} from {base}")
        installer = make_venv(base, venv)
    with open(LOCK, "rb") as f:
        lock_hash = hashlib.sha256(f.read()).hexdigest()
    stamp = os.path.join(venv, "lock.sha256")
    if not (os.path.isfile(stamp) and open(stamp).read().strip() == lock_hash):
        print("installing locks/full.txt with --require-hashes (as the workflow does)")
        has_pip = subprocess.run([vpy, "-m", "pip", "--version"], capture_output=True).returncode == 0
        if has_pip and installer != "uv":
            subprocess.run([vpy, "-m", "pip", "install", "--quiet", "--upgrade", "pip"], check=False)
            subprocess.run([vpy, "-m", "pip", "install", "--quiet", "--require-hashes", "-r", LOCK], check=True)
        else:
            uv = find_uv()
            if not uv:
                raise SystemExit("the environment has no pip and uv is not available")
            subprocess.run([uv, "pip", "install", "--quiet", "--python", vpy, "--require-hashes", "-r", LOCK], check=True)
        with open(stamp, "w") as f:
            f.write(lock_hash)
    r = subprocess.run([vpy, "-c", "import sys, numpy; print(f'Python {sys.version.split()[0]} ({sys.executable}), "
                                   "numpy {numpy.__version__}')"], capture_output=True, text=True)
    print("suites run under:", r.stdout.strip() or r.stderr.strip())
    return subprocess.run([vpy, os.path.join(REPO, "tools", "run_tests.py")] + args.filters, cwd=REPO).returncode


if __name__ == "__main__":
    sys.exit(main())
