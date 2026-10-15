"""Regenerate the dependency lock files in locks/ (what CI installs).

    pip install uv
    python tools/update_locks.py              # re-resolve, keep versions where possible
    python tools/update_locks.py --upgrade    # move everything to the newest allowed
    python tools/update_locks.py --exclude-newer 2026-10-14T18:17:05Z

Why locks: requirements.txt only sets floors, so every CI build used to
install whatever was newest that day. OpenCV 5.0 arrived on PyPI on
2026-07-02 and went into every build after it unannounced; the camera
calibration dialog had been written against OpenCV 4's array shapes and
was broken in every release from v2.7 to v2.85. With locks, a new version
reaches a build only through a commit that changes a lock -- which CI then
tests on all four platforms before anything is built.

Every lock is resolved for CI's Python (3.11) with a hash for every file,
so pip refuses a download that doesn't match (--require-hashes). The full
lock is universal (one file, environment markers for Windows/macOS/Linux);
the Intel Mac and lint sets are subsets, resolved against the full lock so
a package never has two versions across jobs.

requirements.txt / requirements-dev.txt stay the human-facing floors for
anyone running from source; edit those (or locks/*.in) and re-run this.
"""
import argparse
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PYTHON = "3.11"  # must match actions/setup-python in .github/workflows

# (output, inputs, extra uv args)
LOCKS = [
    ("locks/full.txt", ["requirements.txt", "requirements-dev.txt"],
     ["--universal"]),
    # The Intel Mac build ships numpy but no audio stack and no camera
    # stack (see CLAUDE.md); its tests and its build install different
    # sets, and the build must not pick up Pillow it doesn't ship.
    ("locks/intel-mac-test.txt", ["locks/intel-mac-test.in"],
     ["--python-platform", "x86_64-apple-darwin", "--constraint", "locks/full.txt"]),
    ("locks/intel-mac-build.txt", ["locks/intel-mac-build.in"],
     ["--python-platform", "x86_64-apple-darwin", "--constraint", "locks/full.txt"]),
    ("locks/lint.txt", ["locks/lint.in"],
     ["--universal", "--constraint", "locks/full.txt"]),
]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--upgrade", action="store_true",
                    help="allow every package to move to its newest allowed version")
    ap.add_argument("--exclude-newer", metavar="TIMESTAMP",
                    help="ignore releases published after this (RFC 3339)")
    args = ap.parse_args()

    for out, inputs, extra in LOCKS:
        cmd = [sys.executable, "-m", "uv", "pip", "compile", *inputs,
               "--output-file", out, "--python-version", PYTHON,
               "--generate-hashes", "--quiet",
               "--custom-compile-command", "python tools/update_locks.py",
               *extra]
        if args.upgrade and out == "locks/full.txt":
            # Subsets follow the full lock through their --constraint.
            cmd.append("--upgrade")
        if args.exclude_newer:
            cmd += ["--exclude-newer", args.exclude_newer]
        print("$", " ".join(cmd[1:]))
        r = subprocess.run(cmd, cwd=ROOT)
        if r.returncode:
            return r.returncode
    return 0


if __name__ == "__main__":
    sys.exit(main())
