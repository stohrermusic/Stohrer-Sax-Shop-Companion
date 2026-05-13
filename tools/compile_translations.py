#!/usr/bin/env python3
"""Compile every locale/<lang>/LC_MESSAGES/saxshop.po into a .mo file.

The .mo files are what the running app actually reads. Commit both .po
(source) and .mo (compiled) so end-user builds don't need babel.

Run from the repo root:
    python tools/compile_translations.py
"""
import os
import subprocess
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCALE_DIR = os.path.join(REPO_ROOT, "locale")
DOMAIN = "saxshop"


def main():
    if not os.path.isdir(LOCALE_DIR):
        print(f"Missing {LOCALE_DIR}.")
        sys.exit(1)

    cmd = [
        sys.executable, "-m", "babel.messages.frontend",
        "compile",
        "-d", LOCALE_DIR,
        "-D", DOMAIN,
        "-f",  # also compile fuzzy entries so partial translations are visible
    ]
    print("Compiling .po -> .mo for all locales...")
    result = subprocess.run(cmd, cwd=REPO_ROOT)
    if result.returncode != 0:
        print(f"Compile failed (exit {result.returncode}).")
        sys.exit(result.returncode)
    print("Done.")


if __name__ == "__main__":
    main()
