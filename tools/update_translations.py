#!/usr/bin/env python3
"""Merge new strings from locale/saxshop.pot into existing .po files.

Run after extract_strings.py to bring per-language catalogs up to date.
Existing translations are preserved; new entries are marked fuzzy.

Run from the repo root:
    python tools/update_translations.py
"""
import os
import subprocess
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCALE_DIR = os.path.join(REPO_ROOT, "locale")
POT_PATH = os.path.join(LOCALE_DIR, "saxshop.pot")
DOMAIN = "saxshop"


def main():
    if not os.path.isfile(POT_PATH):
        print(f"Missing {POT_PATH} — run extract_strings.py first.")
        sys.exit(1)

    cmd = [
        sys.executable, "-m", "babel.messages.frontend",
        "update",
        "-i", POT_PATH,
        "-d", LOCALE_DIR,
        "-D", DOMAIN,
    ]
    print("Merging template into existing translations...")
    result = subprocess.run(cmd, cwd=REPO_ROOT)
    if result.returncode != 0:
        print(f"Update failed (exit {result.returncode}).")
        sys.exit(result.returncode)
    print("Done. Review fuzzy entries in each locale/<lang>/LC_MESSAGES/saxshop.po")


if __name__ == "__main__":
    main()
