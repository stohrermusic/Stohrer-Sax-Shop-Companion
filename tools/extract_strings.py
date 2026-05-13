#!/usr/bin/env python3
"""Extract translatable strings into locale/saxshop.pot.

Scans every .py file (per babel.cfg) for _() and ngettext() calls and
writes a master template (.pot) file. Translators copy this template
into locale/<lang>/LC_MESSAGES/saxshop.po and fill in msgstr entries.

Run from the repo root:
    python tools/extract_strings.py
"""
import os
import subprocess
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
POT_PATH = os.path.join(REPO_ROOT, "locale", "saxshop.pot")
BABEL_CFG = os.path.join(REPO_ROOT, "babel.cfg")


def main():
    os.makedirs(os.path.dirname(POT_PATH), exist_ok=True)
    cmd = [
        sys.executable, "-m", "babel.messages.frontend",
        "extract",
        "-F", BABEL_CFG,
        "-o", POT_PATH,
        "--project=SaxShopCompanion",
        "--copyright-holder=Matt Stohrer",
        "--msgid-bugs-address=stohrermusic@gmail.com",
        ".",
    ]
    print("Extracting translatable strings...")
    result = subprocess.run(cmd, cwd=REPO_ROOT)
    if result.returncode != 0:
        print(f"Extraction failed (exit {result.returncode}).")
        sys.exit(result.returncode)
    print(f"Wrote {POT_PATH}")


if __name__ == "__main__":
    main()
