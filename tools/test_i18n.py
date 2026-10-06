#!/usr/bin/env python3
"""Tests for the i18n machinery.

Covers:
- locale directory resolution (source + frozen)
- init_translation switches the active catalog
- _() returns translated string when a catalog is present
- _() returns source string for English / unknown languages (fallback)
- available_languages() reflects what's actually on disk
- The Spanish pilot catalog covers every string in saxshop.pot (no
  fuzzy / empty entries shipping in v1)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from i18n import (  # noqa: E402
    DOMAIN, LANGUAGE_NAMES, _, _locale_dir, available_languages,
    current_language, init_translation,
)


PASS = "  PASS"
FAIL = "  FAIL"
results = []


def check(label, ok, detail=""):
    results.append(ok)
    print(f"{PASS if ok else FAIL}  {label}" + (f"  -- {detail}" if not ok and detail else ""))


def main():
    print("i18n Tests")
    print("=" * 60)

    locale_dir = _locale_dir()
    check("locale dir exists", os.path.isdir(locale_dir), locale_dir)

    pot_path = os.path.join(locale_dir, "saxshop.pot")
    check("saxshop.pot template exists", os.path.isfile(pot_path),
          "run tools/extract_strings.py")

    # --- English (source language) ---
    init_translation("en")
    check("init_translation('en') sets current_language",
          current_language() == "en")
    src = "Resonance added!"
    check("_('Resonance added!') returns source under en",
          _(src) == src)

    # --- Unknown language falls back to English ---
    init_translation("xx_invalid")
    check("unknown language falls back to source",
          _(src) == src)

    # --- Spanish (pilot translation) ---
    init_translation("es")
    es_mo = os.path.join(locale_dir, "es", "LC_MESSAGES", f"{DOMAIN}.mo")
    if not os.path.isfile(es_mo):
        check("es .mo file present", False,
              f"missing {es_mo} — run tools/compile_translations.py")
    else:
        check("es .mo file present", True)
        check("init_translation('es') sets current_language",
              current_language() == "es")
        translated = _(src)
        check(f"_('{src}') translates under es",
              translated != src and translated.strip() != "",
              f"got: {translated!r}")

    # --- available_languages() reflects compiled catalogs ---
    init_translation("en")
    langs = available_languages()
    codes = [code for code, _name in langs]
    check("available_languages() includes 'en'", "en" in codes)
    # If es .mo exists, it must show up in the list
    if os.path.isfile(es_mo):
        check("available_languages() includes 'es' when compiled",
              "es" in codes)

    # --- LANGUAGE_NAMES covers all expected v1 languages ---
    expected = {"en", "es", "de", "fr", "it"}
    check("LANGUAGE_NAMES covers en/es/de/fr/it",
          expected.issubset(LANGUAGE_NAMES.keys()),
          f"have: {set(LANGUAGE_NAMES.keys())}")

    # --- Every shipping catalog is complete: no empty, no fuzzy ---
    # The workflow in CLAUDE.md translates every new string in the same
    # commit (extract -> update -> translate -> compile), so a fuzzy or
    # empty entry in ANY catalog is a slip, not a stage. This used to
    # check only Spanish; the other three could ship half-done unnoticed.
    try:
        from babel.messages.pofile import read_po
    except ImportError:
        read_po = None
        check("babel available for catalog checks", False, "pip install babel")
    if read_po is not None:
        for lang in sorted(LANGUAGE_NAMES):
            if lang == "en":
                continue
            po = os.path.join(locale_dir, lang, "LC_MESSAGES", f"{DOMAIN}.po")
            mo = os.path.join(locale_dir, lang, "LC_MESSAGES", f"{DOMAIN}.mo")
            check(f"{lang}: .po and compiled .mo present",
                  os.path.isfile(po) and os.path.isfile(mo))
            if not os.path.isfile(po):
                continue
            with open(po, "rb") as f:
                catalog = read_po(f)
            ids = [m for m in catalog if m.id]
            empty = [m.id for m in ids if not m.string]
            fuzzy = [m.id for m in ids if m.fuzzy]
            check(f"{lang}: every string translated ({len(ids) - len(empty)}/{len(ids)})",
                  not empty, f"empty: {[str(e)[:40] for e in empty[:3]]}")
            check(f"{lang}: no fuzzy entries", not fuzzy,
                  f"fuzzy: {[str(z)[:40] for z in fuzzy[:3]]}")
            if os.path.isfile(mo):
                check(f"{lang}: .mo is not older than .po",
                      os.path.getmtime(mo) >= os.path.getmtime(po) - 1,
                      "run tools/compile_translations.py")

    # --- Summary ---
    print("=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
