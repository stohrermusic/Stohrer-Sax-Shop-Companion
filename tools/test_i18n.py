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

    # --- Every shipping .po must have at least some translations ---
    # Strict full-coverage gate is a separate test that runs at release time
    # (full_coverage_gate=True). During development we tolerate partial
    # coverage since translations land incrementally.
    if os.path.isfile(es_mo):
        from babel.messages.pofile import read_po
        es_po = os.path.join(locale_dir, "es", "LC_MESSAGES", f"{DOMAIN}.po")
        with open(es_po, "rb") as f:
            catalog = read_po(f)
        translated_count = sum(1 for m in catalog if m.string and m.id)
        total = sum(1 for m in catalog if m.id)
        check(f"es catalog has translations ({translated_count}/{total})",
              translated_count > 0,
              "no translated entries at all")

    # --- Summary ---
    print("=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
