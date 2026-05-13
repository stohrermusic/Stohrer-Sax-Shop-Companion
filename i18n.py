"""Translation (gettext) helpers for Stohrer Sax Shop Companion.

Usage at app startup (main.py, before importing UI modules):

    from config import load_settings
    from i18n import init_translation
    settings = load_settings()
    init_translation(settings.get("language", "en"))

Usage in any module:

    from i18n import _, ngettext
    label = _("Save")
    msg = ngettext("Imported {n} capture", "Imported {n} captures", n).format(n=n)

Do NOT use `_()` at module-level constants — modules are imported before
init_translation can run for them. Use a function or defer to instance
attributes inside __init__ if you need translatable strings as data.
"""

import gettext
import os
import sys

DOMAIN = "saxshop"

# Native-name display labels for the language picker. Keep these in the
# native language so users recognize their own language regardless of the
# UI language currently active.
LANGUAGE_NAMES = {
    "en": "English",
    "es": "Español",
    "de": "Deutsch",
    "fr": "Français",
    "it": "Italiano",
}

_translation = gettext.NullTranslations()
_current_lang = "en"


def _locale_dir():
    """Locate the bundled locale directory in both source and frozen builds."""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "locale")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "locale")


def init_translation(lang="en"):
    """Load and install the translation catalog for `lang`.

    Installs `_()` and `ngettext()` into builtins so any module can use
    them without an explicit import. This is the standard gettext pattern
    (see Python docs: gettext.install). Modules using `_` as a throwaway
    loop variable will still shadow the builtin within that local scope,
    so existing `for _, x in ...` patterns continue to work.

    Safe to call multiple times. Silently falls back to source strings if
    the requested language has no compiled .mo file.
    """
    global _translation, _current_lang
    try:
        _translation = gettext.translation(
            DOMAIN, _locale_dir(), languages=[lang], fallback=True
        )
        _current_lang = lang
    except Exception:
        _translation = gettext.NullTranslations()
        _current_lang = "en"
    _translation.install(names=("ngettext",))
    return _translation


def _(message):  # noqa: F811  (intentionally provides a module-level handle)
    """Module-level handle to the active translator. Same result as the
    builtin `_()` installed by `init_translation()`. Use this in tests or
    when you need to explicitly route through this module."""
    return _translation.gettext(message)


def ngettext(singular, plural, n):  # noqa: F811
    """Module-level handle to the active plural translator."""
    return _translation.ngettext(singular, plural, n)


def current_language():
    """Return the active language code (e.g. 'es')."""
    return _current_lang


def available_languages():
    """Return list of (code, display_name) tuples for every language with a
    compiled .mo file, plus English (source language) which always works."""
    available = [("en", LANGUAGE_NAMES["en"])]
    locale_dir = _locale_dir()
    if not os.path.isdir(locale_dir):
        return available
    for code in sorted(os.listdir(locale_dir)):
        if code == "en":
            continue
        mo_path = os.path.join(locale_dir, code, "LC_MESSAGES", f"{DOMAIN}.mo")
        if os.path.isfile(mo_path):
            display = LANGUAGE_NAMES.get(code, code)
            available.append((code, display))
    return available
