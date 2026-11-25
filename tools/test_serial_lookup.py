"""
Serial lookup tests — especially makers with RESTARTED serial series.

Buffet restarted numbering at 1 in 1950, LeBlanc/Vito at 1 in 1970, and
Yanagisawa switched from 8-digit date-style to 6-digit serials in 1980.
The old lookup scanned all entries as one pool and flip-flopped between
series (Buffet 26000 -> 1922 but 26200 -> 1977). The fixed lookup splits
the data into ascending series and reports every series that matches.
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Never the user's real settings folder: an isolated profile, set before any
# project module is imported (tools/test_config_isolation.py checks this).
import tempfile  # noqa: E402
os.environ["SAXSHOP_CONFIG_DIR"] = tempfile.mkdtemp(prefix="ssc-serial-lookup-")

from i18n import init_translation  # noqa: E402
init_translation("en")

from library_features import lookup_serial_year, _split_serial_series
from serials import SERIAL_DATA

passed = 0
failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


# ============================================================
print("\n=== Series splitting ===")
# ============================================================

check("Buffet splits into 2 series",
      len(_split_serial_series(SERIAL_DATA["Buffet"])) == 2)
check("LeBlanc / Vito splits into 2 series",
      len(_split_serial_series(SERIAL_DATA["LeBlanc / Vito"])) == 2)
check("Yanagisawa splits into 2 series",
      len(_split_serial_series(SERIAL_DATA["Yanagisawa"])) == 2)
check("Conn stays 1 series",
      len(_split_serial_series(SERIAL_DATA["Conn"])) == 1)

single_series_makers = sum(
    1 for data in SERIAL_DATA.values()
    if isinstance(data, list) and len(_split_serial_series(data)) == 1)
print(f"  (info: {single_series_makers} single-series makers, "
      f"{len(SERIAL_DATA) - single_series_makers} multi-series)")


# ============================================================
print("\n=== Single-series makers: behavior unchanged ===")
# ============================================================

check("Conn 15000 -> 1908", lookup_serial_year("Conn", "15000") == "1908")
check("Conn 1 -> 1895", lookup_serial_year("Conn", "1") == "1895")
check("Conn 0 -> Too old", lookup_serial_year("Conn", "0") == "Too old / Unknown")
check("Unknown maker handled",
      lookup_serial_year("Atlantis Brass Works", "123") == "Manufacturer data not found.")
check("Non-numeric serial handled",
      lookup_serial_year("Conn", "abc") == "Invalid Serial Number")
check("Empty serial handled", lookup_serial_year("Conn", "") == "")
check("Letters stripped from serial (No. 15000)",
      lookup_serial_year("Conn", "No. 15000") == "1908")
check("Thousands comma accepted (265,316)",
      lookup_serial_year("Conn", "265,316") == "1935")


# ============================================================
print("\n=== Buffet: ambiguous serials report both series ===")
# ============================================================

r = lookup_serial_year("Buffet", "26000")
check(f"Buffet 26000 reports vintage year ({r!r})", "1922" in r)
check("Buffet 26000 reports modern year", "1976" in r)
check("Buffet 26000 joins with 'or'", "or" in r)
check("Buffet 26000 includes series spans", "1866" in r and "1985" in r)

r = lookup_serial_year("Buffet", "27500")
check(f"Buffet 27500 reports both candidates ({r!r})",
      "1924" in r and "1978" in r)

# The old algorithm's flip-flop pair: consecutive-ish serials must now
# BOTH report BOTH series instead of ping-ponging between eras.
r1 = lookup_serial_year("Buffet", "27000")
r2 = lookup_serial_year("Buffet", "27300")
check("Buffet 27000 and 27300 both dual-report (no more flip-flop)",
      "or" in r1 and "or" in r2)

r = lookup_serial_year("Buffet", "5000")
check(f"Buffet 5000 reports both candidates ({r!r})",
      "1882" in r and "1957" in r)


# ============================================================
print("\n=== Inside matches suppress extrapolation ===")
# ============================================================

check("Buffet 33000 -> 1982 only (vintage series ended at 30000)",
      lookup_serial_year("Buffet", "33000") == "1982")
check("Buffet 40000 -> 1985 or later (beyond both series)",
      lookup_serial_year("Buffet", "40000") == "1985 or later")

# Yanagisawa's 8-digit serials sit ABOVE the whole 6-digit series, so the
# 6-digit series would match by extrapolation — it must be suppressed.
check("Yanagisawa 12745400 -> 1974 only (8-digit series, no 2000 ghost)",
      lookup_serial_year("Yanagisawa", "12745400") == "1974")
check("Yanagisawa 150000 -> 1988 only",
      lookup_serial_year("Yanagisawa", "150000") == "1988")


# ============================================================
print("\n=== LeBlanc / Vito ===")
# ============================================================

check("LeBlanc 500 -> 1970 only (old series starts at 16000)",
      lookup_serial_year("LeBlanc / Vito", "500") == "1970")
r = lookup_serial_year("LeBlanc / Vito", "25000")
check(f"LeBlanc 25000 dual-reports ({r!r})", "1966" in r and "or" in r)


# ============================================================
print("\n=== The 2026-11-24 corrections ===")
# ============================================================

def expect(maker, serial, want):
    got = lookup_serial_year(maker, serial)
    check(f"{maker} {serial} -> {want!r} (got {got!r})", got == want)


expect("Selmer Paris", "1130", "1922")
expect("Selmer Paris", "1700", "1923")
expect("Selmer Paris", "2900", "1924")
expect("Selmer Paris", "4200", "1926")
expect("Selmer Paris", "46340", "1951")
expect("Selmer Paris", "155000", "1968")
expect("Selmer Paris", "160210", "1969")
expect("Conn", "341850", "1951")
expect("Conn", "345000", "1951")
expect("Conn", "348296", "1952")
expect("Keilwerth", "19700", "1944–1945")
expect("Keilwerth", "20000", "1944–1945")
expect("Keilwerth", "20001", "1947")
check("Yamaha is not in the list (serials repeat between models)", "Yamaha" not in SERIAL_DATA)

# Past the end of every series: the one whose chart ends closest below.
expect("Yanagisawa", "12791801", "1979")
expect("Yanagisawa", "12791802", "1979 or later")
expect("Yanagisawa", "364566", "2000 or later")

# How Conn 1951 hid: 1952 started one number after it. Keilwerth's 1948
# is one number wide in every chart we have.
NARROW_OK = {("Keilwerth", 1948)}
RESTARTS = {"Buffet", "LeBlanc / Vito", "Yanagisawa"}
narrow, stray = [], []
for maker_name, rows in SERIAL_DATA.items():
    for (a, ya), (b, yb) in zip(rows, rows[1:]):
        if 0 <= b - a <= 1 and (maker_name, ya) not in NARROW_OK:
            narrow.append((maker_name, a, ya, b, yb))
        if b < a and maker_name not in RESTARTS:
            stray.append((maker_name, a, ya, b, yb))
check(f"No year is one number wide or doubled {narrow}", not narrow)
check(f"A serial only goes down where a maker restarted {stray}", not stray)


# ============================================================
print("\n=== 'or later', year ranges ===")
# ============================================================

expect("Selmer Paris", "579971", "1999")
expect("Selmer Paris", "650000", "1999 or later")
expect("Buescher", "350000", "1955")          # the entry itself is that year
expect("Buescher", "352000", "1955–1960")
expect("Buescher", "349999", "1952–1955")
expect("Buffet", "22600", "1914–1916"
       " (1866–1930 series)\nor 1974 (1950–1985 series)")
expect("Hohner", "7000", "1954–1960")
expect("Hohner", "20000", "1960–1972")        # an era already says when it ends
expect("Conn", "1000001", "1963 or later")


# ============================================================
print("\n=== Conn letter years (1963-1972) ===")
# ============================================================

for serial, want in (("C00501", "1963"), ("C73853", "1963"), ("C73854", "1964"), ("D00001", "1964"),
                     ("E54105", "1964"), ("E54106", "1965"), ("H31247", "1966"), ("K35274", "1967"),
                     ("L20454", "1968"), ("L99999", "1968"), ("N12345", "1970"), ("P12345", "1971"),
                     ("R99999", "1972"), ("c73854", "1964"), ("C 73854", "1964"), ("C-73854", "1964")):
    expect("Conn", serial, want)
# An M can also sit beside the number of a vintage Conn.
expect("Conn", "M15000", "1908 (1895–1963 series)\nor 1969 (1963–1972 series)")
expect("Conn", "M 15000", "1908 (1895–1963 series)\nor 1969 (1963–1972 series)")
expect("Conn", "73854", "1921")               # no letter, no letter year
expect("Conn", "M264000", "1935")             # six digits: the vintage numbering
expect("Martin", "C73854", "1926")            # Conn's letters are Conn's only
for serial in ("C00500", "B12345", "S12345"):
    r = lookup_serial_year("Conn", serial)
    check(f"Conn {serial} is outside the letter years and says so ({r!r})", "C00501" in r and "1972" in r)


# ============================================================
print("\n=== Numbers that can't be dated ===")
# ============================================================

for serial in ("86xxx", "86XXX", "86???", "86***", "264 xxx"):
    r = lookup_serial_year("Conn", serial)
    check(f"Conn {serial} asks for the full number ({r!r})", "full" in r and not any(c.isdigit() for c in r))
for serial in ("12345 / 678", "1951 345000 tenor", "9" * 13, "²"):
    r = lookup_serial_year("Conn", serial)
    check(f"Conn {serial!r} is not dated ({r!r})", r and not any(c.isdigit() for c in r))


# ============================================================
print(f"\n=== Total: {passed} passed, {failed} failed ===")
# ============================================================

sys.exit(0 if failed == 0 else 1)
