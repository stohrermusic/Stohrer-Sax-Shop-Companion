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
check("Letters stripped from serial (M 15000)",
      lookup_serial_year("Conn", "M 15000") == "1908")


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
check("Buffet 40000 -> 1985 only (beyond all series: latest one wins)",
      lookup_serial_year("Buffet", "40000") == "1985")

# Yanagisawa's 8-digit serials sit ABOVE the whole 6-digit series, so the
# 6-digit series would match by extrapolation — it must be suppressed.
check("Yanagisawa 12745400 -> 1974 only (8-digit series, no 2000 ghost)",
      lookup_serial_year("Yanagisawa", "12745400") == "1974")
check("Yanagisawa 150000 -> single year, no 'or'",
      "or" not in lookup_serial_year("Yanagisawa", "150000"))


# ============================================================
print("\n=== LeBlanc / Vito ===")
# ============================================================

check("LeBlanc 500 -> 1970 only (old series starts at 16000)",
      lookup_serial_year("LeBlanc / Vito", "500") == "1970")
r = lookup_serial_year("LeBlanc / Vito", "25000")
check(f"LeBlanc 25000 dual-reports ({r!r})", "1966" in r and "or" in r)


# ============================================================
print(f"\n=== Total: {passed} passed, {failed} failed ===")
# ============================================================

sys.exit(0 if failed == 0 else 1)
