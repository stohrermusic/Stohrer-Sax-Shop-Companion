# Web data imports: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Web Data Sync ("Import Matt's")

*Moved from CLAUDE-web.md, section "Web Data Sync & Screw Specs > Web Data Sync ("Import Matt's")" (2026-11-10).*

All use `urllib.request` (stdlib) with a 10-second timeout and `get_ssl_context()` from config.py: certifi's bundle (every build ships it, the Intel Mac build since 2026-11-10), then the system's default certificates, then it raises and the caller shows its "Could not fetch ..." box. It never returns an unverified context: until 2026-11-10 it did (check_hostname False, CERT_NONE) whenever the default certificates could not be loaded, which silently defeated HTTPS for these downloads. `tools/test_tls_context.py` gates it (no code path returns CERT_NONE; a certificate failure reaches the user as the download error), and `--selftest` calls `config.assert_tls_ready()` so the shipped bundle is checked to have certificate authorities on every platform, the Intel Mac included. The JSON formats on the website match the app's internal format exactly, so no conversion is needed.

## Serial lookup

**Corrections and the lookup's rules (2026-11-25).** Seven Selmer Paris starts follow the Selmer chart on getasax.com. Conn 1952 starts at 348296 (The Conn Loyalist's number): Conn-Selmer's own list prints 341851, one above 1951, so 1951 was one number wide and nearly every 1951 Conn came back as 1952. Keilwerth 20001 is 1947 only, and 19601 to 20000 is 1944 to 1945, as on stohrermusic.com. Yamaha is out: its serials repeat from one model to another, and Matt took it off the website on 2026-02-16 ("those numbers seem to be off").

The lookup follows the website's (its commit b0b22c6): "or later" past the last entry, a range where the chart skips years, Conn letter numbers (1963 to 1972; M shows both years because a vintage Conn can have an M beside its number), and no date for a number with hidden digits ("86xxx" kept the digits 86 and was dated 1895). Past the end of every series, the series whose last entry is closest below the number answers; taking the most recent one dated Yanagisawa 12791802 as 2000 instead of 1979. Anything that is not one number is refused, where the digits used to be run together.

Compared with the website's own script in headless Edge: 5,031 inputs, 4,797 dated, the same years for all but M00000, N00000, P00000 and R00000, which the website refuses. Nothing syncs the three copies of the data (here, stohrermusic.com, getasax.com). `test_serial_lookup` also fails on a year one number wide, which is how the Conn slip hid.
