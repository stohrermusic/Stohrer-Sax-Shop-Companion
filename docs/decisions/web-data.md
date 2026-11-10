# Web data imports: decisions and history

This file holds the why and the history behind the rules in the CLAUDE files, and the full detail they summarise: what was tried, measured, rejected and decided, with dates and the user's own words. The paragraphs were moved here verbatim from the CLAUDE files on 2026-11-10 (HEAD 01071f5); each section names where it came from. Add new decisions to the matching section, dated.

## Web Data Sync ("Import Matt's")

*Moved from CLAUDE-web.md, section "Web Data Sync & Screw Specs > Web Data Sync ("Import Matt's")" (2026-11-10).*

All use `urllib.request` (stdlib) with a 10-second timeout and `get_ssl_context()` from config.py: certifi's bundle (every build ships it, the Intel Mac build since 2026-11-10), then the system's default certificates, then it raises and the caller shows its "Could not fetch ..." box. It never returns an unverified context: until 2026-11-10 it did (check_hostname False, CERT_NONE) whenever the default certificates could not be loaded, which silently defeated HTTPS for these downloads. `tools/test_tls_context.py` gates it (no code path returns CERT_NONE; a certificate failure reaches the user as the download error), and `--selftest` calls `config.assert_tls_ready()` so the shipped bundle is checked to have certificate authorities on every platform, the Intel Mac included. The JSON formats on the website match the app's internal format exactly, so no conversion is needed.
