---
description: Audit project documentation against current codebase state. Catches drift; does not restructure without explicit approval.
---

Audit whether the project documentation accurately reflects the current
codebase. Scope: `CLAUDE.md`, every `CLAUDE-*.md` companion file at the
repo root (`CLAUDE-architecture.md`, `CLAUDE-engines.md`, `CLAUDE-toner.md`,
`CLAUDE-web.md`), `README.md`, and the canonical memory index
`~/.claude/projects/C--code-saxshopcompanion/memory/MEMORY.md`.

**Do NOT restructure** by default. The split into per-subsystem
`CLAUDE-*.md` files imported via `@`-statements at the bottom of
`CLAUDE.md`, and the auto-memory layout, are intentional. Your job is to
catch content drift, not reorganize. If you think restructuring is
warranted, flag it as a recommendation and wait for explicit approval
before touching the structure.

## Steps

1. **Read the doc tree.**
   - `CLAUDE.md` (project entrypoint, always loaded by Claude Code).
   - `CLAUDE-architecture.md`, `CLAUDE-engines.md`, `CLAUDE-toner.md`,
     `CLAUDE-web.md` (loaded via `@imports` at the bottom of CLAUDE.md).
   - `README.md` at repo root (user-facing; check version stamp + status
     paragraph + any test-count or msgid claims).
   - `~/.claude/projects/C--code-saxshopcompanion/memory/MEMORY.md`
     (canonical auto-memory index).

2. **Check claimed facts against reality.** Common drift sources for
   this project:
   - **Test count**: count `tools/test_*.py` files and compare to "All
     test suites (N files)" in `CLAUDE.md` and any test-count claim in
     `README.md`. Listed-by-name section in CLAUDE.md should also match.
   - **`APP_VERSION`**: grep `APP_VERSION` in `config.py` and compare
     against the example in CLAUDE.md's Windows installer build command
     (`iscc /DAppVersion=2.21 installer.iss` — should match latest
     release).
   - **i18n msgid count**: `grep -c "^msgid " locale/saxshop.pot` and
     compare against any claim like "~1268 strings" in CLAUDE.md,
     README.md, or recent release notes.
   - **Shipped languages**: `ls locale/*/LC_MESSAGES/saxshop.mo` and
     compare against any "ships in N languages" claim and the
     `LANGUAGE_NAMES` dict in `i18n.py`.
   - **Bundled runtime assets** list in CLAUDE.md's "Bundled Runtime
     Assets" section — grep `build.py` for `--add-data` lines and make
     sure every bundled directory/file is documented.
   - **Branching strategy** — `git branch -a` should still show
     main/beta/gamma; if `gamma` is gone, update CLAUDE.md.

3. **Check for undocumented recent work.** `git log --oneline -25` for
   commits since the last doc refresh. Ask:
   - Are there new modules, new tabs, new configuration keys, or new
     bundled assets the docs haven't caught up to?
   - Anything in the code that contradicts an architectural claim in
     `CLAUDE.md` or the four `CLAUDE-*.md` companions?
   - New test suites that aren't in CLAUDE.md's authoritative test list?
   - Any "we decided X" memory notes that should also be reflected in
     CLAUDE.md as a design constraint (e.g. Phil Noy credit, macOS dark
     mode skip, descriptor live-gauge removal)?

4. **Check `memory/MEMORY.md` is the canonical index.** CLAUDE.md
   should NOT re-list every memory file. It should point readers at the
   memory directory for full context (project decisions, "we tried this
   and it didn't work" lessons, references). If the auto-memory has
   added entries since CLAUDE.md was last touched, that's fine —
   MEMORY.md self-maintains. Don't try to mirror it into CLAUDE.md.

5. **Verify the `@imports`.** The bottom of CLAUDE.md should use
   explicit `@`-import syntax to pull in the four companion docs. If
   they're listed as a prose bullet list but not actually imported,
   Claude Code isn't loading them automatically — that's drift to fix.

6. **Propose targeted edits.** Show the planned changes in diff-like
   form *before* applying. Do not change structure (move files, rename,
   add new companion docs) without explicit consent.

7. **Report.** Summarize: what drifted, what was updated, what was
   noticed but intentionally left alone (with reasoning), anything
   recommended but not changed.

## Do NOT touch

- **Memory file content** — memory has its own lifecycle (handled by
  auto-memory). The per-file content of `~/.claude/projects/.../memory/*.md`
  is not editable by initbig.
- The `@import` list at the bottom of CLAUDE.md — unless a new
  `CLAUDE-*.md` file was genuinely added at root, don't modify this.
- Sections that describe deliberate choices still known to be accurate
  — don't "improve" these for stylistic reasons. Examples:
  - macOS dark mode skip (let Aqua handle it)
  - Phil Noy credit is non-negotiable
  - Live descriptor gauges removed deliberately
  - Restart-required language switch (chosen explicitly over live)
  - macOS code signing deferred ($100/yr cost)
- Frozen historical claims that are part of the project's record (e.g.
  "v2.0 shipped 2026-04-06").

## Trigger conditions for running this

- After a meaningful feature pass or hygiene pass ships.
- When Matt mentions "let's update the docs" or asks for an audit.
- As a final check before claiming a body of work is complete (e.g.
  before a release-notes-style commit).
- When recent commits touched architecture, added tests, added bundled
  assets, or shipped a new release — likely doc drift to verify.
