#!/usr/bin/env bash
# SessionStart hook: fetch from origin and report how this checkout compares.
#
# Matt develops Sax Shop Companion from several different computers, so a local
# checkout is routinely behind origin/beta -- and sometimes also holds unpushed
# local commits, which makes the branch diverged rather than merely stale.
# Starting work off a stale tree has burned us before (see the "Branching
# Strategy" section of CLAUDE.md).
#
# This hook only FETCHES and REPORTS. It never merges, rebases, or checks out --
# reconciling stays a deliberate step, because the right move differs between
# "behind" (pull) and "diverged" (pull --rebase, then check for conflicts).
#
# Output is a SessionStart hook JSON payload whose additionalContext is injected
# into the model's context at the start of the session.

set -u

emit() {
    # JSON-escape backslashes and double quotes, then wrap in the hook payload.
    # Command substitution (not a pipe into `read`) -- sed emits no trailing
    # newline here, so `read` would return non-zero and the message would be lost.
    escaped=$(printf '%s' "$1" | sed -e 's/\\/\\\\/g' -e 's/"/\\"/g')
    printf '{"hookSpecificOutput":{"hookEventName":"SessionStart","additionalContext":"%s"}}\n' "$escaped"
    exit 0
}

# Anchor to this repo regardless of the cwd the hook happens to launch with.
# The script lives at <repo>/.claude/hooks/, so the repo root is three levels up.
repo_root=${CLAUDE_PROJECT_DIR:-$(cd -- "$(dirname -- "$0")/../.." && pwd)}
cd -- "$repo_root" 2>/dev/null || exit 0

git rev-parse --is-inside-work-tree >/dev/null 2>&1 || exit 0

# Network may be down or the remote unreachable; that must not block the session.
git fetch --all --prune --quiet >/dev/null 2>&1 || \
    emit "git sync check: fetch from origin failed (offline?). Branch state below may be stale -- re-run 'git fetch' before trusting it."

branch=$(git rev-parse --abbrev-ref HEAD 2>/dev/null)
upstream=$(git rev-parse --abbrev-ref --symbolic-full-name '@{u}' 2>/dev/null) || exit 0
[ -n "$upstream" ] || exit 0

behind=$(git rev-list --count "HEAD..$upstream" 2>/dev/null || echo 0)
ahead=$(git rev-list --count "$upstream..HEAD" 2>/dev/null || echo 0)
dirty=""
git diff --quiet 2>/dev/null && git diff --cached --quiet 2>/dev/null || \
    dirty=" The working tree also has uncommitted changes -- stash or commit them before rebasing."

if [ "$behind" = "0" ] && [ "$ahead" = "0" ]; then
    emit "git sync check: $branch is in sync with $upstream.${dirty}"
elif [ "$ahead" = "0" ]; then
    emit "git sync check: $branch is BEHIND $upstream by $behind commit(s). Run 'git pull --rebase' BEFORE reading code or making edits -- the local config.py APP_VERSION, release notes, and memory files are all unreliable until you do.${dirty}"
elif [ "$behind" = "0" ]; then
    emit "git sync check: $branch is AHEAD of $upstream by $ahead unpushed commit(s). Nothing to pull, but this work has not reached the other machines yet.${dirty}"
else
    emit "git sync check: $branch has DIVERGED from $upstream -- $ahead local unpushed, $behind on the remote. Run 'git pull --rebase' to replay local work on top, then check those commits against what landed upstream before pushing.${dirty}"
fi
