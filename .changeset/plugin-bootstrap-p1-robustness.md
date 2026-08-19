---
"excelmcp": patch
---

Harden the plugin runtime bootstrap against corrupt caches, concurrent installs, and rate limits.

Follow-up to the offline-fallback fix. `download.ps1` for both `excel-cli` and `excel-mcp` now:

- **Validates the cached archive instead of merely testing for its presence.** A truncated
  download used to wedge the plugin permanently: the release tag still matched, so no download was
  attempted, yet extraction failed on every subsequent run. Recovery required manually deleting
  the cache. The archive is now opened and checked before it is trusted.
- **Downloads to a temp file and renames into place**, so an interrupted transfer can never leave
  a partial archive that a later run mistakes for a complete one.
- **Extracts into a staging directory and swaps it in**, rather than deleting the release
  directory first. This also fixes a destructive failure mode: `Remove-Item -Recurse` deletes every
  sibling file before it reaches a locked executable and fails, leaving a half-destroyed install.
  The executable is now probed for a lock *before* anything is removed, and a runtime that is
  currently in use is kept rather than partially overwritten.
- **Retries once with a fresh download** if an install fails, instead of failing permanently.
- **Serializes installs with a named mutex**, so concurrent sessions cannot race on the same
  archive and release directory.
- **Verifies the resolved runtime's version** from the version stamped into the file. Running the
  runtime with `--version` was deliberately avoided: it performs its own network update check,
  which is exactly wrong inside a bootstrap that must work offline.
- **Sends `GITHUB_TOKEN` / `GH_TOKEN` as a bearer token** when present. Unauthenticated GitHub API
  access is 60 requests/hour per source IP, a budget shared by everyone behind a corporate NAT and
  routinely exhausted — the most common cause of release metadata being unreachable.
- **Re-checks for updates on a time window for non-Copilot installs.** Outside a Copilot session
  the session id is the constant `"standalone"`, so it always equalled the previously recorded one
  and the freshness check never fired again. PATH and shim installs were pinned forever to
  whatever they first downloaded, despite the docs promising the newest runtime.
