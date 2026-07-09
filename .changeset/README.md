# Changesets

This folder holds **changeset fragments** — small markdown files that describe a
user-facing change before it ships. They are compiled into [`CHANGELOG.md`](../CHANGELOG.md)
and GitHub Release notes automatically when a release is cut (see
[Release Strategy](../docs/RELEASE-STRATEGY.md#changelog-generation)).

> This repo is Windows/.NET-first (MCP Server, CLI, VS Code Extension, MCPB). The
> `@changesets/cli` tool is Node-based, but it is used **only** to manage
> `CHANGELOG.md` content — it does not build, version, or publish any ExcelMcp
> component. `package.json` at the repo root exists solely to host this tool.

## Add a changeset for your PR

After making a user-facing change, run from the repo root:

```powershell
npx changeset
```

This asks two questions:

1. **Bump type** (major/minor/patch) — pick whichever feels right; it's metadata
   only and does not control the actual ExcelMcp release version (that's chosen
   by the maintainer when running the release workflow).
2. **Summary** — write 1-3 sentences **for end users**, not for other engineers.

Commit the generated `.changeset/<random-name>.md` file with your PR.

### What makes a good entry

The changelog is end-user facing. Favor plain language over implementation detail.

```md
✅ Good:
**Faster range reads** (#123): `range get-values` now batches COM calls,
cutting read time for large ranges by roughly half.

❌ Too technical / internal:
Refactored ComUtilities.ForEach to use a shared iterator and reduced
Marshal.ReleaseComObject calls in RangeCommands.GetValues by batching
Range.Value2 access via a single COM round-trip instead of per-cell reads.
```

Keep root-cause analysis, stack traces, and implementation notes in the PR
description or commit — not in the changeset summary.

### Which PRs need one

Any PR that changes user-visible behavior of the MCP Server, CLI, VS Code
Extension, or MCPB needs a changeset. Docs-only, test-only, CI-only, and
dependency-bump PRs generally don't — add the `skip-changelog` label instead of
a changeset (a CI check enforces this; see
`.github/workflows/changeset-check.yml`).

### Multiple changes in one PR

Add multiple changeset files if a PR bundles unrelated user-facing changes —
each renders as its own bullet.

## Full changesets documentation

- [Common questions](https://github.com/changesets/changesets/blob/main/docs/common-questions.md)
- [Main repository](https://github.com/changesets/changesets)
