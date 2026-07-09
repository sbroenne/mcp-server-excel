---
applyTo: "**/*.md,README.md,**/README.md,**/index.md"
---

# README Management - Quick Reference

## Four READMEs

| File | Lines | Audience | Purpose |
|------|-------|----------|---------|
| `/README.md` | 250-300 | All users | Comprehensive reference |
| `/src/ExcelMcp.McpServer/README.md` | 80-100 | .NET devs | Concise NuGet gateway |
| `/vscode-extension/README.md` | 100-120 | VS Code users | User benefits focus |
| `/gh-pages/docs/index.md` | 450-5000 | All users | Comprehensive reference |

## Features.md

You need to make sure that the `Features.md` file is up-to-date with the latest features of the project. This file should be updated whenever a new feature is added or an existing feature is modified.

## Critical Rules

### Tool & Action Counts Must Match

**⚠️ IMPORTANT: CLI has FEWER tools/operations than MCP Server!**

**ALWAYS count tools/operations BEFORE updating any README. Never use hardcoded numbers from memory.**

Before updating counts, verify by counting:

- **MCP Server**: Count tool files (excel_batch handled via ExcelTools.cs, not separate tool file)
- **CLI**: Count command group folders (includes Session commands)
- **Operations**: Count separately for each - they differ!

Sync counts across:
  - GitHub Project About: https://github.com/sbroenne/mcp-server-excel (use the GitHub CLI to update)
  - `/README.md`
  - `/src/ExcelMcp.McpServer/README.md`
  - `/src/ExcelMcp.CLI/README.md`
  - `/vscode-extension/README.md`
  - `/gh-pages/docs/index.md`
  - `/FEATURES.md`

### Operation Lists Must Be Complete

**⚠️ IMPORTANT: Where operation lists are documented, they MUST match the actual code!**

The `gh-pages/docs/index.md` file contains detailed tables of all operations for each tool. When adding/removing operations:

1. **Verify section header count** matches actual operation count in code
2. **Verify each operation is listed** in the table - no missing or extra entries
3. **Verify operation names** match the code (kebab-case in docs, PascalCase in code)

**Common discrepancies found:**
- Section header says "25 actions" but code has 30
- Table lists operations that don't exist (stale documentation)
- Table is missing newly added operations

### Version Numbers
- **NEVER** manually update versions in README files
- Versions auto-managed by release workflow
- See `docs/RELEASE-STRATEGY.md`

## Verification Checklist

Before committing README changes:

- [ ] Tool counts match actual code (count, don't assume)
- [ ] Operation counts match actual code per tool
- [ ] Operation LISTS in tables match actual code (no missing/extra entries)
- [ ] All READMEs updated (not just one)
- [ ] FEATURES.md updated if applicable
- [ ] gh-pages/docs/index.md section headers match table row counts

## Common Mistakes

## CHANGELOG.md

The project uses a **centralized changelog** at `/CHANGELOG.md` covering all components. It is generated from [changesets](../../.changeset/README.md), not hand-edited — see `docs/RELEASE-STRATEGY.md#changelog-generation` and Rule 27.

**When to update:**
- Add a changeset (`npx changeset`) with your PR, not by editing CHANGELOG.md directly
- The release workflow (`scripts/Build-Changelog.ps1`) compiles pending changesets into a new version section and uses it verbatim for release notes
- Uses standard Keep a Changelog format: `## [version] - YYYY-MM-DD`

| Mistake | Fix |
|---------|-----|
| Duplicate tool entries | List each tool once |
| Unverified action counts | Count actual switch cases in code |
| Incomplete operation lists | Compare each table row against code |
| Stale operation names | Operations get renamed - verify current names |
| Overclaiming features | Use actual counts, not estimates |
| Missing safety callout | Add COM API benefits |
| Manual version updates | Let workflow handle it |
| Missing changeset | Add via `npx changeset` before merging (CI enforces this) |
| Hand-editing CHANGELOG.md directly | Add a changeset fragment instead — it's compiled automatically |
| External GitHub links in gh-pages | Use local pages (see gh-pages pattern below) |

## gh-pages Local Documentation Pattern

**CRITICAL: All documentation in gh-pages should use LOCAL pages, NOT external GitHub links.**

The site is built with **MkDocs Material** (see `gh-pages/mkdocs.yml`). It preserves a
single-source-of-truth pipeline: canonical repo files (READMEs, FEATURES.md, etc.) are
transformed at build time and pulled into thin wrapper pages — you never hand-copy content.

### Pattern: MkDocs hook + snippet include

`gh-pages/hooks.py` runs on `on_pre_build` and generates `gh-pages/docs/_generated/*.md`
from the canonical sources (stripping the H1/badges, demoting headings, and rewriting
repo-relative links to site/GitHub URLs). Thin wrapper pages then embed the generated
file via a `pymdownx.snippets` include:

1. **Source file** (e.g., `src/ExcelMcp.McpServer/README.md`)
2. **`hooks.py` generates** `docs/_generated/mcp-server.md` (H1/badges stripped, links rewritten)
3. **Page file** (e.g., `gh-pages/docs/mcp-server.md`) includes it: `--8<-- "_generated/mcp-server.md"`
4. **Result**: Local URL `/mcp-server/` instead of GitHub link

### Current Local Pages

| URL | Source | Page File |
|-----|--------|-----------|
| `/features/` | `/FEATURES.md` | `gh-pages/docs/features.md` |
| `/installation/` | `/docs/INSTALLATION.md` | `gh-pages/docs/installation.md` |
| `/changelog/` | `/CHANGELOG.md` | `gh-pages/docs/changelog.md` |
| `/mcp-server/` | `/src/ExcelMcp.McpServer/README.md` | `gh-pages/docs/mcp-server.md` |
| `/cli/` | `/src/ExcelMcp.CLI/README.md` | `gh-pages/docs/cli.md` |
| `/skills/` | `/skills/README.md` | `gh-pages/docs/skills.md` |
| `/contributing/` | `/docs/CONTRIBUTING.md` | `gh-pages/docs/contributing.md` |
| `/security/` | `/docs/SECURITY.md` | `gh-pages/docs/security.md` |
| `/privacy/` | `/PRIVACY.md` | `gh-pages/docs/privacy.md` |

### Adding New Local Pages

1. **Update `gh-pages/hooks.py`** — add a `_write("target.md", "path/to/SOURCE.md", ...)`
   call inside `on_pre_build` (reuse `_strip_header` for demote/strip behavior, or write a
   verbatim copy for pages that keep their own H1).

2. **Create page file** in `gh-pages/docs/` with SEO front matter and a snippet include:
   ```markdown
   ---
   title: Page Title
   description: One-sentence SEO description.
   keywords: relevant, keywords
   ---

   # Page Title

   --8<-- "_generated/target.md"
   ```

3. **Add it to `nav:`** in `gh-pages/mkdocs.yml` so it appears in site navigation.

4. **Update `index.md`** — use the local URL `/url-path/` instead of a GitHub link.

### Why Local Pages

- **Consistent UX** - All docs served from same domain
- **Single source of truth** - Content auto-synced from source files via `hooks.py`
- **SEO** - Better for search engine indexing
- **Offline docs** - Works with `mkdocs serve` locally
