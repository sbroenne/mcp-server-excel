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

### Pattern: MkDocs `include-markdown` plugin

Thin wrapper pages under `gh-pages/docs/` pull content directly from canonical repo
sources using the `mkdocs-include-markdown-plugin`'s `{% include-markdown %}` directive
— no build-time generation script, no gitignored intermediate files. The plugin supports
`start`/`end` HTML-comment markers (to skip a source file's title/badges block) and
`heading-offset` (to demote headings so the wrapper's own `# Title` stays the only H1):

1. **Source file** (e.g., `src/ExcelMcp.McpServer/README.md`) — has a `<!--start-->`
   marker after its title/badge block, and an `<!--end-->` marker before any
   package-specific footer (e.g. "Additional Resources") that would duplicate site nav.
2. **Page file** (e.g., `gh-pages/docs/mcp-server.md`) includes it directly:
   ```jinja
   {%
     include-markdown "../../src/ExcelMcp.McpServer/README.md"
     start="<!--start-->"
     end="<!--end-->"
     heading-offset=1
   %}
   ```
3. **Result**: Local URL `/mcp-server/` renders the source file's content with its own
   H1 replaced by the wrapper's, and no duplicated footer nav.

Pages that are meant to be **verbatim copies** (their own H1 becomes the page title —
`contributing.md`, `security.md`, `privacy.md`) use the directive with no `start`/`end`/
`heading-offset` at all — the whole source file is included as-is.

### Current Local Pages

| URL | Source | Page File |
|-----|--------|-----------|
| `/features/` | `/FEATURES.md` | `gh-pages/docs/features.md` |
| `/installation/` | `/docs/INSTALLATION.md` | `gh-pages/docs/installation.md` |
| `/installation-mcp-server/` | `/docs/INSTALLATION-MCP-SERVER.md` | `gh-pages/docs/installation-mcp-server.md` |
| `/installation-cli/` | `/docs/INSTALLATION-CLI.md` | `gh-pages/docs/installation-cli.md` |
| `/changelog/` | `/CHANGELOG.md` | `gh-pages/docs/changelog.md` |
| `/mcp-server/` | `/src/ExcelMcp.McpServer/README.md` | `gh-pages/docs/mcp-server.md` |
| `/cli/` | `/src/ExcelMcp.CLI/README.md` | `gh-pages/docs/cli.md` |
| `/skills/` | `/skills/README.md` | `gh-pages/docs/skills.md` |
| `/contributing/` | `/docs/CONTRIBUTING.md` | `gh-pages/docs/contributing.md` |
| `/security/` | `/SECURITY.md` | `gh-pages/docs/security.md` |
| `/privacy/` | `/PRIVACY.md` | `gh-pages/docs/privacy.md` |

### Adding New Local Pages

1. **Add markers to the source file** (if it needs its title/badges stripped or a
   footer excluded): a `<!--start-->` comment after the header block, and/or an
   `<!--end-->` comment before any footer that duplicates site nav.

2. **Create page file** in `gh-pages/docs/` with SEO front matter and an
   `include-markdown` directive:
   ```markdown
   ---
   title: Page Title
   description: One-sentence SEO description.
   keywords: relevant, keywords
   ---

   # Page Title

   {%
     include-markdown "../../path/to/SOURCE.md"
     start="<!--start-->"
     end="<!--end-->"
     heading-offset=1
   %}
   ```

3. **Add it to `nav:`** in `gh-pages/mkdocs.yml` so it appears in site navigation.

4. **Update `index.md`** — use the local URL `/url-path/` instead of a GitHub link.

### Why Local Pages

- **Consistent UX** - All docs served from same domain
- **Single source of truth** - Content pulled directly from canonical source files at build time
- **SEO** - Better for search engine indexing
- **Offline docs** - Works with `mkdocs serve` locally
