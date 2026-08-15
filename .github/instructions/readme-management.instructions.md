---
applyTo: "**/*.md,README.md,**/README.md,**/index.md"
---

# README Management - Quick Reference

## Core Entry Documents

| File | Target Lines | Audience | Purpose |
|------|-------|----------|---------|
| `/README.md` | 120-150 | GitHub visitors | Acquisition page, requirements, quick start, and website routing |
| `/src/ExcelMcp.McpServer/README.md` | 80-100 | .NET devs | Concise NuGet gateway |
| `/src/ExcelMcp.CLI/README.md` | 180-200 | CLI users | CLI installation, syntax, and workflows |
| `/vscode-extension/README.md` | 100-120 | VS Code users | User benefits focus |
| `/gh-pages/docs/index.md` | 180-220 | Website visitors | Website landing page and task routing |

## Feature Documentation

`FEATURES.md` is the repository-level navigation hub. Detailed canonical feature
documentation lives in the four category files under `docs/features/`. Update
the appropriate category file first whenever a feature changes, then update the
hub if its navigation or tool-selection guidance also changed.

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

The canonical `docs/features/*.md` files contain the operation lists rendered on
the website. When adding/removing operations:

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
- [ ] `docs/features/*.md` section headers match their operation lists and sum to the canonical count

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

### Canonical Source First and Content Preservation

**ALWAYS improve the canonical repository document before changing its GitHub Pages
presentation.** Never create a website-only information architecture for content that has a
canonical source elsewhere in the repository.

**Never lose substantive information while shortening or moving documentation.**
Before removing content from an entry page such as `README.md`, map every
behavioral note, caveat, command example, installation option, architecture
detail, and workflow to a canonical destination. Add or improve that destination
first, then shorten the entry page.

Required order:

1. Restructure or improve the canonical source (`FEATURES.md`, `docs/features/`, component
   README, installation guide, etc.).
2. Compare the before/after content and record where each substantive section now lives.
3. Keep detailed content in that canonical source only.
4. Make `gh-pages/docs/` pages thin presentation wrappers that add navigation, SEO metadata,
   images, or Material components without becoming a second source of truth.
5. Update `gh-pages/hooks.py`, site navigation, workflow path filters, and documentation
   validation scripts to consume the new canonical structure.

For example, split an overly long feature reference under `docs/features/` first, then expose
those canonical files as focused website pages. Do not split only the rendered website while
leaving the canonical reference monolithic.

The site is built with **MkDocs Material** (see `gh-pages/mkdocs.yml`). It preserves a
single-source-of-truth pipeline: canonical repo files (READMEs, FEATURES.md, etc.) are
transformed at build time and pulled into thin wrapper pages — you never hand-copy content.

### Pattern: `hooks.py` and `pymdownx.snippets`

Thin wrappers under `gh-pages/docs/` render canonical repository sources through
the build hook and snippet extension:

1. **Canonical source:** The authoritative file under `docs/`, `src/`, or the
   repository root.
2. **Build hook:** `gh-pages/hooks.py` strips package headers where needed,
   rewrites repo-relative links through `SITE_PAGE_MAP`, adds stable feature
   anchors, and writes gitignored files under `gh-pages/docs/_generated/`.
3. **Wrapper:** A page under `gh-pages/docs/` contains SEO front matter, one H1,
   optional presentation content, and a snippet:

   ```markdown
   --8<-- "_generated/mcp-server.md"
   ```

4. **Result:** The website gets clean navigation and metadata while detailed
   content remains in one canonical repository file.

### Current Local Pages

| URL | Source | Page File |
|-----|--------|-----------|
| `/features/` | `/FEATURES.md` | `gh-pages/docs/features.md` |
| `/features/data-analytics/` | `/docs/features/DATA-ANALYTICS.md` | `gh-pages/docs/features/data-analytics.md` |
| `/features/cells-workbooks/` | `/docs/features/CELLS-WORKBOOKS.md` | `gh-pages/docs/features/cells-workbooks.md` |
| `/features/charts-visuals/` | `/docs/features/CHARTS-VISUALS.md` | `gh-pages/docs/features/charts-visuals.md` |
| `/features/automation-advanced/` | `/docs/features/AUTOMATION-ADVANCED.md` | `gh-pages/docs/features/automation-advanced.md` |
| `/guides/` | `/docs/guides/README.md` | `gh-pages/docs/guides/index.md` |
| `/guides/<slug>/` | `/docs/guides/*.md` (5 task guides) | `gh-pages/docs/guides/<slug>.md` |
| `/reference/<slug>/` | `/skills/shared/*.md` (24 agent reference files) | `gh-pages/docs/reference/<slug>.md` |
| `/installation/` | `/docs/INSTALLATION.md` | `gh-pages/docs/installation.md` |
| `/installation-mcp-server/` | `/docs/INSTALLATION-MCP-SERVER.md` | `gh-pages/docs/installation-mcp-server.md` |
| `/installation-cli/` | `/docs/INSTALLATION-CLI.md` | `gh-pages/docs/installation-cli.md` |
| `/architecture/` | `/docs/ARCHITECTURE.md` | `gh-pages/docs/architecture.md` |
| `/use-cases/` | `/docs/USE-CASES.md` | `gh-pages/docs/use-cases.md` |
| `/changelog/` | `/CHANGELOG.md` | `gh-pages/docs/changelog.md` |
| `/mcp-server/` | `/src/ExcelMcp.McpServer/README.md` | `gh-pages/docs/mcp-server.md` |
| `/cli/` | `/src/ExcelMcp.CLI/README.md` | `gh-pages/docs/cli.md` |
| `/skills/` | `/skills/README.md` | `gh-pages/docs/skills.md` |
| `/contributing/` | `/docs/CONTRIBUTING.md` | `gh-pages/docs/contributing.md` |
| `/security/` | `/SECURITY.md` | `gh-pages/docs/security.md` |
| `/privacy/` | `/PRIVACY.md` | `gh-pages/docs/privacy.md` |

### Adding New Local Pages

1. **Create or update the canonical source** under `docs/`, `src/`, or the
   repository root.
2. **Register it in `gh-pages/hooks.py`** with `SITE_PAGE_MAP` and an `_write()`
   call using the appropriate transformation.
3. **Create a wrapper** in `gh-pages/docs/` with SEO front matter and a snippet:
   ```markdown
   ---
   title: Page Title
   description: One-sentence SEO description.
   keywords: relevant, keywords
   ---

   # Page Title

   --8<-- "_generated/page-name.md"
   ```

4. **Add it to `nav:`** in `gh-pages/mkdocs.yml`.
5. **Update site links** to use the local URL instead of a GitHub URL.
6. **Build strictly** from `gh-pages/`: `python -m mkdocs build --strict --clean`.

### Generated machine-readable outputs

The site build also derives these from the same page content, so they never need
manual maintenance and cannot drift:

| Output | Contents |
|---|---|
| `/llms.txt` | llmstxt.org index: every page with its description, in nav order |
| `/llms-full.txt` | Full Markdown of every page, `--8<--` includes resolved |
| `/<path>/index.md` | Markdown mirror of each page, advertised via `<link rel="alternate">` |
| `/tools.json` | Tool and operation catalogue parsed from `docs/features/*.md` |
| `FAQPage` JSON-LD | Built from the `??? question` blocks on the troubleshooting page |

`tools.json` generation **fails the build** if the operation totals in
`docs/features/*.md` stop matching the headline count in `FEATURES.md`.


### Why Local Pages

- **Consistent UX** - All docs served from same domain
- **Single source of truth** - Content pulled directly from canonical source files at build time
- **SEO** - Better for search engine indexing
- **Offline docs** - Works with `mkdocs serve` locally
