---
applyTo: "**/*.md,README.md,**/README.md, **/index.md"
---

# README Management - Quick Reference

## Four READMEs

| File | Lines | Audience | Purpose |
|------|-------|----------|---------|
| `/README.md` | 250-300 | All users | Comprehensive reference |
| `/src/ExcelMcp.McpServer/README.md` | 80-100 | .NET devs | Concise NuGet gateway |
| `/vscode-extension/README.md` | 100-120 | VS Code users | User benefits focus |
| `/gh-pages/index.md` | 450-5000 | All users | Comprehensive reference |

## Features.md

You need to make sure that the `Features.md` file is up-to-date with the latest features of the project. This file should be updated whenever a new feature is added or an existing feature is modified.

## Critical Rules

### Tool & Action Counts Must Match

**IMPORTANT: CLI has FEWER tools/operations than MCP Server!**

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
 - `/gh-pages/index.md`
 - `/FEATURES.md`

### Operation Lists Must Be Complete

**IMPORTANT: Where operation lists are documented, they MUST match the actual code!**

The `gh-pages/index.md` file contains detailed tables of all operations for each tool. When adding/removing operations:

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
- [ ] gh-pages/index.md section headers match table row counts

## Common Mistakes

## CHANGELOG.md

The project uses a **centralized changelog** at `/CHANGELOG.md` covering all components.

**When to update:**
- Before creating a `v*` tag, ensure the version section exists in CHANGELOG.md
- The release workflow extracts the specific version's changes for release notes
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
| Missing CHANGELOG entry | Add before creating release tag |
