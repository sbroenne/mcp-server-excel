---
applyTo: "**/*.md,README.md,**/README.md"
---

# README Management - Quick Reference

## Three READMEs

| File | Lines | Audience | Purpose |
|------|-------|----------|---------|
| `/README.md` | 250-300 | All users | Comprehensive reference |
| `/src/ExcelMcp.McpServer/README.md` | 80-100 | .NET devs | Concise NuGet gateway |
| `/vscode-extension/README.md` | 100-120 | VS Code users | User benefits focus |

## Critical Rules

### Tool Counts Must Match
### Tool & Action Counts Must Match
- All READMEs: **11 specialized tools**
- Current total actions: **154 operations** (update if tool schema changes)
- Sync counts across:
	- `/README.md`
	- `/src/ExcelMcp.McpServer/README.md`
	- `/src/ExcelMcp.CLI/README.md`
	- `/vscode-extension/README.md`
	- `/gh-pages/index.md`
- Verify against code: `git grep "case.*:" src/ExcelMcp.McpServer/Tools/`

# Check READMEs/sites have same tool count
git grep "11 specialized tools" README.md src/ExcelMcp.McpServer/README.md src/ExcelMcp.CLI/README.md vscode-extension/README.md gh-pages/index.md

# Check total operations text (154)
- All READMEs mention: **COM API** (not "Excel's internal API")
- Highlight: zero corruption, interactive development, growing features

### Version Numbers
- **NEVER** manually update versions in README files
- Versions auto-managed by release workflow
- See `docs/RELEASE-STRATEGY.md`

## Before Committing README Changes

```bash
# Verify tool counts match code
git grep -A 100 "switch.*action" src/ExcelMcp.McpServer/Tools/*.cs | grep "case" | wc -l

# Check all 3 READMEs have same tool count
git grep "12 specialized tools" README.md src/ExcelMcp.McpServer/README.md vscode-extension/README.md
```

## Common Mistakes

| Mistake | Fix |
|---------|-----|
| Duplicate tool entries | List each tool once |
| Unverified action counts | Count actual switch cases in code |
| Overclaiming features | Use "80+ operations" not "all features" |
| Missing safety callout | Add COM API benefits |
| Manual version updates | Let workflow handle it |

**Full details**: See old backup if needed
