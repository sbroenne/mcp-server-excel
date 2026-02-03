---
applyTo: ".github/**/*.md,.github/instructions/**"
---

# Copilot Instructions - Meta

> **Note for maintainers**: This file provides minimal guidance for editing instruction files. For complete documentation, see `docs/COPILOT-INSTRUCTIONS-GUIDE.md`

## Quick Rules for Maintainers

### File Naming
- ✅ MUST end with `.instructions.md`
- ✅ Use descriptive names: `excel-com-interop.instructions.md`
- ❌ DON'T use generic names: `notes.instructions.md`

### Frontmatter Required
```markdown
---
applyTo: "**/*.cs,**/*.csproj"
---
```
- NO spaces after commas in glob patterns
- Use `**` for recursive directories

### Content Focus
**Write FOR the LLM** (actionable instructions):
- ✅ "When X, run command Y"
- ✅ "Follow these N steps"
- ✅ Quick reference tables
- ✅ ✅/❌ code examples

**Don't write ABOUT the system** (documentation):
- ❌ "The problem we solved was..."
- ❌ "This strategy has 5 layers..."
- ❌ Historical context

### Testing
1. Open file matching `applyTo` pattern
2. Ask Copilot Chat a related question
3. Check "References" section for your instruction file

**Full Guide**: See `docs/COPILOT-INSTRUCTIONS-GUIDE.md` for complete patterns, best practices, and examples.

