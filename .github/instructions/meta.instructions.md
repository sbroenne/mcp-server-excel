---
applyTo: ".github/**/*.md,.github/instructions/**"
---

# Meta Instructions - GitHub Copilot Instructions

**Reference:** [GitHub Docs - Repository Custom Instructions](https://docs.github.com/en/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot)

## Instruction File Types

1. **Repository-Wide:** `.github/copilot-instructions.md` - Applies to ALL requests (no frontmatter)
2. **Path-Specific:** `.github/instructions/NAME.instructions.md` - Applies to glob patterns (requires frontmatter)
3. **Agent-Specific:** `AGENTS.md`, `CLAUDE.md`, `GEMINI.md` - AI agent-specific (anywhere in tree)

## Path-Specific Instructions - Critical Rules

### Naming Convention
- ✅ **CORRECT:** `powerquery.instructions.md`, `csharp.instructions.md`
- ❌ **WRONG:** `powerquery.md`, `copilot-powerquery.md`

### Frontmatter Requirements
```markdown
---
applyTo: "glob,patterns,here"
---
```

**Glob Examples:**
```markdown
applyTo: "**/*.cs"                           # Single pattern
applyTo: "**/*.cs,**/*.csproj,util/**"       # Multiple (NO spaces after commas)
applyTo: "**"                                # All files
```

**Pattern Rules:**
- `**` = recursive directories
- `*` = single-level wildcard
- `,` = separator (NO spaces)
- Case-sensitive on case-sensitive filesystems

## Automatic Behavior

Copilot **automatically** loads instructions when working on matching files. NO cross-references needed.

**❌ Don't do this:**
```markdown
For Power Query, see [powerquery.instructions.md](powerquery.instructions.md)
```

**Why:** Copilot loads `powerquery.instructions.md` automatically when working on `.pq` files.

## Content Guidelines

### Repository-Wide (`copilot-instructions.md`)
**Include:** Project overview, tech stack, directory structure, build commands, general principles  
**Exclude:** Technology deep dives (use path-specific), cross-references

### Path-Specific (`*.instructions.md`)
**Include:** Language/framework patterns, best practices, code snippets, error handling, tool workflows  
**Exclude:** Project-wide architecture, cross-references

## Best Practices

1. **Single Responsibility** - One technology per file
2. **Precise Patterns** - Avoid conflicts (use specific patterns, not `**`)
3. **Combine Related** - Group related file types
4. **Descriptive Names** - Clear purpose (`api-development.instructions.md` ✅, `stuff.instructions.md` ❌)

## Testing Instructions

### Verify Loading
1. Open Copilot Chat
2. Ask related question
3. Check "References" for instruction files

### Debug Common Issues

**Issue:** Instructions not loading
```markdown
# WRONG - Missing suffix
.github/instructions/powerquery.md

# CORRECT
.github/instructions/powerquery.instructions.md
```

**Issue:** Invalid frontmatter
```markdown
# WRONG - Missing closing ---
---
applyTo: "**/*.cs"

# CORRECT
---
applyTo: "**/*.cs"
---
```

**Issue:** Pattern doesn't match
```markdown
# WRONG - Spaces break patterns
applyTo: "**/*.cs, **/*.csproj"

# CORRECT - No spaces
applyTo: "**/*.cs,**/*.csproj"
```

## References

- [GitHub Docs - Repository Custom Instructions](https://docs.github.com/en/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot)
- [Glob Pattern Reference](https://en.wikipedia.org/wiki/Glob_(programming))
