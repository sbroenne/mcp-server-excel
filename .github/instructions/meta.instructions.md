---
applyTo: ".github/**/*.md,.github/instructions/**"
---

# Meta Instructions - How to Create GitHub Copilot Instructions

## Official Documentation

**Primary Reference:** [GitHub Docs - Adding Repository Custom Instructions](https://docs.github.com/en/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot)

## Instruction File Types

GitHub Copilot supports three types of repository custom instructions:

### 1. Repository-Wide Instructions
- **File:** `.github/copilot-instructions.md`
- **Scope:** Applies to ALL requests in the repository
- **Purpose:** Project overview, general guidelines, architecture
- **No frontmatter required**

### 2. Path-Specific Instructions
- **Files:** `.github/instructions/NAME.instructions.md`
- **Naming:** MUST end with `.instructions.md`
- **Scope:** Applies to specific files/paths via glob patterns
- **Required frontmatter:** `applyTo` property

**Example:**
```markdown
---
applyTo: "**/*.cs,**/*.csproj,util/**"
---

# C# Development Guidelines
...
```

### 3. Agent Instructions
- **Files:** `AGENTS.md`, `CLAUDE.md`, or `GEMINI.md`
- **Location:** Anywhere in repository (nearest file in tree takes precedence)
- **Purpose:** AI agent-specific instructions

## Path-Specific Instructions - Critical Rules

### Naming Convention
- ✅ **CORRECT:** `powerquery.instructions.md`, `csharp.instructions.md`, `api.instructions.md`
- ❌ **WRONG:** `powerquery.md`, `copilot-powerquery.md`, `instructions-powerquery.md`

### Directory Structure
```
.github/
├── copilot-instructions.md          # Repository-wide
└── instructions/                     # Path-specific files
    ├── powerquery.instructions.md
    ├── csharp.instructions.md
    ├── powerbi.instructions.md
    └── excel.instructions.md
```

### Frontmatter Requirements

**Required format:**
```markdown
---
applyTo: "glob,patterns,here"
---
```

**Glob Pattern Examples:**
```markdown
# Single pattern
applyTo: "**/*.cs"

# Multiple patterns (comma-separated)
applyTo: "**/*.cs,**/*.csproj,util/**"

# All files
applyTo: "**"

# Specific directories
applyTo: "src/**,tests/**"

# Multiple file types
applyTo: "**/*.ts,**/*.tsx,**/*.js,**/*.jsx"
```

**Pattern Matching:**
- Use `**` for recursive directory matching
- Use `*` for single-level wildcard
- Use `,` to separate multiple patterns (NO spaces after commas)
- Patterns are case-sensitive on case-sensitive file systems

## Automatic Behavior

### When Instructions Apply
GitHub Copilot **automatically** includes instructions when:
- You're working on a file matching an `applyTo` pattern
- You reference the repository in Copilot Chat
- You use Copilot in the context of the repository

### No Cross-References Needed
- ❌ **Don't do this:** Link between instruction files
- ✅ **Copilot handles it:** Files are loaded automatically based on context

**Example of what NOT to do:**
```markdown
<!-- WRONG - Don't add cross-references -->
For Power Query guidelines, see [powerquery.instructions.md](powerquery.instructions.md)
```

**Why:** Copilot already loads `powerquery.instructions.md` when you work on `.pq` files.

## Content Guidelines

### Repository-Wide Instructions (`copilot-instructions.md`)
**Should include:**
- Project overview and purpose
- Technology stack summary
- Directory structure
- Build/test/run commands
- General coding principles
- Key architectural patterns

**Should NOT include:**
- Technology-specific deep dives (use path-specific instead)
- Redundant cross-references to other instruction files

### Path-Specific Instructions (`*.instructions.md`)
**Should include:**
- Language/framework-specific patterns
- Technology best practices
- Common code snippets and examples
- Error handling patterns
- Performance optimization tips
- Tool-specific workflows

**Should NOT include:**
- Project-wide architecture (use copilot-instructions.md)
- Cross-references to other instruction files

## Best Practices

### 1. Single Responsibility
Each instruction file should focus on ONE technology or concern:
- `csharp.instructions.md` → C# language patterns
- `powerquery.instructions.md` → Power Query M language
- `testing.instructions.md` → Testing strategies

### 2. Precise applyTo Patterns
Make patterns specific enough to avoid conflicts:
```markdown
# Good - Specific and clear
applyTo: "**/*.pq,**/expressions.tmdl,powerquery/**"

# Avoid - Too broad, may conflict with other instructions
applyTo: "**"
```

### 3. Combine Related Patterns
Group related file types in one instruction file:
```markdown
---
applyTo: "**/*.cs,**/*.csproj,**/*.sln,util/**"
---
```

### 4. Use Descriptive Names
Choose instruction file names that clearly indicate their purpose:
- ✅ `api-development.instructions.md`
- ✅ `database-migrations.instructions.md`
- ❌ `stuff.instructions.md`
- ❌ `misc.instructions.md`

## Testing Instructions

### Verify Instructions Are Loaded
1. Open Copilot Chat
2. Ask a question related to your instruction files
3. Check "References" in the response
4. Look for `.github/copilot-instructions.md` or `*.instructions.md` files

### Debug applyTo Patterns
If instructions aren't loading:
1. Verify file ends with `.instructions.md`
2. Check frontmatter syntax (YAML format)
3. Test glob pattern matches your file paths
4. Ensure no syntax errors in frontmatter

### Common Issues

**Issue:** Instructions not loading
```markdown
# WRONG - Missing .instructions.md suffix
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
# WRONG - Spaces after commas break patterns
applyTo: "**/*.cs, **/*.csproj"

# CORRECT - No spaces
applyTo: "**/*.cs,**/*.csproj"
```

## Example: Complete Path-Specific Instruction File

```markdown
---
applyTo: "**/*.pq,**/expressions.tmdl,powerquery/**"
---

# Power Query M - Development Patterns

## Overview
Power Query M language usage in this project.

## Best Practices

### Data Source References
- Use `Excel.CurrentWorkbook(){[Name="RangeName"]}[Content]`
- Reference queries directly by name

### Function Design
- Define typed parameters
- Use descriptive step names
- Handle errors with try...otherwise

## Common Patterns

### Load Data
\`\`\`powerquery
Source = Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
\`\`\`

### Transform Data
\`\`\`powerquery
Filtered = Table.SelectRows(Source, each [Amount] > 0)
\`\`\`
```

## Migration from Old Format

If you have files without `.instructions.md`:

### Step 1: Rename Files
```powershell
# Old format
.github/copilot-powerquery.md

# New format
.github/instructions/powerquery.instructions.md
```

### Step 2: Add Frontmatter
```markdown
---
applyTo: "**/*.pq,powerquery/**"
---

# (existing content)
```

### Step 3: Remove Cross-References
Delete any links between instruction files - Copilot loads them automatically.

### Step 4: Update Main Instructions
Remove references to technology-specific files from `copilot-instructions.md`.

## Tools and Validation

### Check File Structure
```powershell
# List all instruction files
Get-ChildItem ".github" -Recurse -Filter "*.instructions.md"

# Verify naming convention
Get-ChildItem ".github/instructions" | Where-Object { $_.Name -notlike "*.instructions.md" }
```

### Validate Frontmatter
```powershell
# Check for frontmatter in instruction files
Select-String -Path ".github/instructions/*.instructions.md" -Pattern "^---$" -Context 0,5
```

## References

- [GitHub Docs - Repository Custom Instructions](https://docs.github.com/en/copilot/customizing-copilot/adding-custom-instructions-for-github-copilot)
- [Glob Pattern Reference](https://en.wikipedia.org/wiki/Glob_(programming))
- [Custom Instructions Examples](https://docs.github.com/en/copilot/tutorials/customization-library/custom-instructions)
