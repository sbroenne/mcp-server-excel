---
applyTo: "**/*.md,README.md,**/README.md,src/ExcelMcp.McpServer/.mcp/server.json"
---

# README Management Instructions

> **Guidelines for maintaining ExcelMcp's three user-facing README files and MCP server metadata**

## 📋 Documentation Inventory

ExcelMcp has **three README files** and **one metadata file** that must stay synchronized:

| File | Location | Audience | Discovery Channel |
|------|----------|----------|-------------------|
| **Main README** | `/README.md` | All users | GitHub repo, Google search, npm/NuGet "homepage" link |
| **NuGet Package README** | `/src/ExcelMcp.McpServer/README.md` | .NET developers | NuGet.org package page, `dotnet tool install` search |
| **VS Code Extension README** | `/vscode-extension/README.md` | VS Code users | VS Code Marketplace, Extension search panel |
| **MCP Server Metadata** | `/src/ExcelMcp.McpServer/.mcp/server.json` | MCP registry | MCP.run registry, Claude Desktop, VS Code MCP discovery |

---

## 🎯 Purpose & Scope

### Main README (`/README.md`)

**Purpose:** Comprehensive project documentation and primary landing page

**Audience:**
- Developers evaluating the project (GitHub visitors)
- Users finding via Google search
- Contributors reviewing architecture
- Anyone clicking "Homepage" from NuGet or Marketplace

**Scope:** COMPREHENSIVE (272 lines)
- ✅ Full project description with Quick Example
- ✅ Complete safety explanation (3-point COM API benefits)
- ✅ Who Should Use This / Not Suitable For sections
- ✅ Detailed feature list (80+ operations) in collapsible `<details>`
- ✅ Quick Start with VS Code Extension AND manual installation
- ✅ Complete tool overview (11 tools with action counts)
- ✅ Links to all documentation (CLI, MCP Server, Installation)
- ✅ Contributing, license, acknowledgments, SEO keywords

**Content Strategy:** Be thorough. This is the definitive reference.

---

### NuGet Package README (`/src/ExcelMcp.McpServer/README.md`)

**Purpose:** Package description for NuGet.org search and package page

**Audience:**
- .NET developers searching for "MCP server" or "Excel automation"
- Users running `dotnet tool search excel`
- Developers evaluating NuGet packages

**Scope:** CONCISE GATEWAY (89 lines)
- ✅ One-line description + badges
- ✅ Brief safety callout (COM API advantage)
- ✅ Quick installation (global .NET tool)
- ✅ Brief tool list (11 tools with action counts only)
- ✅ 4 example use cases (1-2 lines each)
- ✅ Links to full GitHub documentation

**Content Strategy:** Focus on discoverability and conversion. Answer "What is this?" and "How do I install it?" then link to comprehensive docs.

**What NOT to Include:**
- ❌ Detailed examples or tutorials
- ❌ Complete feature lists (use brief bullets)
- ❌ Architecture explanations
- ❌ Long use case scenarios
- ❌ Duplicate content from main README

---

### VS Code Extension README (`/vscode-extension/README.md`)

**Purpose:** Extension description for VS Code Marketplace

**Audience:**
- VS Code users searching for "Excel" or "Copilot Excel"
- GitHub Copilot users looking for MCP servers
- Developers browsing Marketplace AI extensions

**Scope:** FOCUSED USER BENEFITS (105 lines)
- ✅ Natural language examples (5 bullet points)
- ✅ Quick Example workflow
- ✅ Safety callout (3-point COM API benefits)
- ✅ Who Should Use This / Not Suitable For
- ✅ Quick Start (3 steps)
- ✅ Requirements (Windows, Excel, .NET auto-installed)
- ✅ Tool list (11 tools with brief descriptions)
- ✅ Troubleshooting (3 common issues)
- ✅ Minimal documentation links (2 essential links)

**Content Strategy:** Match main README's user benefits while staying concise. VS Code users and GitHub visitors are the same audience.

**What NOT to Include:**
- ❌ Detailed example categories (already in opening)
- ❌ Verbose "Common Use Cases" scenarios
- ❌ "How It Works" technical explanations
- ❌ Multiple documentation links (main README link covers it)

---

### MCP Server Metadata (`/src/ExcelMcp.McpServer/.mcp/server.json`)

**Purpose:** Machine-readable metadata for MCP registry and discovery tools

**Audience:**
- MCP.run registry indexing bots
- Claude Desktop MCP server discovery
- VS Code MCP extension discovery
- Developers installing via MCP registry

**Scope:** MACHINE-READABLE METADATA (JSON)
- ✅ Server name (must match registry format: `io.github.sbroenne/mcp-server-excel`)
- ✅ Title (brief, human-readable: "Excel COM Automation")
- ✅ Description (one-line: tools listed, max 100 chars)
- ✅ Version (must match NuGet package version)
- ✅ NuGet package identifier and version
- ✅ Repository URL

**Content Strategy:** Keep synchronized with NuGet package version. Description should list key tools.

**Critical Fields:**
```json
{
  "name": "io.github.sbroenne/mcp-server-excel",
  "title": "Excel COM Automation",
  "description": "Excel COM automation - Power Query, DAX measures, VBA, Tables, ranges, connections",
  "version": "1.0.0",  // MUST match NuGet package version
  "packages": [
    {
      "identifier": "Sbroenne.ExcelMcp.McpServer",
      "version": "1.0.0"  // MUST match NuGet package version
    }
  ]
}
```

**What to Update When:**
- ✅ New version → Update `version` and `packages[0].version` to match NuGet
- ✅ New major tool → Update `description` to include tool name
- ✅ Repository moved → Update `repository.url`
- ❌ Don't change `name` (breaks MCP registry references)

---

## ✅ Critical Rules

### Rule 1: Consistency Across All Documentation

**Tool Counts Must Match:**
- All 3 READMEs must show **11 specialized tools**
- All 3 READMEs must show identical action counts per tool
- server.json description should mention key tools
- Example: `excel_datamodel` = **14 actions** (not 15, not 20)

**Version Numbers Must Match:**
- server.json `version` field = NuGet package version
- server.json `packages[0].version` = NuGet package version
- Update both simultaneously during releases

**Safety Messaging Must Match:**
- All READMEs must mention COM API (not "Excel's internal API")
- All READMEs must highlight "zero risk of document corruption"
- All READMEs must mention "interactive development" / "real-time"
- All READMEs must state "active development" / "growing feature set"

**Who Should Use This Must Match:**
- Perfect for: Data analysts, Developers, Business users, Teams
- Not suitable for: Linux/macOS, High-volume batch operations
- Main README also includes: "Server-side data processing" (with library references)

### Rule 2: Verify Against Code Before Claiming Features

**Before documenting tool action counts:**
1. Search for the tool's switch statement (e.g., `ExcelDataModelTool.cs`)
2. Count actual implemented actions in switch cases
3. Check RegularExpression attribute for allowed actions
4. Update all three READMEs with verified count

**Example:**
```csharp
// File: ExcelDataModelTool.cs
[RegularExpression(@"^(list-tables|list-measures|...|update-relationship)$")]
// Count actions in switch statement (line 150-175)
// Result: 14 actions (NOT 15)
```

**Never trust:**
- ❌ Private methods that exist but aren't wired to switch
- ❌ Documentation claims without code verification
- ❌ Old comments or dead code

### Rule 3: No Overclaiming Features

**Accurate Language:**
- ✅ "Most Excel features supported" or "80+ operations"
- ✅ "Currently supports" (implies active development)
- ✅ "Growing feature set" (sets expectations)
- ❌ "All Excel features" (impossible claim)
- ❌ "Complete support" (overpromise)
- ❌ "Full compatibility" (too broad)

**Known Limitations to Acknowledge:**
- DAX calculated columns NOT supported (Excel UI only)
- Subset of Excel capabilities implemented (not all)
- Active development expanding features

### Rule 4: Maintain Appropriate Length

**Length Guidelines:**
- Main README: 250-300 lines (comprehensive reference)
- NuGet README: 80-100 lines (concise gateway)
- VS Code README: 100-120 lines (focused benefits)

**If a README exceeds its target:**
1. Check for duplicate content across sections
2. Look for verbose examples that could be condensed
3. Move detailed content to linked documentation
4. Use collapsible `<details>` sections (main README only)

### Rule 5: Update All Documentation Files Together

**When making changes:**
- Tool count changes → Update all 3 READMEs + server.json description (mention key tools)
- Action count changes → Update all 3 READMEs
- Safety messaging changes → Update all 3 READMEs
- Repository URL changes → Update server.json repository.url field
- Major tool additions → Update server.json description to list new capabilities
- User benefit changes → Update main + VS Code READMEs
- New features → Update main README first, then summarize in others

**⚠️ Version numbers are NEVER manually updated:**
- server.json version field is updated automatically by release workflow
- NuGet package version is updated automatically by release workflow
- See `docs/RELEASE-STRATEGY.md` for the release process

**Sequential Update Process:**
1. Update main README (most comprehensive)
2. Extract brief version for NuGet README
3. Extract user-benefit version for VS Code README
4. Verify consistency across all three

---

## 🔍 Quality Checklist

Before committing README changes, verify:

### Content Accuracy
- [ ] Tool count = 11 in all READMEs
- [ ] Action counts verified against code (not assumptions)
- [ ] No features claimed that aren't implemented
- [ ] Safety messaging consistent (COM API, zero corruption, interactive)
- [ ] "Who Should Use This" sections match (where present)
- [ ] server.json version matches NuGet package version
- [ ] server.json description mentions key tools (Power Query, DAX, VBA, Tables, ranges, connections)

### Appropriate Scope
- [ ] Main README is comprehensive (250-300 lines)
- [ ] NuGet README is concise (80-100 lines, gateway to full docs)
- [ ] VS Code README focuses on benefits (100-120 lines)
- [ ] No major duplication between READMEs

### Markdown Correctness
- [ ] No broken links
- [ ] Code blocks properly closed (triple backticks)
- [ ] HTML tags properly closed (`<details>`, `<summary>`)
- [ ] Headers use proper hierarchy (no skipping levels)
- [ ] Lists consistently formatted (bullets or numbers)

### Discoverability
- [ ] Clear one-line description at top
- [ ] Examples use natural language (actual questions users ask)
- [ ] Requirements clearly stated (Windows, Excel, .NET)
- [ ] Installation steps are accurate and tested

---

## 🚫 Common Mistakes to Avoid

### ❌ Mistake 1: Duplicate Tool Entries
```markdown
❌ WRONG:
3. **excel_table** (22 actions) - Excel Tables
...
9. **excel_table** (22 actions) - Excel Tables management
```

**Why wrong:** Same tool listed twice with different descriptions

**Fix:** List each tool exactly once with consistent description

---

### ❌ Mistake 2: Unverified Action Counts
```markdown
❌ WRONG:
**excel_datamodel** (15 actions) - DAX measures, relationships, calculated columns
```

**Why wrong:**
1. Code has 14 actions (not 15)
2. Calculated columns NOT supported

**Fix:** Verify code, update count, remove unsupported features

---

### ❌ Mistake 3: Verbose NuGet README
```markdown
❌ WRONG: 527-line NuGet README with detailed examples, architecture diagrams, etc.
```

**Why wrong:** NuGet README should be concise gateway, not duplicate of main README

**Fix:** Reduce to 80-100 lines with brief tool list and link to full docs

---

### ❌ Mistake 4: Overclaiming Compatibility
```markdown
❌ WRONG:
- ✅ **Full Excel Feature Access** - All Excel features work perfectly
- ✅ **Complete feature support** - Access to everything
```

**Why wrong:** Subset of features implemented, not all

**Fix:** Use "Most features" or "80+ operations" or "Growing feature set"

---

### ❌ Mistake 5: Missing Safety Positioning
```markdown
❌ WRONG: No mention of COM API safety advantage
```

**Why wrong:** Key differentiator vs third-party .xlsx manipulation libraries

**Fix:** Add "🛡️ 100% Safe - Uses Excel's Native API" callout with 3 benefits

---

### ❌ Mistake 6: Manually Updating Version Numbers
```json
❌ WRONG: 
// Manually editing server.json version field
"version": "1.1.0"
// Or manually editing .csproj version
```

**Why wrong:** Version numbers are automatically managed by release workflow

**Fix:** 
- Never manually update version numbers
- Release workflow updates server.json and .csproj automatically when tags are pushed
- See `docs/RELEASE-STRATEGY.md` for release process

---

## 📚 Related Documentation

- **Main README**: `/README.md` - Comprehensive reference
- **NuGet README**: `/src/ExcelMcp.McpServer/README.md` - Package description
- **VS Code README**: `/vscode-extension/README.md` - Extension description
- **MCP Server Metadata**: `/src/ExcelMcp.McpServer/.mcp/server.json` - Registry metadata
- **Tool Implementation**: `src/ExcelMcp.McpServer/Tools/*.cs` - Source of truth for action counts

---

## 🔄 Maintenance Workflow

**When adding a new tool:**
1. Implement tool in `src/ExcelMcp.McpServer/Tools/`
2. Count actions in switch statement
3. Update main README tool list (add new row)
4. Update main README "11 tools" → "12 tools"
5. Update NuGet README tool list (add new row)
6. Update NuGet README "11 specialized tools" → "12 specialized tools"
7. Update VS Code README tool list (add new row)
8. Update VS Code README "11 specialized tools" → "12 specialized tools"
9. Update server.json description if it's a major tool (add to key tools list)
10. Verify consistency across all four files

**When modifying existing tool:**
1. Update implementation in `src/ExcelMcp.McpServer/Tools/`
2. Count new action count in switch statement
3. Update all three README tool lists with new count
4. Verify descriptions still accurate

**When changing safety messaging:**
1. Update main README safety callout
2. Update NuGet README safety callout (concise version)
3. Update VS Code README safety callout (same as main)
4. Verify all three mention: zero corruption, interactive, growing features

**⚠️ Version Updates:**
- **NEVER** manually update version numbers in server.json or .csproj files
- Version numbers are automatically managed by the release workflow
- See `docs/RELEASE-STRATEGY.md` for release process

---

## 📝 Templates

### Main README Tool List Entry
```markdown
1. **excel_toolname** (X actions) - Feature area: action1, action2, action3, action4
```

### NuGet README Tool List Entry
```markdown
| **excel_toolname** | X actions | Brief purpose (5-8 words) |
```

### VS Code README Tool List Entry
```markdown
| **excel_toolname** | X actions | Brief purpose (5-8 words) |
```

### Safety Callout (Main + VS Code)
```markdown
**🛡️ 100% Safe - Uses Excel's Native API**

Unlike third-party libraries that manipulate `.xlsx` files directly (risking file corruption), ExcelMcp uses **Excel's official COM API**. This ensures:
- ✅ **Zero risk of document corruption** - Excel handles all file operations safely
- ✅ **Interactive development** - See changes in real-time as you work with live Excel files
- ✅ **Growing feature set** - Currently supports 80+ operations across Power Query, Power Pivot, VBA, PivotTables, Tables, and more (active development)
```

### Safety Callout (NuGet - Concise)
```markdown
**🛡️ 100% Safe - Uses Excel's Native COM API**

Unlike third-party libraries that manipulate `.xlsx` files (risking corruption), ExcelMcp uses **Excel's official COM automation API**. This guarantees zero risk of file corruption while you work interactively with live Excel files - see your changes happen in real-time. Currently supports 80+ operations with active development expanding capabilities.
```

---

## 🎓 Key Lessons Learned

1. **Three READMEs, Three Purposes** - Main (comprehensive), NuGet (gateway), VS Code (benefits). Don't duplicate, differentiate.

2. **Always Verify Counts** - Code is source of truth. Don't trust documentation, comments, or assumptions.

3. **Safety is Key Differentiator** - COM API vs .xlsx manipulation is critical positioning for NuGet/Marketplace discoverability.

4. **Active Development = Manage Expectations** - "Growing feature set" and "Currently supports 80+ operations" sets realistic expectations while showing momentum.

5. **Consistency Builds Trust** - When tool counts, action counts, and messaging match across all READMEs, users trust the documentation.

6. **Concise Converts Better** - NuGet README reduced from 527 lines to 89 lines increased clarity and conversion potential.

7. **Dead Code Misleads** - `ListTableColumnsAsync` exists but not in switch statement. Private methods ≠ available features.

8. **VS Code Users = GitHub Users** - Same audience, same benefits. Main README and VS Code README should mirror each other in structure and messaging.
