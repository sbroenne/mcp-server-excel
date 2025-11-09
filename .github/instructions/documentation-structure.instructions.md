---
applyTo: "**/*.md,docs/**,specs/**"
---

# Documentation Structure & Standards

> **Prevent temporary documentation accumulation with clear hierarchy and naming conventions**

## 📁 Documentation Hierarchy

### Root Level - Essential User-Facing Only
- ✅ `README.md` - Main project overview
- ✅ `SECURITY.md` - Security policy (GitHub standard)
- ✅ `LICENSE` - License file
- ❌ **NO** temporary files (SUMMARY, FIX, BUG, TESTS, DOCS, etc.)

### `docs/` - Implementation Documentation
**Purpose:** How things work, how to use them, architectural decisions

**Categories:**
- **User Guides:** `CONTRIBUTING.md`
- **Developer Guides:** `DEVELOPMENT.md`, `PRE-COMMIT-SETUP.md`
- **Process Docs:** `RELEASE-STRATEGY.md`, `MCP_REGISTRY_PUBLISHING.md`, `NUGET-GUIDE.md`
- **Architecture:** `ADR-*.md` (Architecture Decision Records)
- **Infrastructure:** `AZURE_SELFHOSTED_RUNNER_SETUP.md`
- **Standards:** `TEST-NAMING-STANDARD.md`

**Naming Convention:**
- ✅ `TOPIC-NAME.md` (ALL CAPS for discoverability)
- ✅ `ADR-NNN-DECISION-NAME.md` (Architecture Decision Records)
- ❌ NO `SUMMARY.md`, `FIX.md`, `TESTS.md` (temporary naming)

### `specs/` - Feature Specifications
**Purpose:** What should be built (requirements, design, before implementation)

**Categories:**
- Feature specifications (before implementation)
- API design documents
- Technical requirements

**Naming Convention:**
- ✅ `FEATURE-NAME-SPEC.md`
- ✅ `COMPONENT-API-SPECIFICATION.md`
- Examples: `RANGE-API-SPECIFICATION.md`, `TABLE-API-SPECIFICATION.md`

### `src/[Component]/` - Component Documentation
**Purpose:** Component-specific overview and usage

**Files:**
- ✅ `README.md` - Component overview, quick start
- ❌ NO implementation details (those go in `docs/`)

**Current Components:**
- `src/ExcelMcp.CLI/README.md`
- `src/ExcelMcp.Core/README.md`
- `src/ExcelMcp.ComInterop/README.md`
- `src/ExcelMcp.McpServer/README.md`

### `src/ExcelMcp.McpServer/Prompts/Content/` - LLM Guidance
**Purpose:** Teach LLMs how to use the MCP server

**Categories:**
- **Tool Prompts:** `excel_[tool].md` (one per tool)
- **Completions:** `Completions/[parameter].md` (autocomplete values)
- **Elicitations:** `Elicitations/[workflow].md` (pre-flight checklists)
- **Guidance:** `tool_selection_guide.md`, `server_quirks.md`, `user_request_patterns.md`

**Rules:**
- ✅ ONE file per tool (all actions in same file)
- ✅ SHORT (50-150 lines max)
- ❌ NO Excel tutorials (LLMs already know Excel)

### `tests/` - Test Documentation
**Purpose:** Testing strategy, test data setup, test-specific guides

**Files:**
- ✅ `tests/README.md` - Main test documentation
- ✅ `tests/[Component].Tests/docs/*.md` - Component-specific test docs

**Current:**
- `tests/ExcelMcp.Core.Tests/docs/DATA-MODEL-SETUP.md` (reference docs)

### Subdirectories - Scoped Documentation
**Infrastructure:**
- `infrastructure/azure/README.md` - Azure infrastructure overview
- `infrastructure/azure/GITHUB_ACTIONS_DEPLOYMENT.md`

**VS Code Extension:**
- `vscode-extension/README.md` - Extension overview
- `vscode-extension/DEVELOPMENT.md` - Development guide
- `vscode-extension/CHANGELOG.md` - Version history
- `vscode-extension/INSTALL.md`, `MARKETPLACE-PUBLISHING.md`, etc.

---

## 🚫 Anti-Patterns: What NOT to Create

### ❌ Temporary Fix/Summary Files
**NEVER create these:**
- `BUG-FIX-*.md` in root or docs/
- `TESTS-*.md` in root or docs/
- `DOCS-*.md` in root or docs/
- `SUMMARY.md` anywhere
- `*-FIXES-SUMMARY.md` in root
- `*-MIGRATION-PLAN.md` in root (goes in docs/ or specs/)

**Why:** These accumulate and create clutter. Information belongs in permanent docs or git history.

**Instead:**
1. **Bug fixes:** Document in GitHub Issues + PR description + update relevant permanent docs
2. **Test coverage:** Update `tests/README.md` with new test categories
3. **Documentation changes:** Update the actual docs being changed
4. **Summaries:** Use PR description, not separate file

### ❌ Duplicate Information
**NEVER:**
- Copy same information to multiple files
- Create "summary" of existing documentation
- Keep outdated versions of docs

**Instead:**
- ONE source of truth per topic
- Link to canonical documentation
- Delete old versions (git preserves history)

### ❌ Wrong Location
**NEVER:**
- Put feature specs in `docs/` (they go in `specs/`)
- Put implementation guides in `specs/` (they go in `docs/`)
- Put component docs in root (they go in `src/[Component]/`)
- Put temporary files anywhere (delete after PR merge)

---

## ✅ Correct Documentation Workflow

### Adding a New Feature

**Phase 1: Specification (Before Implementation)**
1. Create `specs/FEATURE-NAME-SPEC.md`
2. Document requirements, design, API surface
3. Get feedback via PR review

**Phase 2: Implementation**
1. Implement feature
2. Add/update component README if needed
3. Update MCP prompts in `src/ExcelMcp.McpServer/Prompts/Content/`

**Phase 3: Post-Merge**
1. Optionally create `docs/FEATURE-NAME-GUIDE.md` if complex
2. Delete any temporary files created during development
3. Keep spec in `specs/` for future reference

### Fixing a Bug

**During Development:**
1. Reproduce bug
2. Create GitHub Issue describing bug
3. Create feature branch

**During PR:**
1. PR description documents:
   - What was broken
   - Root cause
   - How it's fixed
   - Test coverage added
2. Update relevant permanent docs:
   - MCP prompts if tool behavior changed
   - Component README if API changed

**After Merge:**
1. Close GitHub Issue
2. Delete feature branch
3. **NO** separate BUG-FIX-*.md file needed

**Why:** Git history + GitHub Issues + PR descriptions = complete record

### Updating Documentation

**When updating existing docs:**
1. Edit the canonical file directly
2. Don't create `*-UPDATES.md` or `*-CHANGES.md`
3. Use git commit message to describe changes

**When restructuring docs:**
1. Create proposal in PR description
2. Make changes
3. Delete old files (git preserves history)

---

## 📚 Documentation Standards

### File Naming
- ✅ `TOPIC-NAME.md` (ALL CAPS) for docs/
- ✅ `feature-name-spec.md` for specs/ (lowercase with hyphens)
- ✅ `README.md` for component overviews
- ✅ `ADR-NNN-DECISION.md` for architecture decisions

### File Length
- **Guides:** 200-500 lines (split if longer)
- **Prompts:** 50-150 lines max
- **READMEs:** 100-300 lines
- **Specs:** 300-1000 lines (can be longer)

### Writing Style
- **User guides:** Imperative ("Run this command")
- **Developer guides:** Explanatory ("This pattern works because...")
- **Specs:** Declarative ("The system shall...")
- **Prompts:** Action-oriented bullet points

### Markdown Standards
- Use `#` for titles, `##` for sections, `###` for subsections
- Code blocks with language tags: ```bash, ```csharp, ```json
- Tables for comparisons
- Emoji for visual structure (✅ ❌ 📁 🚀 ⚠️)
- Links to related docs

---

## 🎯 GitHub Copilot Guidelines

**When asked to create documentation:**

### 1. Check Existing Documentation First
```bash
# Search for related docs
git grep -i "topic name" -- "*.md"

# List docs in category
ls docs/
ls specs/
```

### 2. Determine Correct Location
- **Specification?** → `specs/FEATURE-NAME-SPEC.md`
- **User guide?** → `docs/TOPIC-NAME.md`
- **Developer guide?** → `docs/DEVELOPMENT.md` or component README
- **Test documentation?** → `tests/README.md`
- **LLM guidance?** → `src/ExcelMcp.McpServer/Prompts/Content/`

### 3. Update Existing Docs When Possible
- ✅ Add section to existing file
- ✅ Update relevant paragraphs
- ❌ Don't create new file if topic already covered

### 4. Never Create Temporary Files
- ❌ NO `SUMMARY.md`, `FIX.md`, `TESTS.md`, `DOCS.md`
- ❌ NO files with dates in names (`2025-01-TOPIC.md`)
- ✅ Use PR descriptions for temporary summaries

### 5. Clean Up After Yourself
- Delete temporary files before PR submission
- Merge temporary notes into permanent docs
- Remove outdated documentation

---

## 🔍 Verification Commands

**Check for temporary files:**
```powershell
# Find potential temporary files
Get-ChildItem -Recurse -Filter "*.md" | 
    Where-Object { $_.Name -match "(SUMMARY|FIX|TESTS|DOCS|MIGRATION|AUDIT|PLAN)" } |
    Select-Object FullName
```

**List documentation by category:**
```powershell
# Root level docs (should be minimal)
Get-ChildItem -Filter "*.md" -File

# Implementation docs
Get-ChildItem docs/*.md

# Specifications
Get-ChildItem specs/*.md

# Test docs
Get-ChildItem tests/**/*.md -Recurse
```

---

## 📖 Examples

### ✅ CORRECT: Adding Range Feature Documentation

**Spec (before implementation):**
```
specs/RANGE-API-SPECIFICATION.md
```

**Implementation:**
```
# Update existing docs
src/ExcelMcp.McpServer/Prompts/Content/excel_range.md (add LLM guidance)
```

**No temporary files created.**

### ❌ WRONG: Creating Temporary Summary

**Don't do this:**
```
docs/RANGE-IMPLEMENTATION-SUMMARY.md   # ❌ Temporary
docs/RANGE-TESTS-ADDED.md               # ❌ Temporary
docs/RANGE-DOCS-UPDATED.md              # ❌ Temporary
```

**Do this instead:**
```
# PR Description includes:
- Feature overview
- Tests added
- Documentation updated
```

### ✅ CORRECT: Bug Fix Documentation

**GitHub Issue:**
```
Title: Bug: refresh parameter ignored
Description: [problem description]
```

**PR Description:**
```
## Bug Fix: Refresh + LoadDestination

**Problem:** Parameter was ignored
**Root Cause:** Missing parameter in method signature
**Fix:** Added parameter, wired through all layers
**Tests:** 13 new tests added
**Docs Updated:** excel_powerquery.md
```

**No separate BUG-FIX-*.md file needed.**

---

## 🎓 Key Takeaways

1. **One source of truth** - Don't duplicate information
2. **Permanent over temporary** - Update existing docs, don't create summaries
3. **Clear hierarchy** - docs/ vs specs/ vs component READMEs
4. **No clutter** - Delete temporary files after PR merge
5. **Git preserves history** - Don't keep old versions "just in case"

---

**Last Updated:** 2025-01-28  
**Related:** `.github/instructions/bug-fixing-checklist.instructions.md`, `.github/instructions/readme-management.instructions.md`
