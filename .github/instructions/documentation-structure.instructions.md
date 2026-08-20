---
applyTo: "README.md,FEATURES.md,CHANGELOG.md,SECURITY.md,PRIVACY.md,docs/**/*.md,specs/**/*.md,src/**/README.md,skills/**/*.md,gh-pages/docs/**/*.md"
excludeAgent: "code-review"
---

# Documentation Structure & Standards

> **Clear hierarchy prevents temporary doc accumulation**

## 📁 Documentation Hierarchy

### Root Level - Essential User-Facing Only
- ✅ `README.md` - GitHub acquisition page and quick start
- ✅ `FEATURES.md` - Feature navigation hub and tool-selection guide
- ✅ `CHANGELOG.md` - Generated release history
- ✅ `SECURITY.md` - Security policy (GitHub standard)
- ✅ `PRIVACY.md` - Privacy policy
- ✅ `LICENSE` - License file
- ❌ **NO** temporary files (SUMMARY, FIX, BUG, TESTS, DOCS, etc.)

### `docs/` - Canonical Documentation
**Purpose:** Feature references, user guides, architecture, and development processes

**Categories:**
- **Feature References:** `features/*.md`
- **User Guides:** `INSTALLATION*.md`, `USE-CASES.md`, `CONTRIBUTING.md`
- **Architecture:** `ARCHITECTURE.md`, `ADR-*.md`
- **Developer Guides:** `DEVELOPMENT.md`, `PRE-COMMIT-SETUP.md`
- **Process Docs:** `RELEASE-STRATEGY.md`, `MCP_REGISTRY_PUBLISHING.md`, `NUGET-GUIDE.md`
- **Infrastructure:** `infrastructure/azure/README.md`
- **Standards:** `TEST-NAMING-STANDARD.md`

**Naming Convention:**
- ✅ `TOPIC-NAME.md` (ALL CAPS for discoverability)
- ✅ `ADR-NNN-DECISION-NAME.md` (Architecture Decision Records)
- ❌ NO `SUMMARY.md`, `FIX.md`, `TESTS.md` (temporary naming)

### `specs/` - Feature Specifications
**Purpose:** What should be built (requirements, design, before implementation)

**Naming Convention:**
- ✅ `FEATURE-NAME-SPEC.md`
- ✅ `COMPONENT-API-SPECIFICATION.md`

### `src/[Component]/` - Component Documentation
**Purpose:** Component-specific overview and usage

**Files:**
- ✅ `README.md` - Component overview, quick start

---

## Temporary Documentation Rules

### ❌ Forbidden at Root Level
- `FIX-*.md` - Document fixes in PRs, not files
- `BUG-*.md` - Track bugs in GitHub Issues
- `TESTS-*.md` - Test info belongs in test files
- `DOCS-*.md` - Update actual docs, don't create meta-docs
- `SUMMARY-*.md` - Summarize in PR descriptions

### ✅ Where to Put Content Instead
- Bug investigation → GitHub Issue comments
- Fix summary → PR description
- Architecture decisions → `docs/ADR-NNN-DECISION-NAME.md`
- Temporary notes → Branch commit messages (deleted after merge)

---

## Document Lifecycle

### Before Creating a Doc
1. **Canonical source first:** Update canonical repository docs before Pages wrappers.
2. **Preserve information:** Map substantive content before shortening or relocating a document.
3. **Wrappers stay thin:** `gh-pages/docs/` adds presentation and SEO, not a second source of truth.
4. **Is this permanent?** → YES: Use proper location above.
5. **Is this temporary?** → Put in PR/Issue/commit message instead.
6. **Does equivalent doc exist?** → Update existing, don't duplicate.

### During PR Review
- ❌ Root-level temporary docs → Move to proper location or delete
- ✅ Permanent docs → Must follow naming conventions above

### After PR Merge
- Delete temporary docs if any slipped through
- Verify permanent docs in correct location

---

## Canonical sources published to the website

`gh-pages/hooks.py` transforms canonical repo docs into `gh-pages/docs/_generated/`
(git-ignored) and thin wrappers include them with `--8<--`. Anything listed below
is authored **once**, in the repo, and published automatically:

| Canonical source | Site route |
|---|---|
| `README.md`, `FEATURES.md`, `CHANGELOG.md`, `SECURITY.md`, `PRIVACY.md` | home, `/features/`, `/changelog/`, `/security/`, `/privacy/` |
| `docs/features/*.md` | `/features/<slug>/` |
| `docs/guides/*.md` | `/guides/<slug>/` |
| `docs/INSTALLATION*.md`, `docs/ARCHITECTURE.md`, `docs/USE-CASES.md`, `docs/CONTRIBUTING.md` | matching routes |
| `src/ExcelMcp.McpServer/README.md`, `src/ExcelMcp.CLI/README.md`, `skills/README.md` | `/mcp-server/`, `/cli/`, `/skills/` |
| `skills/shared/*.md` | `/reference/<slug>/` |

When adding a new canonical doc that should appear on the site:

1. Write it under `docs/` (or `skills/shared/` for agent reference material).
2. Add it to the matching source map in `gh-pages/hooks.py` **and** to
   `SITE_PAGE_MAP` so cross-references resolve on-site instead of linking to GitHub.
3. Add a thin wrapper under `gh-pages/docs/` owning only `title`, `description`,
   `keywords`, the H1, and the `--8<--` include.
4. Add a nav entry in `gh-pages/mkdocs.yml`.
5. Build with `python -m mkdocs build --strict` from `gh-pages/`.

The build also derives `/llms.txt`, `/llms-full.txt`, `/tools.json`, and a
Markdown mirror for every page from the same content — no separate maintenance.

---

## Naming Standards

### ✅ Good Names (Discoverable)
- `DEVELOPMENT.md`
- `ADR-001-NO-UNIT-TESTS.md`
- `RANGE-API-SPECIFICATION.md`
- `PRE-COMMIT-SETUP.md`

### ❌ Bad Names (Temporary/Vague)
- `notes.md`
- `temp.md`
- `SUMMARY.md`
- `FIX-123.md`
- `NEW-FEATURE.md`

---

## File Organization Rules

1. **Root level** = Permanent, essential, user-facing only
2. **docs/** = Permanent implementation/process documentation
3. **specs/** = Permanent feature specifications
4. **src/Component/** = Permanent component-specific docs
5. **Nowhere** = Temporary documentation (use PRs/Issues instead)

---

## Quick Checklist

Before committing a `.md` file, verify:
- [ ] File is permanent (not temporary investigation/fix notes)
- [ ] File location matches hierarchy above
- [ ] File name follows naming conventions (ALL CAPS or kebab-case)
- [ ] No duplicate documentation exists
- [ ] Content is complete (not placeholder)
- [ ] Removed content has a verified canonical destination
- [ ] GitHub Pages wrappers contain no duplicated operational reference
