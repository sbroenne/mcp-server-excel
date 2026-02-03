---
applyTo: "**/*.md,docs/**,specs/**"
---

# Documentation Structure & Standards

> **Clear hierarchy prevents temporary doc accumulation**

## Documentation Hierarchy

### Root Level - Essential User-Facing Only
- `README.md` - Main project overview
- `SECURITY.md` - Security policy (GitHub standard)
- `LICENSE` - License file
- **NO** temporary files (SUMMARY, FIX, BUG, TESTS, DOCS, etc.)

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
- `TOPIC-NAME.md` (ALL CAPS for discoverability)
- `ADR-NNN-DECISION-NAME.md` (Architecture Decision Records)
- NO `SUMMARY.md`, `FIX.md`, `TESTS.md` (temporary naming)

### `specs/` - Feature Specifications
**Purpose:** What should be built (requirements, design, before implementation)

**Naming Convention:**
- `FEATURE-NAME-SPEC.md`
- `COMPONENT-API-SPECIFICATION.md`

### `src/[Component]/` - Component Documentation
**Purpose:** Component-specific overview and usage

**Files:**
- `README.md` - Component overview, quick start

---

## Temporary Documentation Rules

### Forbidden at Root Level
- `FIX-*.md` - Document fixes in PRs, not files
- `BUG-*.md` - Track bugs in GitHub Issues
- `TESTS-*.md` - Test info belongs in test files
- `DOCS-*.md` - Update actual docs, don't create meta-docs
- `SUMMARY-*.md` - Summarize in PR descriptions

### Where to Put Content Instead
- Bug investigation → GitHub Issue comments
- Fix summary → PR description
- Architecture decisions → `docs/ADR-NNN-DECISION-NAME.md`
- Temporary notes → Branch commit messages (deleted after merge)

---

## Document Lifecycle

### Before Creating a Doc
1. **Is this permanent?** → YES: Use proper location above
2. **Is this temporary?** → Put in PR/Issue/commit message instead
3. **Does equivalent doc exist?** → Update existing, don't duplicate

### During PR Review
- Root-level temporary docs → Move to proper location or delete
- Permanent docs → Must follow naming conventions above

### After PR Merge
- Delete temporary docs if any slipped through
- Verify permanent docs in correct location

---

## Naming Standards

### Good Names (Discoverable)
- `DEVELOPMENT.md`
- `ADR-001-NO-UNIT-TESTS.md`
- `RANGE-API-SPECIFICATION.md`
- `PRE-COMMIT-SETUP.md`

### Bad Names (Temporary/Vague)
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
