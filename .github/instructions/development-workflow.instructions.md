---
applyTo: ".github/workflows/**/*.yml,**/*.csproj,global.json"
---

# Development Workflow

> **Required process for all contributions**

## Branch Protection

**⛔ NEVER commit directly to `main`**

Enforced: PR reviews, CI/CD checks, create a branch first, up-to-date branches, no force pushes

**Merge Strategy: Squash Merge**

All PRs are merged using **squash merge** (single commit to `main`). This keeps git history clean and makes it easy to revert changes if needed. When you merge a PR:
- GitHub automatically squashes all commits from your branch into one commit
- Commit message is auto-populated from PR title/description — verify it's accurate before confirming merge
- Your feature branch can be safely deleted after merge

## Development Process

1. **Create feature branch**: `git checkout -b feature/name`
2. **Standards**: Zero warnings, tests pass, docs updated, security rules followed
4. **PR Checklist**: Build passes, tests pass, docs updated, patterns followed, changeset added (see Rule 27)
4. **Check PR review comments**: After creating PR, retrieve automated review feedback and fix all issues
5. **Versions**: Automated via release workflow - don't update manually

## PR Review Comment Workflow

**After creating a PR, ALWAYS check for automated review comments:**

```powershell
# Retrieve inline code review comments using GitHub CLI
# ⚠️ IMPORTANT: for this public repo, gh CLI must be authenticated as a PERSONAL GitHub account.
# Enterprise Managed User (EMU) accounts cannot access public repos via gh CLI.
# Verify with: gh auth status
# If needed, select the personal account's token (Copilot CLI exposes it as an env var):
#   $env:GH_TOKEN = $env:COPILOT_GH_ACCOUNT_github_2E_com_sbroenne   # then verify: gh api user --jq '.login'
# (Admin ops — rulesets, disabling workflows, deleting runs — require this personal token, not the EMU account.)
gh api repos/sbroenne/mcp-server-excel/pulls/PULL_NUMBER/comments --paginate

# Or use the mcp_github tool if available
mcp_github_github_pull_request_read(method="get_review_comments", owner="sbroenne", repo="mcp-server-excel", pullNumber=PULL_NUMBER)
```

**Common automated reviewers:**
- **Copilot** (code quality, performance, style)
- **github-advanced-security** (security scanning, code analysis)

**Common issues to fix:**
- Improper `/// <inheritdoc/>` on constructors/test methods that don't override
- `.AsSpan().ToString()` inefficiency - use `[..n]` range operator instead
- Nullable type access without null checks
- `foreach` → `.Select()` for functional style
- Nested if statements that can be combined
- Generic catch clauses - use specific exceptions or add justification
- Path.Combine security warnings - suppress with justification for test code

**Fix all automated review comments before requesting human review.**

## Test Execution

**See testing-strategy.instructions.md for complete test commands.**

Quick reference:
- Development: `Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust`
- Session/batch changes: run `RunType=OnDemand` in `ExcelMcp.ComInterop.Tests`
- Core OnDemand tests are environment-specific diagnostics and stay outside required CI
- VBA tests: `(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand`

## CI/CD Workflows

**Automated on Pull Requests:**
- `ci.yml` (**CI Gate**) - Release build + Excel-free audit gates + Excel-free CLI/MCP build smoke (always runs on PRs to `main`, so it can be a required check). GitHub-hosted runners have no Excel, so Excel-dependent gates stay local-only in the pre-commit hook.
- `codeql.yml` - Security analysis
- `dependency-review.yml` - Dependency security scanning

**Excel Integration Tests:**
- `integration-tests.yml` starts the cost-optimized Azure runner, runs the real Excel integration suite, and deallocates the VM.
- A same-repository PR runs the full suite once when it is opened as ready or moves from draft to ready. There is no nightly schedule.
- Manual dispatch defaults to one Core feature and also supports individual ComInterop, MCP, CLI, and ComInterop OnDemand session scopes.
- The full merge gate includes the ComInterop OnDemand session tests. It excludes Core OnDemand diagnostics that intentionally require optional Python licensing, IRM-protected files, or manual CPU analysis.
- During development, reproduce one test locally and run only the affected feature remotely. The ruleset requires `Full Excel integration suite` on the latest PR commit; scoped runs use a different check name and cannot satisfy the merge gate.
- The VM has a five-hour auto-shutdown watchdog in case workflow cleanup cannot run.
- See `docs/AZURE_SELFHOSTED_RUNNER_SETUP.md` for provisioning and maintenance.

## Workflow Config Updates

**⚠️ Update ALL workflows when changing:**
- .NET SDK version (`global.json` + all workflows)
- Assembly/package names (`.csproj` + workflow references)
- Runtime requirements (target framework + release notes)
- Project structure (path filters + build commands)

## Quality Enforcement

**Build Settings:** `TreatWarningsAsErrors=true`, analyzers enabled

**Security Rules (Errors):** CA2100 (SQL injection), CA3003 (file path injection), CA3006 (process injection), CA5389 (archive traversal), CA5390 (hardcoded encryption), CA5394 (insecure randomness)

## Release Process (Maintainers)

**Release:** Use `workflow_dispatch` on the release workflow with version bump (major/minor/patch) or custom version. Releases ALL components with same version:
- MCP Server → NuGet + ZIP
- CLI → NuGet + ZIP
- VS Code Extension → Marketplace + VSIX
- MCPB → Claude Desktop bundle

**Before Releasing:**
1. Nothing manual — changesets accumulate in `.changeset/` from individual PRs (see Rule 27); the release workflow compiles them into `CHANGELOG.md` automatically
2. Go to Actions → Release All Components → Run workflow
3. Select version bump type (patch/minor/major) or enter a custom version

Workflow calculates version → builds all components → creates git tag → GitHub release with all artifacts

**Quick release from terminal:**
```powershell
gh workflow run release.yml -f bump=patch   # or minor/major
```

## GitHub Issue Comment Protocol

**ALWAYS verify @mention usernames before posting comments on issues or PRs.**

1. Read the issue/PR to confirm the actual author's GitHub handle
2. Use the correct handle in @mentions — never guess from display names
3. Wrong @mentions are embarrassing and may notify the wrong person
4. When posting on multiple issues in sequence, re-verify each author (they differ!)

## Key Principles

1. Feature branches mandatory
2. Tests required
3. CI/CD must pass
4. Documentation updated
5. Version management automated
6. Security enforced
