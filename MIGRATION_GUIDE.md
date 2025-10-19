# Migration Guide: Moving to mcp-server-excel

This guide outlines the process for migrating the ExcelMcp project to the new repository at `https://github.com/sbroenne/mcp-server-excel.git` without preserving Git history.

## Prerequisites

1. Ensure you have created the new repository at GitHub: `https://github.com/sbroenne/mcp-server-excel.git`
2. Have write access to the new repository
3. Have all your changes committed in the current repository

## Migration Steps

### Option 1: Direct Push (No History)

```powershell
# 1. Navigate to your current repository
cd d:\source\ExcelCLI

# 2. Add the new remote
git remote add new-origin https://github.com/sbroenne/mcp-server-excel.git

# 3. Create a fresh branch without history
git checkout --orphan main-new

# 4. Stage all files
git add -A

# 5. Create initial commit
git commit -m "Initial commit: ExcelMcp - Excel Command Line Interface and MCP Server

Migrated from github.com/sbroenne/ExcelCLI
- Complete CLI tool for Excel automation
- MCP Server for AI-assisted Excel development
- Power Query, VBA, and worksheet operations
- Comprehensive test suite and documentation"

# 6. Push to new repository
git push new-origin main-new:main

# 7. Clean up old remote reference
git remote remove new-origin
```

### Option 2: Pull Request Approach (Recommended)

This approach creates a PR in the new repository for review before merging:

```powershell
# 1. Navigate to your current repository
cd d:\source\ExcelCLI

# 2. Add the new remote
git remote add new-origin https://github.com/sbroenne/mcp-server-excel.git

# 3. Fetch the new repository (if it has any initial content)
git fetch new-origin

# 4. Create a fresh branch without history
git checkout --orphan migration-pr

# 5. Stage all files
git add -A

# 6. Create initial commit
git commit -m "Initial commit: ExcelMcp - Excel Command Line Interface and MCP Server

Migrated from github.com/sbroenne/ExcelCLI

Key Features:
- CLI tool for Excel automation (Power Query, VBA, worksheets)
- MCP Server for AI-assisted Excel development workflows
- Comprehensive test suite with unit, integration, and round-trip tests
- Security-focused with input validation and resource limits
- Well-documented with extensive developer guides

Components:
- ExcelMcp.Core: Shared Excel COM interop operations
- ExcelMcp: Command-line interface executable
- ExcelMcp.McpServer: Model Context Protocol server
- ExcelMcp.Tests: Comprehensive test suite

Documentation:
- README.md: Project overview and quick start
- docs/COMMANDS.md: Complete command reference
- docs/DEVELOPMENT.md: Developer guide for contributors
- docs/COPILOT.md: GitHub Copilot integration patterns
- .github/copilot-instructions.md: AI assistant instructions"

# 7. Push to new repository as a feature branch
git push new-origin migration-pr

# 8. Create Pull Request via GitHub UI
# - Go to https://github.com/sbroenne/mcp-server-excel
# - You'll see a prompt to create PR from migration-pr branch
# - Add detailed description about the migration
# - Review changes and merge when ready

# 9. After PR is merged, update your local repository
git checkout main
git remote set-url origin https://github.com/sbroenne/mcp-server-excel.git
git fetch origin
git reset --hard origin/main

# 10. Clean up migration branch
git branch -D migration-pr
```

## Post-Migration Tasks

### 1. Update Repository References

Update all references to the old repository URL in:

- [ ] `README.md` - Update any links or references
- [ ] `docs/INSTALLATION.md` - Update installation instructions
- [ ] `.github/workflows/*.yml` - Update any workflow references
- [ ] `ExcelMcp.csproj` and other project files - Check for repository URLs
- [ ] `Directory.Build.props` - Update package metadata

### 2. Configure New Repository

In the new GitHub repository, configure:

- [ ] Branch protection rules for `main`
- [ ] Required status checks (CI/CD)
- [ ] Code owners (if applicable)
- [ ] Repository description and topics
- [ ] Enable GitHub Pages (if documentation site exists)
- [ ] Configure secrets for CI/CD

### 3. Update CI/CD Workflows

Verify all GitHub Actions workflows work in the new repository:

- [ ] Build and test workflow
- [ ] Release workflow
- [ ] Any deployment workflows

### 4. Update Documentation

- [ ] Update README badges with new repository path
- [ ] Update contributing guidelines
- [ ] Update issue/PR templates
- [ ] Add migration note to documentation

### 5. Notify Users

If this is a public project:

- [ ] Create an issue in the old repository noting the migration
- [ ] Update old repository README with migration notice
- [ ] Archive the old repository (optional)

## Verification Checklist

After migration, verify:

- [ ] All files are present in new repository
- [ ] `.gitignore` is working correctly
- [ ] CI/CD pipelines execute successfully
- [ ] All documentation links work
- [ ] NuGet package publishing works (if applicable)
- [ ] MCP server configuration is correct
- [ ] Tests pass in new repository

## Rollback Plan

If issues arise, you can:

1. Keep the old repository active until migration is confirmed successful
2. The old repository remains at `https://github.com/sbroenne/ExcelCLI.git`
3. No data is lost in the original location

## Notes

- This migration does NOT preserve Git history (as requested)
- All files are preserved, only commit history is reset
- The old repository can be archived or deleted after successful migration
- Consider adding a "migrated from" note in the new repository's README

## Questions?

Review the documentation:

- GitHub's guide on renaming/moving repositories
- Git documentation on orphan branches
- Model Context Protocol server setup
