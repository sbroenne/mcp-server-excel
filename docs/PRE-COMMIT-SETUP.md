# Pre-Commit Hook Setup

This repository includes automated pre-commit checks to prevent code quality issues and coverage regression.

## What Gets Checked

1. **Branch Protection** - Blocks direct commits to `main` branch (Rule 6)
2. **COM Object Leaks** - Ensures all dynamic COM objects are properly released
3. **Core Commands Coverage** - Verifies 100% of Core methods are exposed via MCP Server
4. **Naming Consistency** - Ensures enum action names match Core method names exactly
5. **Success Flag Violations** - Ensures Success=true never paired with ErrorMessage (Rule 1)
6. **CLI Actions Audit** - Verifies CLI action catalog matches Core action enums
7. **CLI Workflow Smoke Test** - Validates the end-to-end CLI workflow
8. **MCP Server Smoke Test** - Validates all 11 MCP tools work correctly

## Setup Instructions

### Option 1: Git Bash (Recommended for cross-platform)

The bash hook at `.git/hooks/pre-commit` works automatically if you have Git Bash installed (default with Git for Windows).

**Test it:**
```bash
bash .git/hooks/pre-commit
```

### Option 2: PowerShell (Windows-specific, more reliable output)

Use the PowerShell script for better formatting and error messages on Windows:

**Manual execution:**
```powershell
.\scripts\pre-commit.ps1
```

**Configure Git to use PowerShell hook:**
```powershell
# Create a wrapper in .git/hooks/pre-commit
@"
#!/bin/sh
pwsh -ExecutionPolicy Bypass -File "scripts/pre-commit.ps1"
"@ | Out-File -FilePath .git/hooks/pre-commit -Encoding ASCII
```

## What Happens on Failure

### Branch Protection Violation
```
❌ BLOCKED: Cannot commit directly to 'main' branch!

   Rule 6: All Changes Via Pull Requests
   'Never commit to main. Create feature branch → PR → CI/CD + review → merge.'

   To fix:
   1. git stash                                    # Save your changes
   2. git checkout -b feature/your-feature-name    # Create feature branch
   3. git stash pop                                # Restore changes
   4. git add <files>                              # Stage changes
   5. git commit -m 'your message'                 # Commit to feature branch
```

**Fix:** Follow the 5 steps above to move your work to a feature branch.

### COM Leak Detected
```
❌ COM object leaks detected! Fix them before committing.
```

**Fix:** Run `.\scripts\check-com-leaks.ps1` to see which files have leaks, then add proper `finally` blocks with `ComUtilities.Release(ref obj!)` calls.

### Coverage Gap Detected
```
❌ Coverage gaps detected! All Core methods must be exposed via MCP Server.
   Fix the gaps before committing (add enum values and mappings).
```

**Fix:** Follow the 5-step process:
1. Add enum values to `ToolActions.cs`
2. Add `ToActionString` mappings to `ActionExtensions.cs`
3. Add switch cases to appropriate MCP Tool
4. Implement MCP methods
5. Build and verify

See `.github/instructions/coverage-prevention-strategy.instructions.md` for details.

## Bypass Pre-Commit Checks (Emergency Only)

If you absolutely must commit without passing the checks (NOT recommended):

```bash
git commit --no-verify -m "Emergency commit message"
```

**⚠️ Warning:** This should only be used in emergencies. Coverage gaps and COM leaks must be fixed before merging to main.

## Testing the Hook

Run manually before committing:

```powershell
# PowerShell
.\scripts\pre-commit.ps1

# Git Bash
bash .git/hooks/pre-commit
```

## Troubleshooting

### PowerShell not found
Install PowerShell 7+ from https://github.com/PowerShell/PowerShell/releases

### Scripts disabled on Windows
Run once as Administrator:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

### Hook not running automatically
Verify the file is executable:
```bash
chmod +x .git/hooks/pre-commit
```

## Continuous Integration

These same checks run in CI/CD pipelines:
- `check-com-leaks.ps1` runs in build workflow
- `audit-core-coverage.ps1` runs **after every build** in:
  - `build-mcp-server.yml` - MCP Server builds
  - `build-cli.yml` - CLI builds  
  - `integration-tests.yml` - Integration test runs

**Pipeline enforcement ensures:**
- Pre-commit hook provides **instant local feedback**
- CI/CD provides **safety net** if hook bypassed with `--no-verify`
- **Double protection** against coverage regression

The pre-commit hook gives you **instant feedback** before pushing to remote.
