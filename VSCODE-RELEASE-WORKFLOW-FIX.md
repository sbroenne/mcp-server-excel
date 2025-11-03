# VS Code Extension Release Workflow Fix

## Problem

The VS Code extension release workflow (`release-vscode-extension.yml`) was failing because it used **bash commands on Windows**.

**Failed workflow**: https://github.com/sbroenne/mcp-server-excel/actions/runs/19029585463

### Root Cause

The workflow runs on `windows-latest` (line 13) but three steps used `shell: bash`:

1. **Update Extension Version** (lines 76-94)
   - Used `sed -i` to update CHANGELOG.md
   - `sed -i` doesn't work properly on Windows Git Bash
   - Used `date +%Y-%m-%d` (bash date command)

2. **Package Extension** (lines 111-127)
   - Used `mv` to rename files
   - Used `ls -lh` for file listing
   - Used bash variable syntax `${VERSION}`

3. **Create GitHub Release** (lines 138-244)
   - Used `cat > file << EOF` heredoc syntax
   - Used bash variable syntax throughout

### Why It Failed

On Windows, Git Bash provides minimal POSIX compatibility but:
- `sed -i` behavior differs from Linux (file editing issues)
- Path handling is inconsistent (backslashes vs forward slashes)
- Heredoc syntax can fail with complex multiline content
- Environment variable setting (`echo "VAR=value" >> $GITHUB_ENV`) may have encoding issues

## Solution

**Converted all bash steps to PowerShell (`pwsh`)**

PowerShell is the native Windows shell and works consistently on `windows-latest` runners.

### Changes Made

#### 1. Update Extension Version (lines 76-93)

**Before** (bash):
```yaml
- name: Update Extension Version
  run: |
    TAG_NAME="${{ github.ref_name }}"
    VERSION="${TAG_NAME#vscode-v}"
    
    echo "Updating VS Code extension to version $VERSION"
    
    cd vscode-extension
    npm version "$VERSION" --no-git-tag-version
    
    DATE=$(date +%Y-%m-%d)
    sed -i "0,/## \[[0-9.]*\] - [0-9-]*/s//## [$VERSION] - $DATE/" CHANGELOG.md
    
    echo "Updated extension version to $VERSION"
    echo "PACKAGE_VERSION=$VERSION" >> $GITHUB_ENV
  shell: bash
```

**After** (PowerShell):
```yaml
- name: Update Extension Version
  run: |
    $tagName = "${{ github.ref_name }}"
    $version = $tagName -replace '^vscode-v', ''
    
    Write-Output "Updating VS Code extension to version $version"
    
    cd vscode-extension
    npm version "$version" --no-git-tag-version
    
    $date = Get-Date -Format "yyyy-MM-dd"
    $changelogPath = "CHANGELOG.md"
    $changelogContent = Get-Content $changelogPath -Raw
    $changelogContent = $changelogContent -replace '(?m)^## \[\d+\.\d+\.\d+\] - \d{4}-\d{2}-\d{2}', "## [$version] - $date"
    Set-Content $changelogPath $changelogContent
    
    Write-Output "Updated extension version to $version"
    "PACKAGE_VERSION=$version" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
  shell: pwsh
```

**Key improvements**:
- `-replace` operator instead of `sed`
- `Get-Date -Format` instead of `date +%`
- `Get-Content -Raw` + `Set-Content` instead of `sed -i`
- `Out-File -Encoding utf8` instead of `>>`

#### 2. Package Extension (lines 111-125)

**Before** (bash):
```yaml
- name: Package Extension
  run: |
    cd vscode-extension
    npx @vscode/vsce package --no-dependencies --allow-missing-repository
    
    VERSION="${{ env.PACKAGE_VERSION }}"
    VSIX_FILE="excelmcp-${VERSION}.vsix"
    
    if [ -f "excelmcp-*.vsix" ]; then
      mv excelmcp-*.vsix "$VSIX_FILE"
    fi
    
    echo "Created $VSIX_FILE"
    ls -lh "$VSIX_FILE"
    echo "VSIX_PATH=vscode-extension/$VSIX_FILE" >> $GITHUB_ENV
  shell: bash
```

**After** (PowerShell):
```yaml
- name: Package Extension
  run: |
    cd vscode-extension
    npx @vscode/vsce package --no-dependencies --allow-missing-repository
    
    $version = "${{ env.PACKAGE_VERSION }}"
    $vsixFile = "excelmcp-$version.vsix"
    
    $existingVsix = Get-ChildItem -Path . -Filter "excelmcp-*.vsix" -File | Select-Object -First 1
    if ($existingVsix -and $existingVsix.Name -ne $vsixFile) {
      Rename-Item $existingVsix.FullName -NewName $vsixFile
    }
    
    Write-Output "Created $vsixFile"
    Get-Item $vsixFile | Format-Table Name, Length
    "VSIX_PATH=vscode-extension/$vsixFile" | Out-File -FilePath $env:GITHUB_ENV -Encoding utf8 -Append
  shell: pwsh
```

**Key improvements**:
- `Get-ChildItem` instead of wildcard file matching
- `Rename-Item` instead of `mv`
- `Format-Table` instead of `ls -lh`

#### 3. Create GitHub Release (lines 138-250)

**Before** (bash heredoc):
```bash
cat > release_notes.md << EOF
## ExcelMcp VS Code Extension $TAG_NAME
...
EOF
```

**After** (PowerShell here-string):
```powershell
$releaseNotes = @"
## ExcelMcp VS Code Extension $tagName
...
"@

Set-Content -Path "release_notes.md" -Value $releaseNotes
```

**Key improvements**:
- `@"..."@` here-string instead of `<< EOF`
- `Set-Content` instead of `cat >`
- PowerShell variables (`$tagName`) instead of bash (`$TAG_NAME`)
- Consistent `pwsh` shell

## Testing

The fix has been committed and pushed to branch `fix/code-scanning-issues`.

**Commit**: `6af4e47` - "Fix VS Code extension release workflow - convert bash to PowerShell"

### Next Steps

1. **Merge to main** via PR
2. **Test the workflow** by pushing a new tag:
   ```bash
   git tag vscode-v1.2.1
   git push origin vscode-v1.2.1
   ```
3. **Verify workflow** completes successfully at:
   https://github.com/sbroenne/mcp-server-excel/actions

### Expected Behavior After Fix

✅ Update Extension Version step completes  
✅ CHANGELOG.md updated with correct version and date  
✅ Package Extension step creates VSIX file  
✅ VSIX file properly renamed to `excelmcp-{version}.vsix`  
✅ GitHub Release created with release notes  
✅ VSIX file attached to release  

## Prevention

**Updated workflow config rules**:
- When workflow runs on `windows-latest`, use `shell: pwsh`
- When workflow runs on `ubuntu-latest`, use `shell: bash`
- Never mix shells within same job
- Test workflows locally with `act` before pushing tags

**See**: `.github/instructions/development-workflow.instructions.md`

## Summary

| Aspect | Before | After |
|--------|--------|-------|
| Shell | bash | pwsh |
| File editing | sed -i | Get-Content + regex + Set-Content |
| Date formatting | date +%Y-%m-%d | Get-Date -Format "yyyy-MM-dd" |
| File operations | mv, ls | Rename-Item, Format-Table |
| Heredocs | cat << EOF | @"..."@ here-string |
| Variables | ${VAR} | $var |
| Env vars | echo >> $GITHUB_ENV | Out-File -Encoding utf8 |

**Result**: Workflow now uses native Windows commands that work reliably on `windows-latest` runners.
