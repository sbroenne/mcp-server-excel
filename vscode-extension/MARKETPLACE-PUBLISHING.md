# VS Code Marketplace Publishing Setup

This document explains how to set up automated publishing to the VS Code Marketplace.

## Required GitHub Secret

The release workflow requires the following secret to be configured in your GitHub repository:

### VSCE_TOKEN (VS Code Marketplace)

**Purpose:** Allows automated publishing to the Visual Studio Code Marketplace

**How to create:**

1. **Create a Microsoft Account** (if you don't have one)
   - Go to https://login.live.com/

2. **Create an Azure DevOps organization**
   - Go to https://dev.azure.com/
   - Sign in with your Microsoft account
   - Create a new organization (if needed)

3. **Create a Personal Access Token (PAT)**
   - In Azure DevOps, go to User Settings (top right) → Personal Access Tokens
   - Click "New Token"
   - Name: `VS Code Marketplace Publishing`
   - Organization: Select your organization
   - Expiration: Custom defined (e.g., 1 year)
   - Scopes: Select "Custom defined" → Check "Marketplace (Manage)"
   - Click "Create"
   - **Copy the token** (you won't see it again!)

4. **Create a publisher account** (if you don't have one)
   - Go to https://marketplace.visualstudio.com/manage
   - Click "Create publisher"
   - Publisher ID: Should match `package.json` publisher field (e.g., `sbroenne`)
   - Display name, description, etc.

5. **Add to GitHub Secrets**
   - Go to your GitHub repo → Settings → Secrets and variables → Actions
   - Click "New repository secret"
   - Name: `VSCE_TOKEN`
   - Value: Paste your PAT from step 3
   - Click "Add secret"

## Workflow Behavior

**Note:** The VS Code extension is now released as part of the unified release workflow (`.github/workflows/release.yml`).

When you run the release workflow (via `workflow_dispatch`):

1. **Calculates version** from latest git tag (or custom version input)
2. **Updates `package.json`** version for VS Code extension
3. **Compiles changesets** into `CHANGELOG.md` and release notes
4. **Builds the extension** from source
5. **Packages and verifies Windows x64 and ARM64 VSIX files**
6. **Publishes both platform packages to VS Code Marketplace**
7. **Creates GitHub Release** with all components (MCP Server, CLI, VS Code, MCPB)

### Publishing failures

Building and publishing use the same `@vscode/vsce` version from the extension's
lockfile. The Marketplace job checks out the exact release commit, installs its
locked tools, and uploads the verified VSIX files without rebuilding them.
Publishing runs on Windows because the extension declares `os: ["win32"]`;
installing its locked tools on Linux fails with npm `EBADPLATFORM`.
Each upload uses `--skip-duplicate`, so retrying a partially completed job can
publish the missing platform without failing on the already published one.

The Marketplace job reports publication failures. A missing/expired token or
Marketplace outage can leave one or both platform packages unpublished.
Inspect both publish steps after every release and use the manual fallback
below only with authorization.

### Repair an existing release

The unified release always calls `.github/workflows/publish-vscode.yml` after
creating the GitHub release. The same workflow supports an authorized repair:

```powershell
gh workflow run publish-vscode.yml --ref main -f release_tag=v2.1.2
```

Repair verifies that the tag belongs to `main`, uses publishing tools from that
exact release, and downloads both existing VSIX assets. It does not rebuild the
extension, create another release, or change tags. Both uploads retain
`--skip-duplicate`. A failed Marketplace upload stays a visible workflow failure.

## Checking Packages Without Publishing

Do not dispatch a release to test packaging. From the repository root, after a
successful Release solution build, use:

```powershell
.\scripts\Build-ReleasePackages.ps1 -Components Extension
```

This runs the compile, metadata, lint, type, and Vitest checks and inspects both
VSIX files without publishing. Install the matching package in an isolated
VS Code profile to check activation and real Excel behavior. Release dispatch
and publication require separate authorization.

## Troubleshooting

### "Failed to publish to VS Code Marketplace"

- **Check PAT permissions**: Ensure your Azure DevOps PAT has "Marketplace (Manage)" scope
- **Check PAT expiration**: Tokens expire - you may need to regenerate
- **Check publisher ownership**: Ensure your Azure DevOps account owns the publisher
- **Check package.json**: Publisher field must match your marketplace publisher ID

### "Workflow runs but marketplace shows old version"

- Marketplace updates can take 5-15 minutes to appear
- Clear browser cache or use incognito mode
- Check marketplace directly: https://marketplace.visualstudio.com/items?itemName=PUBLISHER.EXTENSION

## Manual Publishing (Fallback)

If automated publishing fails, publish the **verified, release-stamped VSIX
files** from that release. Do not publish directly from the unprepared source
folder or rebuild an existing release with different inputs. After authorization,
use the locked tool from the extension folder (replace the example version and
artifact paths):

```powershell
Set-Location vscode-extension
npm ci
npm exec -- vsce login sbroenne
npm exec -- vsce publish --packagePath ..\artifacts\release\excel-mcp-2.1.0.vsix
npm exec -- vsce publish --packagePath ..\artifacts\release\excel-mcp-2.1.0-win32-arm64.vsix
```

Both packages use the existing Marketplace PAT authentication. On Windows
ARM64, the server uses Windows' x64 emulation; the ARM64 target describes the
VS Code installation, not a new native ARM64 server build.

## Security Best Practices

1. **Rotate tokens regularly** (every 6-12 months)
2. **Use minimal permissions** (only Marketplace Manage, not all scopes)
3. **Monitor secret usage** in GitHub Actions logs
4. **Revoke tokens immediately** if compromised
5. **Don't share tokens** via email, chat, or public channels

## References

- [VS Code Publishing Documentation](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
- [VS Code Extension Manager (vsce)](https://github.com/microsoft/vsce)
- [Azure DevOps PAT Documentation](https://learn.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate)
