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
5. **Packages as VSIX** file
6. **Publishes to VS Code Marketplace**
7. **Creates GitHub Release** with all components (MCP Server, CLI, VS Code, MCPB)

### Publishing is Best-Effort

The marketplace step uses `continue-on-error: true`, so a missing/expired token
or Marketplace outage does not block the GitHub release. Inspect the publish
step after every release and use the manual fallback below if needed.

## Testing the Workflow

To test the release workflow:

1. Ensure the `VSCE_TOKEN` secret is configured
2. Dispatch a real semantic release when it is ready:
   ```powershell
   gh workflow run release.yml -f version_bump=patch
   ```
3. Go to GitHub Actions and watch the unified workflow run
4. Check that the release was created and marketplace publishing succeeded

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

If automated publishing fails, you can publish manually:

```powershell
cd vscode-extension
npm install -g @vscode/vsce
vsce login <publisher-name>
vsce publish
```

## Security Best Practices

1. **Rotate tokens regularly** (every 6-12 months)
2. **Use minimal permissions** (only Marketplace Manage, not all scopes)
3. **Monitor secret usage** in GitHub Actions logs
4. **Revoke tokens immediately** if compromised
5. **Don't share tokens** via email, chat, or public channels

## References

- [VS Code Publishing Documentation](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
- [HaaLeo/publish-vscode-extension Action](https://github.com/marketplace/actions/publish-vs-code-extension)
- [Azure DevOps PAT Documentation](https://learn.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate)
