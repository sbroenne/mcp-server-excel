# VS Code Marketplace Publishing Setup

This document explains how to set up automated publishing to the VS Code Marketplace and Open VSX Registry.

## Required GitHub Secrets

The release workflow requires two secrets to be configured in your GitHub repository:

### 1. VSCE_TOKEN (VS Code Marketplace)

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

### 2. OPEN_VSX_TOKEN (Open VSX Registry)

**Purpose:** Allows automated publishing to the Open VSX Registry (used by VS Codium, Gitpod, etc.)

**How to create:**

1. **Create an account on Open VSX**
   - Go to https://open-vsx.org/
   - Click "Sign In" → Use GitHub to sign in

2. **Generate an Access Token**
   - Once logged in, go to https://open-vsx.org/user-settings/tokens
   - Click "Generate New Token"
   - Description: `GitHub Actions Publishing`
   - Click "Generate"
   - **Copy the token** immediately

3. **Create a namespace** (if publishing for the first time)
   - Go to https://open-vsx.org/admin/create-namespace
   - Namespace: Should match your publisher name (e.g., `sbroenne`)
   - Description: Brief description of your publisher
   - Click "Create"

4. **Add to GitHub Secrets**
   - Go to your GitHub repo → Settings → Secrets and variables → Actions
   - Click "New repository secret"
   - Name: `OPEN_VSX_TOKEN`
   - Value: Paste your token from step 2
   - Click "Add secret"

## Workflow Behavior

When you push a tag matching `vscode-v*` (e.g., `vscode-v1.0.0`):

1. **Builds the extension** from source
2. **Packages as VSIX** file
3. **Publishes to VS Code Marketplace** (if `VSCE_TOKEN` is configured)
4. **Publishes to Open VSX Registry** (if `OPEN_VSX_TOKEN` is configured)
5. **Creates GitHub Release** with VSIX attachment and status of marketplace publishing

### Publishing is Optional

- If either token is not configured, that marketplace publishing step will be skipped (uses `continue-on-error: true`)
- The GitHub release will still be created with the VSIX file
- Users can always install from the VSIX file manually

### Publishing Status

The GitHub release notes will show:
```
### Publishing Status

- ✅ Published to VS Code Marketplace
- ✅ Published to Open VSX Registry
```

Or if tokens are not configured:
```
### Publishing Status

- ❌ Not published (to VS Code Marketplace)
- ❌ Not published (to Open VSX Registry)
```

## Testing the Workflow

To test the release workflow:

1. Ensure both secrets are configured (or at least one)
2. Push a test tag:
   ```bash
   git tag vscode-v0.0.1-test
   git push origin vscode-v0.0.1-test
   ```
3. Go to GitHub Actions and watch the workflow run
4. Check the release was created and marketplace publishing succeeded

## Troubleshooting

### "Failed to publish to VS Code Marketplace"

- **Check PAT permissions**: Ensure your Azure DevOps PAT has "Marketplace (Manage)" scope
- **Check PAT expiration**: Tokens expire - you may need to regenerate
- **Check publisher ownership**: Ensure your Azure DevOps account owns the publisher
- **Check package.json**: Publisher field must match your marketplace publisher ID

### "Failed to publish to Open VSX"

- **Check token validity**: Regenerate if needed at https://open-vsx.org/user-settings/tokens
- **Check namespace**: Ensure you've created a namespace matching your publisher name
- **Check package.json**: Publisher field must match your Open VSX namespace

### "Workflow runs but marketplaces show old version"

- Marketplace updates can take 5-15 minutes to appear
- Clear browser cache or use incognito mode
- Check marketplace directly:
  - VS Code: https://marketplace.visualstudio.com/items?itemName=PUBLISHER.EXTENSION
  - Open VSX: https://open-vsx.org/extension/PUBLISHER/EXTENSION

## Manual Publishing (Fallback)

If automated publishing fails, you can publish manually:

### VS Code Marketplace
```bash
cd vscode-extension
npm install -g @vscode/vsce
vsce login <publisher-name>
vsce publish
```

### Open VSX Registry
```bash
npm install -g ovsx
ovsx publish -p <your-token>
```

## Security Best Practices

1. **Rotate tokens regularly** (every 6-12 months)
2. **Use minimal permissions** (only Marketplace Manage, not all scopes)
3. **Monitor secret usage** in GitHub Actions logs
4. **Revoke tokens immediately** if compromised
5. **Don't share tokens** via email, chat, or public channels

## References

- [VS Code Publishing Documentation](https://code.visualstudio.com/api/working-with-extensions/publishing-extension)
- [Open VSX Publishing Documentation](https://github.com/eclipse/openvsx/wiki/Publishing-Extensions)
- [HaaLeo/publish-vscode-extension Action](https://github.com/marketplace/actions/publish-vs-code-extension)
- [Azure DevOps PAT Documentation](https://learn.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate)
