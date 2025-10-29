# Using VS Code Marketplace as a Publisher - Complete Guide

This guide walks you through publishing your VS Code extension to the marketplace for the first time.

## Quick Summary

1. Create Azure DevOps organization + PAT token
2. Create VS Code Marketplace publisher account
3. Configure GitHub secret (`VSCE_TOKEN`)
4. Push a version tag → Automated publishing!

---

## Step-by-Step First-Time Publishing

### Step 1: Create Microsoft/Azure DevOps Account

**Why needed:** VS Code Marketplace uses Azure DevOps for authentication.

1. **Go to Azure DevOps**: https://dev.azure.com/
2. **Sign in** with your Microsoft account (or create one at https://login.live.com/)
3. **Create a new organization** (if prompted):
   - Organization name: Can be anything (e.g., "MyPublishing")
   - Region: Choose closest to you
   - Click "Continue"

**✅ You now have an Azure DevOps organization**

---

### Step 2: Create Personal Access Token (PAT)

**Why needed:** The GitHub Actions workflow needs this to publish on your behalf.

1. **In Azure DevOps**, click your profile icon (top right) → **Personal Access Tokens**
2. Click **"+ New Token"**
3. **Configure the token:**
   - **Name**: `VS Code Marketplace Publishing`
   - **Organization**: Select your organization from dropdown
   - **Expiration**: Custom defined → 1 year (you'll need to renew annually)
   - **Scopes**: 
     - Click **"Custom defined"**
     - Scroll down to **"Marketplace"**
     - Check **"Manage"** (this includes Acquire and Publish)
   - **Important**: Only select Marketplace → Manage, nothing else needed
4. Click **"Create"**
5. **CRITICAL**: Copy the token immediately and save it securely
   - You'll only see this once!
   - If you lose it, you'll need to create a new one

**✅ You now have a PAT token** (keep it safe, you'll use it in Step 4)

---

### Step 3: Create VS Code Marketplace Publisher

**Why needed:** Your publisher identity on the marketplace.

1. **Go to Marketplace Management**: https://marketplace.visualstudio.com/manage
2. **Sign in** with the same Microsoft account from Step 1
3. Click **"Create publisher"**
4. **Fill in the form:**
   - **Publisher ID**: `sbroenne` (must match the `publisher` field in `package.json`)
     - ⚠️ This MUST be exactly what's in your package.json
     - ⚠️ Cannot be changed later
     - Can only contain letters, numbers, and hyphens
   - **Display name**: `Stefan Broenne` (shown to users)
   - **Description**: Brief description of your publisher account
   - **Logo** (optional): Upload a square image (recommended 128x128px)
5. Click **"Create"**

**✅ You now have a publisher account**

---

### Step 4: Configure GitHub Secret

**Why needed:** GitHub Actions needs the PAT to publish.

1. **Go to your GitHub repository**: https://github.com/sbroenne/mcp-server-excel
2. Click **Settings** (top right, near the repo name)
3. In left sidebar, click **Secrets and variables** → **Actions**
4. Click **"New repository secret"**
5. **Configure the secret:**
   - **Name**: `VSCE_TOKEN` (exactly this, case-sensitive)
   - **Value**: Paste the PAT from Step 2
6. Click **"Add secret"**

**✅ GitHub Actions can now publish to the marketplace**

---

### Step 5: Publish Your First Release

**Now you're ready to publish!**

1. **Update version in package.json** (if needed):
   ```bash
   cd vscode-extension
   npm version 1.0.0  # or patch, minor, major
   ```

2. **Update CHANGELOG.md** with release notes

3. **Commit changes**:
   ```bash
   git add .
   git commit -m "Prepare v1.0.0 release"
   git push
   ```

4. **Create and push version tag**:
   ```bash
   git tag vscode-v1.0.0
   git push origin vscode-v1.0.0
   ```

5. **Watch the magic happen**:
   - Go to **Actions** tab in GitHub
   - Watch the "Release VS Code Extension" workflow run
   - It will:
     - Build the extension
     - Publish to VS Code Marketplace
     - Create GitHub release

6. **Verify publication** (takes 5-15 minutes):
   - VS Code Marketplace: https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp
   - Or search "ExcelMcp" in VS Code Extensions panel

**✅ Your extension is now live on the marketplace!**

---

## Daily Publishing Workflow (After Setup)

Once you've done the above setup once, publishing new versions is easy:

```bash
# 1. Make your code changes
# ... edit files ...

# 2. Update version
cd vscode-extension
npm version patch  # or minor, or major

# 3. Update CHANGELOG.md
# Add release notes under new version

# 4. Commit and tag
git add .
git commit -m "Release v1.0.1"
git push

git tag vscode-v1.0.1
git push origin vscode-v1.0.1

# 5. Done! Automated workflow handles the rest
```

---

## Verifying Your Publisher Account

**Check your publisher page:**
- Go to https://marketplace.visualstudio.com/manage/publishers/sbroenne
- You should see your publisher details
- Any published extensions will appear here

**Check extension page:**
- Go to https://marketplace.visualstudio.com/items?itemName=sbroenne.excelmcp
- Should show your extension (after first publish)

---

## Common First-Time Issues

### ❌ "Publisher 'sbroenne' not found"

**Solution:** 
- Go to https://marketplace.visualstudio.com/manage
- Verify you created a publisher with ID `sbroenne` (exact match to package.json)
- Make sure you're signed in with the correct Microsoft account

### ❌ "Personal Access Token expired or invalid"

**Solution:**
- Create a new PAT following Step 2
- Update GitHub secret `VSCE_TOKEN` with new token

### ❌ "Extension validation failed"

**Solution:**
- Check `package.json` has all required fields:
  - `publisher`: Must match your marketplace publisher ID
  - `name`: Extension identifier
  - `displayName`: User-friendly name
  - `description`: Brief description
  - `version`: Semantic version
  - `engines.vscode`: VS Code version requirement
  - `repository`: GitHub repository URL
  - `license`: License type
- Icon must be PNG (not SVG)
- README images must use HTTPS URLs

### ❌ "Cannot publish - version already exists"

**Solution:**
- Increment version in package.json:
  ```bash
  npm version patch  # 1.0.0 → 1.0.1
  ```
- You cannot republish the same version

---

## Managing Your Publisher Account

### Update Publisher Details

1. Go to https://marketplace.visualstudio.com/manage/publishers/sbroenne
2. Click "Edit" to update:
   - Display name
   - Description
   - Logo
   - Contact links

### View Extension Statistics

1. Go to https://marketplace.visualstudio.com/manage/publishers/sbroenne
2. Click on your extension
3. See:
   - Download/install counts
   - Ratings and reviews
   - Usage trends

### Respond to Reviews

1. Go to your extension's marketplace page
2. Users can leave reviews and ratings
3. You can respond to reviews (good for support)

---

## Security Best Practices

1. **Never commit tokens** to Git (use GitHub secrets only)
2. **Set token expiration** to 1 year, renew before expiry
3. **Use minimum scopes** (only Marketplace → Manage)
4. **Rotate tokens annually** for security
5. **Monitor usage** in Azure DevOps PAT settings

---

## Getting Help

- **VS Code Publishing Docs**: https://code.visualstudio.com/api/working-with-extensions/publishing-extension
- **Azure DevOps PAT Docs**: https://learn.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate
- **Marketplace Publisher Portal**: https://marketplace.visualstudio.com/manage

---

## Quick Checklist for First Publish

- [ ] Created Azure DevOps organization
- [ ] Created Personal Access Token (PAT) with Marketplace → Manage scope
- [ ] Created VS Code Marketplace publisher (ID matches package.json)
- [ ] Added `VSCE_TOKEN` GitHub secret
- [ ] Verified package.json has all required fields
- [ ] Updated CHANGELOG.md with release notes
- [ ] Tagged release with `vscode-v*` format
- [ ] Watched GitHub Actions workflow succeed
- [ ] Verified extension appears on marketplace (wait 5-15 min)

**🎉 Once all checked, you're a published VS Code extension author!**
