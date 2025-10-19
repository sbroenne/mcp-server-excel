# CodeQL Setup Issue - Quick Fix

## Error Message

```
Code Scanning could not process the submitted SARIF file:
CodeQL analyses from advanced configurations cannot be processed when the default setup is enabled
```

## Problem

The repository has **both** GitHub's automatic default CodeQL setup **and** a custom advanced CodeQL workflow. GitHub only allows one at a time.

## Quick Fix (Choose One)

### Option 1: Use Advanced Setup (Recommended)

This keeps the custom configuration with COM interop exceptions and path filters.

**Steps:**

1. Go to **Settings** → **Code security and analysis**
2. Find **"Code scanning"** section
3. Look for **"CodeQL analysis"**
4. If it shows **"Default"**:
   - Click the **"..."** (three dots) menu
   - Select **"Switch to advanced"**
5. If prompted, **do not create a new workflow** (we already have one)
6. Done! GitHub will now use `.github/workflows/codeql.yml`

**Benefits:**

- ✅ Custom COM interop false positive filters
- ✅ Path filters (only scans code changes, not docs)
- ✅ Custom query suites (security-extended)
- ✅ Windows runner for proper .NET build

### Option 2: Use Default Setup (Simpler)

GitHub manages everything, but you lose customization.

**Steps:**

1. Delete the custom workflow files:

   ```bash
   git rm .github/workflows/codeql.yml
   git rm -r .github/codeql/
   git commit -m "Remove custom CodeQL config for default setup"
   git push
   ```

2. Go to **Settings** → **Code security and analysis**
3. Click **"Set up"** for CodeQL analysis
4. Choose **"Default"**

**Trade-offs:**

- ❌ No custom exclusions for COM interop
- ❌ Scans on all file changes (including docs)
- ✅ Zero maintenance required
- ✅ Automatic updates from GitHub

## Verification

After switching to advanced:

1. Push a code change or trigger the workflow manually
2. Go to **Actions** tab
3. You should see **"CodeQL Advanced Security"** workflow running
4. No more SARIF processing errors

## Why This Happens

GitHub's default CodeQL setup (enabled through Settings) conflicts with custom workflows that use:

- `github/codeql-action/init@v3` with `config-file`
- Custom query packs
- Advanced configuration options

You must choose **one approach** - either fully automated (default) or fully customized (advanced).

## Recommended Choice

For **ExcelMcp/mcp-server-excel**: Use **Advanced Setup (Option 1)**

**Reasons:**

1. COM interop requires specific exclusions (weak crypto, unmanaged code)
2. Path filters save CI minutes (docs don't need code scanning)
3. Windows runner needed for proper .NET/Excel build
4. Custom query suites provide better coverage

## Need Help?

- See `.github/SECURITY_SETUP.md` for detailed security configuration guide
- Check GitHub's [CodeQL documentation](https://docs.github.com/en/code-security/code-scanning/automatically-scanning-your-code-for-vulnerabilities-and-errors/configuring-code-scanning)
