# Development Workflow

## 🚨 **IMPORTANT: All Changes Must Use Pull Requests**

**Direct commits to `main` are not allowed.** All changes must go through the Pull Request (PR) process to ensure:

- Code review and quality control
- Proper version management
- CI/CD validation
- Documentation updates

## 📋 **Standard Development Workflow**

### 1. **Create Feature Branch**

```powershell
# Create and switch to feature branch
git checkout -b feature/your-feature-name

# Or for bug fixes
git checkout -b fix/issue-description

# Or for documentation updates  
git checkout -b docs/update-description
```

### 2. **Make Your Changes**

```powershell
# Make code changes, add tests, update docs
# Commit frequently with clear messages

git add .
git commit -m "Add feature X with tests and documentation

- Implement core functionality
- Add comprehensive unit tests  
- Update command documentation
- Include usage examples"
```

### 3. **Push Feature Branch**

```powershell
# Push your feature branch to GitHub
git push origin feature/your-feature-name
```

### 4. **Create Pull Request**

1. Go to [GitHub Repository](https://github.com/sbroenne/mcp-server-excel)
2. Click **"New Pull Request"**
3. Select your feature branch
4. Fill out the PR template:
   - **Clear title** describing the change
   - **Detailed description** of what was changed and why
   - **Testing information** - what tests were added/run
   - **Breaking changes** - if any
   - **Documentation updates** - what docs were updated

### 5. **PR Review Process**

- **Automated checks** will run (build, tests, linting)
- **Code review** by maintainers
- **Address feedback** if requested
- **Merge** once approved and all checks pass

### 6. **After Merge**

```powershell
# Switch back to main and pull latest
git checkout main
git pull origin main

# Delete the feature branch (cleanup)
git branch -d feature/your-feature-name
git push origin --delete feature/your-feature-name
```

## 🏷️ **Release Process**

### Creating a New Release

**Only maintainers** can create releases. The process is:

1. **Ensure all changes are merged** to `main` via PRs

2. **Create and push a version tag**:

```powershell
# Create version tag (semantic versioning)
git tag v1.1.0

# Push the tag (triggers release workflow)
git push origin v1.1.0
```

1. **Automated Release Workflow**:
   - ✅ Updates version numbers in project files
   - ✅ Builds the release binaries  
   - ✅ Creates GitHub release with ZIP file
   - ✅ Updates release notes

### Version Numbering

We follow [Semantic Versioning](https://semver.org/):

- **Major** (v2.0.0): Breaking changes
- **Minor** (v1.1.0): New features, backward compatible  
- **Patch** (v1.0.1): Bug fixes, backward compatible

## 🔒 **Branch Protection Rules**

The `main` branch is protected with:

- **Require pull request reviews** - Changes must be reviewed
- **Require status checks** - CI/CD must pass
- **Require up-to-date branches** - Must be current with main
- **No direct pushes** - All changes via PR only

## 🧪 **Testing Requirements**

Before creating a PR, ensure:

```powershell
# All tests pass
dotnet test

# Code builds without warnings  
dotnet build -c Release

# Code follows style guidelines (automatic via EditorConfig)
```

## 📝 **PR Template Checklist**

When creating a PR, verify:

- [ ] **Code builds** with zero warnings
- [ ] **All tests pass** (unit tests minimum)
- [ ] **New features have tests**
- [ ] **Documentation updated** (README, COMMANDS.md, etc.)
- [ ] **Breaking changes documented**
- [ ] **Follows existing code patterns**
- [ ] **Commit messages are clear**

## 🚫 **What NOT to Do**

- ❌ **Don't commit directly to `main`**
- ❌ **Don't create releases without PRs**
- ❌ **Don't skip tests**
- ❌ **Don't ignore build warnings**
- ❌ **Don't update version numbers manually** (release workflow handles this)

## 💡 **Tips for Good PRs**

### Commit Messages

```text
✅ Good: "Add PowerQuery batch refresh command with error handling"
❌ Bad: "fix stuff"
```

### PR Titles

```text  
✅ Good: "Add batch operations for Power Query refresh"
❌ Bad: "Update code"
```

### PR Size

- **Keep PRs focused** - One feature/fix per PR
- **Break large changes** into smaller, reviewable chunks
- **Include tests and docs** in the same PR as the feature

## 🔧 **Local Development Setup**

```powershell
# Clone the repository
git clone https://github.com/sbroenne/mcp-server-excel.git
cd ExcelMcp

# Install dependencies
dotnet restore

# Run all tests
dotnet test

# Build release version
dotnet build -c Release

# Test the built executable
.\src\ExcelMcp.CLI\bin\Release\net10.0\ExcelMcp.CLI.exe --version
```

## 📞 **Need Help?**

- **Read the docs**: [Contributing Guide](CONTRIBUTING.md)
- **Check command reference**: [Commands Documentation](COMMANDS.md)  
- **Ask questions**: Create a GitHub Issue with the `question` label
- **Report bugs**: Use the bug report template

---

**Remember: Every change, no matter how small, must go through a Pull Request!**

This ensures code quality, proper testing, and maintains the project's reliability for all users.