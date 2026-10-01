---
applyTo: "vscode-extension/**"
excludeAgent: "code-review"
---

# Extension constraints

- The extension bundles self-contained Windows x64 and Apple Silicon macOS MCP
  Servers plus a generated skill; no separate user .NET install. Keep provider
  IDs and platform runtime paths in `package.json`, build scripts, and
  `src/extension.ts` aligned.
- `bin/`, `skills/excel-mcp/`, and extension `CHANGELOG.md` are packaging outputs.
  Edit repository sources instead.
- Do not bump versions for local tests: `npm version` can create commits/tags.
  The unified release workflow owns versions.
- Run `npm run compile` (includes metadata validation) and `npm run lint`.
  Packaging changes also require `npm run package` and VSIX content inspection.

Build, activation testing, and packaging procedures:
[DEVELOPMENT.md](../../DEVELOPMENT.md). Authorized publication:
[MARKETPLACE-PUBLISHING.md](../../MARKETPLACE-PUBLISHING.md).
