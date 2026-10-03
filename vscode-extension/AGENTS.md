# Extension constraints

Follow the [repository rules](../AGENTS.md). These are implementation
instructions; review tasks use the [shared review checks](../docs/agents/review.md).
Excel usage guidance comes from native MCP descriptions and repository docs.
The optional report-formatting skill is authored under `skills`; do not maintain
a separate tool reference in this extension.

- The extension bundles a self-contained Windows MCP Server and generated skill;
  no separate user .NET install. Keep provider IDs in `package.json` and
  `src/extension.ts` aligned.
- `bin/`, `skills/excel-mcp-report-formatting/`, and extension `CHANGELOG.md` are packaging outputs.
  Edit repository sources instead.
- Do not bump versions for local tests: `npm version` can create commits/tags.
  The unified release workflow owns versions.
- Run `npm run compile` (includes metadata validation) and `npm run lint`.
  For code or test changes, also run `npm test` and `npm run typecheck:tests`.
  Packaging changes also require `npm run package` and VSIX content inspection.

Build, activation testing, and packaging procedures:
[DEVELOPMENT.md](DEVELOPMENT.md). Authorized publication:
[MARKETPLACE-PUBLISHING.md](MARKETPLACE-PUBLISHING.md).
