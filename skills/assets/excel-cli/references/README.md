# Excel CLI References

This folder contains the generated command/action/flag reference plus shared Excel domain guidance for the CLI.

**Note for developers:** After the Release solution build, run `scripts\Build-AgentSkills.ps1 -GenerateOnly`. It renders complete skill directories under `artifacts\generated-skills`, including live CLI help and adapted shared guidance. Do not edit generated output.

**Note for users:** Start with `cli-commands.md` for the generated command index
and linked command pages. Shared domain guides contain native CLI examples
selected during generation; no translation from MCP calls is needed.
Use their platform notices to
distinguish Windows workflows from the experimental Apple Silicon macOS subset.

## Contents

- `cli-commands.md` - Auto-generated command groups, actions, parameters, and common pitfalls
- Shared domain guides - Excel workflows, gotchas, charts, screenshots, Power Query, Data Model, ranges, and other feature guidance
