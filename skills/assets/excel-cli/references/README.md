# Excel CLI References

This folder contains the generated command/action/flag reference plus shared Excel domain guidance for the CLI.

**Note for developers:** After the Release solution build, run `scripts\Build-AgentSkills.ps1 -GenerateOnly`. It renders complete skill directories under `artifacts\generated-skills`, including live CLI help and adapted shared guidance. Do not edit generated output.

**Note for users:** The exact CLI syntax comes from `cli-commands.md`. Shared domain guides may use MCP-style calls as conceptual shorthand; translate them according to the syntax notice at the top of each file.

## Contents

- `cli-commands.md` - Auto-generated command groups, actions, parameters, and common pitfalls
- Shared domain guides - Excel workflows, gotchas, charts, screenshots, Power Query, Data Model, ranges, and other feature guidance
