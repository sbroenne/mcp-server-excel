# Excel CLI References

This folder contains the generated command/action/flag reference plus shared Excel domain guidance for the CLI.

**Note for developers:** Run `dotnet build src\ExcelMcp.CLI\ExcelMcp.CLI.csproj -c Release` to generate SKILL.md. Run `scripts\Build-AgentSkills.ps1 -PopulateReferences` to regenerate `cli-commands.md` from the built CLI help output and adapt `skills\shared\*.md` for the CLI package.

**Note for users:** The exact CLI syntax comes from `cli-commands.md`. Shared domain guides may use MCP-style calls as conceptual shorthand; translate them according to the syntax notice at the top of each file.

## Contents

- `cli-commands.md` - Auto-generated command groups, actions, parameters, and common pitfalls
- Shared domain guides - Excel workflows, gotchas, charts, screenshots, Power Query, Data Model, ranges, and other feature guidance
