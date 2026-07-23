# Excel CLI References

This folder contains the generated command/action/flag reference for the Excel CLI.

**Note for developers:** Run `dotnet build src\ExcelMcp.CLI\ExcelMcp.CLI.csproj -c Release` to generate SKILL.md. Run `scripts\Build-AgentSkills.ps1 -PopulateReferences` to regenerate `cli-commands.md` from the built CLI help output.

**Note for users:** When distributed as a skill package, this folder includes the exact CLI command reference generated from `excelcli --help`.

## Contents

- `cli-commands.md` - Auto-generated command groups, actions, parameters, and common pitfalls
