namespace Sbroenne.ExcelMcp.Build.Tasks;

/// <summary>
/// Model passed to Scriban templates for skill generation.
/// Properties are exposed to templates using snake_case naming (Scriban convention).
/// </summary>
public class SkillTemplateModel
{
    /// <summary>Version from Directory.Build.props</summary>
    public string Version { get; set; } = "";

    /// <summary>Number of CLI commands or MCP tools</summary>
    public int ToolCount { get; set; }

    /// <summary>Total number of operations/actions</summary>
    public int OperationCount { get; set; }

    /// <summary>Repository URL</summary>
    public string Repository { get; set; } = "https://github.com/sbroenne/mcp-server-excel";

    /// <summary>Documentation URL</summary>
    public string Documentation { get; set; } = "https://excelmcpserver.dev/";

    /// <summary>For CLI skill: parsed command reference from --help</summary>
    public List<CliCommand>? CliCommands { get; set; }
}

/// <summary>
/// Represents a CLI command parsed from excelcli --help output.
/// </summary>
public class CliCommand
{
    /// <summary>Command name (e.g., "worksheet", "range")</summary>
    public string Name { get; set; } = "";

    /// <summary>Command description from interface XML doc</summary>
    public string Description { get; set; } = "";

    /// <summary>List of actions (e.g., "list", "create", "delete")</summary>
    public List<string> Actions { get; set; } = new();

    /// <summary>List of parameters</summary>
    public List<CliParameter> Parameters { get; set; } = new();
}

/// <summary>
/// Represents a CLI parameter parsed from command --help output.
/// </summary>
public class CliParameter
{
    /// <summary>Parameter name without dashes (e.g., "sheet-name")</summary>
    public string Name { get; set; } = "";

    /// <summary>Description from help text</summary>
    public string Description { get; set; } = "";

    /// <summary>Whether parameter has short form (e.g., -s for --session)</summary>
    public string? ShortForm { get; set; }
}
