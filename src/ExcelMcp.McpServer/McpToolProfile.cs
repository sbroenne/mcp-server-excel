namespace Sbroenne.ExcelMcp.McpServer;

/// <summary>Registration profiles for the public MCP tool surface.</summary>
internal enum McpToolProfile
{
    Full,
    CopilotCompact,
}

/// <summary>
/// Defines stable, registration-time tool profiles. Profiles only control discovery/context;
/// they are not an authorization boundary and do not alter the service command catalog.
/// </summary>
internal static class McpToolProfileCatalog
{
    internal const string EnvironmentVariable = "EXCELMCP_TOOL_PROFILE";
    internal const string FullId = "full";
    internal const string CopilotCompactId = "copilot-compact";
    internal const string Version = "1";

    private static readonly string[] CopilotCompactTools =
    [
        "file",
        "workflow",
        "worksheet",
        "range",
        "range_edit",
        "range_format",
        "worksheet_style",
        "layout",
        "calculation_mode",
    ];

    private static readonly HashSet<string> CopilotCompactToolSet =
        new(CopilotCompactTools, StringComparer.Ordinal);

    private static readonly HashSet<string> CompactOnlyToolSet =
        new(["layout"], StringComparer.Ordinal);

    internal const string CompactServerInstructions = """
        ExcelMCP compact profile; live tools/list is authoritative.
        For 2+ edits prefer workflow(open-and-describe), then execute-plan; reuse session_id.
        Use layout for report formatting and outlines; direct domain tools for one-offs.
        Close with file(close) and an explicit save; ask before closing a visible workbook.
        Never replay a timeout or unknown outcome. Restart with EXCELMCP_TOOL_PROFILE=full for omitted tools.
        """;

    internal static IReadOnlyList<string> CompactTools => CopilotCompactTools;

    /// <summary>
    /// Missing and unrecognized values resolve to full for backward-compatible startup.
    /// A typo may forgo the optimization, but must never hide tools unexpectedly.
    /// </summary>
    internal static McpToolProfile Resolve(string? value) => value?.Trim().ToLowerInvariant() switch
    {
        CopilotCompactId or "compact" => McpToolProfile.CopilotCompact,
        _ => McpToolProfile.Full,
    };

    internal static string GetId(McpToolProfile profile) => profile switch
    {
        McpToolProfile.Full => FullId,
        McpToolProfile.CopilotCompact => CopilotCompactId,
        _ => throw new ArgumentOutOfRangeException(nameof(profile)),
    };

    internal static bool Includes(McpToolProfile profile, string toolName) => profile switch
    {
        McpToolProfile.Full => !CompactOnlyToolSet.Contains(toolName),
        McpToolProfile.CopilotCompact => CopilotCompactToolSet.Contains(toolName),
        _ => throw new ArgumentOutOfRangeException(nameof(profile)),
    };
}

/// <summary>
/// Process-scoped profile selected before MCP registration. The production server hosts one MCP
/// catalog; tests reset this value with the in-memory transport lifecycle.
/// </summary>
internal static class McpToolProfileRuntime
{
    private static int _current;

    internal static McpToolProfile Current => (McpToolProfile)Volatile.Read(ref _current);

    internal static void Configure(McpToolProfile profile) =>
        Volatile.Write(ref _current, (int)profile);

    internal static void Reset() => Configure(McpToolProfile.Full);
}
