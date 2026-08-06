using System.ComponentModel;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>Compact Copilot actions for deterministic report layout.</summary>
[JsonConverter(typeof(JsonStringEnumConverter<LayoutAction>))]
public enum LayoutAction
{
    [JsonStringEnumMemberName("apply-report")]
    ApplyReport,

    [JsonStringEnumMemberName("get-report")]
    GetReport,

    [JsonStringEnumMemberName("set-outline")]
    SetOutline,

    [JsonStringEnumMemberName("get-outline")]
    GetOutline,
}

/// <summary>
/// Copilot-compact facade for deterministic report formatting and outlines. The full profile keeps
/// the complete report_format and outline tools; this facade avoids loading those broad
/// union schemas into every compact conversation.
/// </summary>
[McpServerToolType]
public static class ExcelLayoutTool
{
    [McpServerTool(Name = "layout", Title = "Deterministic Excel Layout", Destructive = true)]
    [McpMeta("category", "layout")]
    [McpMeta("requiresSession", true)]
    [Description("Apply/read deterministic report styles and outlines. Explicit ranges; no selection or toggles. Mutations return readback and preserve cell content.")]
    public static string ExcelLayout(
        [Description("Action")] LayoutAction action,
        [Description("Session ID")] string session_id,
        [Description("Sheet for report/outline")] string? sheet_name = null,
        string? title_range = null,
        [Description("Report header; required for report actions")] string? header_range = null,
        [Description("Report body; required for report actions")] string? body_range = null,
        string? total_range = null,
        [Description("professional or minimal")] string preset = "professional",
        [Description("Report accent as #RRGGBB")] string accent_color = "#1F4E78",
        bool auto_fit_columns = true,
        [Description("Rows 5:10 or columns B:D")] string? range_address = null,
        [Description("row or column")] string? axis = null,
        [Description("Outline level 0-7")] int level = 0,
        bool? collapsed = null,
        CancellationToken cancellationToken = default)
    {
        using var cancellationScope = ExcelToolsBase.PushCancellationToken(cancellationToken);
        var actionName = ToActionString(action);

        return ExcelToolsBase.ExecuteToolAction(
            "layout",
            actionName,
            () => action switch
            {
                LayoutAction.ApplyReport => ExcelToolsBase.ForwardToService(
                    "reportformat.apply",
                    session_id,
                    new
                    {
                        sheetName = Required(sheet_name, nameof(sheet_name), actionName),
                        titleRange = title_range,
                        headerRange = Required(header_range, nameof(header_range), actionName),
                        bodyRange = Required(body_range, nameof(body_range), actionName),
                        totalRange = total_range,
                        preset,
                        accentColor = accent_color,
                        autoFitColumns = auto_fit_columns,
                    }),
                LayoutAction.GetReport => ExcelToolsBase.ForwardToService(
                    "reportformat.get-state",
                    session_id,
                    new
                    {
                        sheetName = Required(sheet_name, nameof(sheet_name), actionName),
                        titleRange = title_range,
                        headerRange = Required(header_range, nameof(header_range), actionName),
                        bodyRange = Required(body_range, nameof(body_range), actionName),
                        totalRange = total_range,
                    }),
                LayoutAction.SetOutline => ExcelToolsBase.ForwardToService(
                    "outline.set-level",
                    session_id,
                    new
                    {
                        sheetName = Required(sheet_name, nameof(sheet_name), actionName),
                        rangeAddress = Required(range_address, nameof(range_address), actionName),
                        level,
                        axis = Required(axis, nameof(axis), actionName),
                        collapsed,
                    }),
                LayoutAction.GetOutline => ExcelToolsBase.ForwardToService(
                    "outline.get-state",
                    session_id,
                    new
                    {
                        sheetName = Required(sheet_name, nameof(sheet_name), actionName),
                        rangeAddress = Required(range_address, nameof(range_address), actionName),
                        axis = Required(axis, nameof(axis), actionName),
                    }),
                _ => throw new ArgumentOutOfRangeException(nameof(action), action, "Unknown layout action"),
            });
    }

    private static string Required(string? value, string parameterName, string action) =>
        !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new ArgumentException($"{parameterName} is required for {action}", parameterName);

    private static string ToActionString(LayoutAction action) => action switch
    {
        LayoutAction.ApplyReport => "apply-report",
        LayoutAction.GetReport => "get-report",
        LayoutAction.SetOutline => "set-outline",
        LayoutAction.GetOutline => "get-outline",
        _ => throw new ArgumentOutOfRangeException(nameof(action), action, "Unknown layout action"),
    };
}
