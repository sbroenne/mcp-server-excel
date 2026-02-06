using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel VBA script management tool for MCP server.
/// Manages VBA macro operations, code import/export, and script execution in macro-enabled workbooks.
///
/// IMPORTANT: Requires .xlsm files! VBA operations only work with macro-enabled Excel files.
///
/// Prerequisites: VBA trust must be enabled for automation. Use setup-vba-trust command to configure.
/// </summary>
[McpServerToolType]
public static partial class ExcelVbaTool
{
    /// <summary>
    /// VBA scripts (requires .xlsm and VBA trust enabled).
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action (required for all VBA operations)</param>
    /// <param name="moduleName">VBA module name or procedure name (format: 'Module.Procedure' for run)</param>
    /// <param name="vbaCode">VBA code content as string (for import/update actions)</param>
    /// <param name="vbaCodeFile">Full path to .bas or .vba file containing VBA code. Alternative to vbaCode parameter - use for large/complex modules.</param>
    /// <param name="parameters">Parameters for VBA procedure execution (comma-separated)</param>
    [McpServerTool(Name = "excel_vba", Title = "Excel VBA Operations", Destructive = true)]
    [McpMeta("category", "automation")]
    [McpMeta("requiresSession", true)]
    [McpMeta("fileFormat", ".xlsm")]
    public static partial string ExcelVba(
        VbaAction action,
        string sessionId,
        [DefaultValue(null)] string? moduleName,
        [DefaultValue(null)] string? vbaCode,
        [DefaultValue(null)] string? vbaCodeFile,
        [DefaultValue(null)] string? parameters)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_vba",
            ServiceRegistry.Vba.ToActionString(action),
            () =>
            {
                // Parse comma-separated parameters into array for run action
                string[]? parameterArray = null;
                if (!string.IsNullOrWhiteSpace(parameters))
                {
                    parameterArray = parameters.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                               .Select(p => p.Trim())
                                               .ToArray();
                }

                return ServiceRegistry.Vba.RouteAction(
                    action,
                    sessionId,
                    ExcelToolsBase.ForwardToServiceFunc,
                    moduleName: moduleName,
                    vbaCode: vbaCode,
                    vbaCodeFile: vbaCodeFile,
                    procedureName: moduleName, // run action uses moduleName as procedureName
                    timeout: null,
                    parameters: parameterArray);
            });
    }
}





