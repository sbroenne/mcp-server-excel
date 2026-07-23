using ModelContextProtocol;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.McpServer.Progress;

/// <summary>
/// Adapts the MCP SDK's <see cref="IProgress{ProgressNotificationValue}"/> to our
/// domain <see cref="IProgress{ProgressInfo}"/> so Core commands stay MCP-agnostic.
/// </summary>
internal sealed class McpProgressAdapter : IProgress<ProgressInfo>
{
    private readonly IProgress<ProgressNotificationValue> _mcpProgress;

    public McpProgressAdapter(IProgress<ProgressNotificationValue> mcpProgress)
    {
        _mcpProgress = mcpProgress;
    }

    public void Report(ProgressInfo value)
    {
        _mcpProgress.Report(new ProgressNotificationValue
        {
            Progress = value.Current,
            Total = value.Total,
            Message = value.Message
        });
    }
}
