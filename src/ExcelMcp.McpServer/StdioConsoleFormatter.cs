using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Logging.Console;

namespace Sbroenne.ExcelMcp.McpServer;

internal sealed class StdioConsoleFormatter() : ConsoleFormatter(FormatterName)
{
    internal const string FormatterName = "excelmcp-stdio";

    public override void Write<TState>(
        in LogEntry<TState> logEntry,
        IExternalScopeProvider? scopeProvider,
        TextWriter textWriter)
    {
        if (logEntry.LogLevel == LogLevel.Warning
            && logEntry.Category == "ModelContextProtocol.Server.McpServer"
            && logEntry.Exception is OperationCanceledException)
        {
            return;
        }

        var message = logEntry.Formatter(logEntry.State, logEntry.Exception);
        if (string.IsNullOrEmpty(message) && logEntry.Exception is null)
            return;

        textWriter.Write(logEntry.LogLevel.ToString().ToLowerInvariant());
        textWriter.Write(": ");
        textWriter.Write(logEntry.Category);
        textWriter.Write('[');
        textWriter.Write(logEntry.EventId.Id);
        textWriter.WriteLine(']');

        if (!string.IsNullOrEmpty(message))
        {
            textWriter.Write("      ");
            textWriter.WriteLine(message);
        }

        if (logEntry.Exception is not null)
            textWriter.WriteLine(logEntry.Exception);
    }
}
