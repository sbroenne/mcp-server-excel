using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.McpServer.Tools;

namespace Sbroenne.ExcelMcp.McpServer.Status;

/// <summary>
/// Lightweight local IPC for VS Code extension status polling.
///
/// This intentionally exposes only telemetry-like operations (list sessions) and
/// user-initiated session management (close) without going through VS Code's tool
/// invocation flow, which can trigger confirmation UI when no tool token exists.
///
/// Protocol: client connects to named pipe, sends a single JSON line request,
/// receives a single JSON line response, then disconnects.
/// </summary>
public sealed partial class ExcelMcpStatusPipeService : BackgroundService
{
    public const string PipeNameEnvVar = "EXCELMCP_STATUS_PIPE_NAME";

    private readonly ILogger<ExcelMcpStatusPipeService> _logger;

    public ExcelMcpStatusPipeService(ILogger<ExcelMcpStatusPipeService> logger)
    {
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var pipeName = Environment.GetEnvironmentVariable(PipeNameEnvVar);
        if (string.IsNullOrWhiteSpace(pipeName))
        {
            Log.StatusPipeDisabled(_logger, PipeNameEnvVar);
            return;
        }

        // Serve connections sequentially. This is a lightweight endpoint for polling.
        // ACL is enforced at the OS level when the pipe is created.
        Log.StatusPipeCreated(_logger, WindowsIdentity.GetCurrent().Name ?? "unknown");
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await using var pipe = CreatePipe(pipeName);
                await pipe.WaitForConnectionAsync(stoppingToken);

                var requestJson = await ReadSingleLineAsync(pipe, stoppingToken);
                var responseJson = HandleRequest(requestJson);

                var responseBytes = Encoding.UTF8.GetBytes(responseJson + "\n");
                await pipe.WriteAsync(responseBytes, stoppingToken);
                await pipe.FlushAsync(stoppingToken);
            }
            catch (OperationCanceledException)
            {
                // normal shutdown
                break;
            }
            catch (Exception ex)
            {
                // Best-effort: log and continue so polling can recover.
                Log.StatusPipeServerError(_logger, ex);
                await Task.Delay(250, stoppingToken);
            }
        }
    }

    private static partial class Log
    {
        [LoggerMessage(EventId = 8401, Level = LogLevel.Debug, Message = "Status pipe disabled: env var {EnvVar} not set")]
        public static partial void StatusPipeDisabled(ILogger logger, string envVar);

        [LoggerMessage(EventId = 8402, Level = LogLevel.Warning, Message = "Status pipe server error")]
        public static partial void StatusPipeServerError(ILogger logger, Exception ex);

        [LoggerMessage(EventId = 8403, Level = LogLevel.Debug, Message = "Status pipe created with restricted ACL for user {UserName}")]
        public static partial void StatusPipeCreated(ILogger logger, string userName);
    }

    private static NamedPipeServerStream CreatePipe(string pipeName)
    {
        // Restrict pipe access to the current user only. This prevents unprivileged processes
        // on the same machine from connecting to the status pipe.
        var identity = WindowsIdentity.GetCurrent();
        if (identity.User == null)
        {
            throw new InvalidOperationException("Unable to determine current Windows identity for named pipe ACL");
        }

        var pipeSecurity = new PipeSecurity();
        // Allow read/write only for the current user
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            identity.User,
            PipeAccessRights.ReadWrite,
            AccessControlType.Allow));

        // Explicitly deny everyone else (defense in depth)
        var everyoneSid = new SecurityIdentifier(WellKnownSidType.WorldSid, null);
        pipeSecurity.AddAccessRule(new PipeAccessRule(
            everyoneSid,
            PipeAccessRights.FullControl,
            AccessControlType.Deny));

        return NamedPipeServerStreamAcl.Create(
            pipeName,
            PipeDirection.InOut,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            0,
            0,
            pipeSecurity);
    }

    private static async Task<string?> ReadSingleLineAsync(Stream stream, CancellationToken cancellationToken)
    {
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);

        // If the client sends nothing, treat as implicit list.
        // StreamReader.ReadLineAsync(CancellationToken) exists in .NET 8.
        return await reader.ReadLineAsync(cancellationToken);
    }

    private static string HandleRequest(string? requestJson)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(requestJson))
            {
                return ListSessions();
            }

            using var doc = JsonDocument.Parse(requestJson);
            var root = doc.RootElement;

            var action = root.TryGetProperty("action", out var a) && a.ValueKind == JsonValueKind.String
                ? a.GetString()
                : "list";

            return action?.ToLowerInvariant() switch
            {
                "list" => ListSessions(),
                "close" => CloseSession(root),
                _ => JsonSerializer.Serialize(new
                {
                    success = false,
                    errorMessage = $"Unknown action '{action}'. Supported actions: list, close",
                    isError = true
                }, ExcelToolsBase.JsonOptions)
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"Status pipe request failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ListSessions()
    {
        var sessionManager = ExcelToolsBase.GetSessionManager();
        var sessions = sessionManager.GetActiveSessions();

        var sessionList = sessions.Select(s => new
        {
            sessionId = s.SessionId,
            filePath = s.FilePath,
            activeOperations = sessionManager.GetActiveOperationCount(s.SessionId),
            isExcelVisible = sessionManager.IsExcelVisible(s.SessionId),
            canClose = sessionManager.GetActiveOperationCount(s.SessionId) == 0
        }).ToList();

        return JsonSerializer.Serialize(new
        {
            success = true,
            sessions = sessionList,
            count = sessionList.Count
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CloseSession(JsonElement root)
    {
        var sessionId = root.TryGetProperty("sessionId", out var sid) && sid.ValueKind == JsonValueKind.String
            ? sid.GetString()
            : null;

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = "sessionId is required for close action",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }

        var save = root.TryGetProperty("save", out var s) && (s.ValueKind == JsonValueKind.True || s.ValueKind == JsonValueKind.False)
            ? s.GetBoolean()
            : false;

        var sessionManager = ExcelToolsBase.GetSessionManager();

        try
        {
            var closed = sessionManager.CloseSession(sessionId, save);
            return JsonSerializer.Serialize(new
            {
                success = closed,
                sessionId,
                saved = save,
                isError = closed ? (bool?)null : true,
                errorMessage = closed ? (string?)null : $"Session '{sessionId}' not found"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                sessionId,
                saved = save,
                isError = true,
                errorMessage = ex.Message
            }, ExcelToolsBase.JsonOptions);
        }
    }
}
