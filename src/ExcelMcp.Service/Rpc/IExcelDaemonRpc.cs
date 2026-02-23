using PolyType;
using StreamJsonRpc;

namespace Sbroenne.ExcelMcp.Service.Rpc;

/// <summary>
/// Typed RPC interface for CLI↔daemon communication over named pipes.
/// Replaces the hand-rolled newline-delimited JSON protocol with StreamJsonRpc (JSON-RPC 2.0).
///
/// The interface has a single method that delegates to <see cref="ExcelMcpService.ProcessAsync"/>,
/// preserving the existing command routing while gaining:
/// - Proper error propagation (JSON-RPC error objects instead of swallowed exceptions)
/// - Standard protocol (JSON-RPC 2.0 with Content-Length framing)
/// - Type safety (compile-time contract between client and server)
/// - Testability via <c>Nerdbank.FullDuplexStream.CreatePair()</c>
/// </summary>
[JsonRpcContract]
[GenerateShape(IncludeMethods = MethodShapeFlags.PublicInstance)]
public partial interface IExcelDaemonRpc
{
    /// <summary>
    /// Sends a command to the daemon for execution.
    /// Wraps <see cref="ExcelMcpService.ProcessAsync"/> over the pipe transport.
    /// </summary>
    /// <param name="request">The service request with command, sessionId, and args.</param>
    /// <returns>The service response indicating success/failure with optional result data.</returns>
    Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request);
}
