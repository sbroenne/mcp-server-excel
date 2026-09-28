namespace Sbroenne.ExcelMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to <see cref="ExcelMcpService.ProcessAsync"/>.
/// One instance is attached per pipe connection via <c>JsonRpc.Attach(stream, target)</c>.
/// </summary>
internal sealed class DaemonRpcTarget : IExcelDaemonRpc
{
    private readonly Func<ServiceRequest, Task<ServiceResponse>> _requestHandler;
    private readonly Action _recordActivity;

    internal DaemonRpcTarget(
        Func<ServiceRequest, Task<ServiceResponse>> requestHandler,
        Action recordActivity)
    {
        ArgumentNullException.ThrowIfNull(requestHandler);
        ArgumentNullException.ThrowIfNull(recordActivity);
        _requestHandler = requestHandler;
        _recordActivity = recordActivity;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
    {
        _recordActivity();
        return await _requestHandler(request);
    }
}
