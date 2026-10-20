namespace Sbroenne.ExcelMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to <see cref="ExcelMcpService.ProcessAsync"/>.
/// One instance is attached per pipe connection.
/// </summary>
internal sealed class DaemonRpcTarget : IExcelDaemonRpc
{
    private readonly Func<ServiceRequest, Task<ServiceResponse>> _requestHandler;
    private readonly Action<ServiceRequest, ServiceResponse?>? _requestAbandoned;

    internal DaemonRpcTarget(
        Func<ServiceRequest, Task<ServiceResponse>> requestHandler,
        Action<ServiceRequest, ServiceResponse?>? requestAbandoned = null)
    {
        ArgumentNullException.ThrowIfNull(requestHandler);
        _requestHandler = requestHandler;
        _requestAbandoned = requestAbandoned;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        request.CallerCancellation = cancellationToken;
        using var registration = cancellationToken.Register(() => _requestAbandoned?.Invoke(request, null));
        var response = await _requestHandler(request);
        if (cancellationToken.IsCancellationRequested && string.IsNullOrWhiteSpace(request.SessionId))
            _requestAbandoned?.Invoke(request, response);
        return response;
    }
}
