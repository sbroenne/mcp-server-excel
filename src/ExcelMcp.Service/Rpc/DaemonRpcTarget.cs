namespace Sbroenne.ExcelMcp.Service.Rpc;

/// <summary>
/// Server-side RPC target that delegates incoming JSON-RPC calls to <see cref="ExcelMcpService.ProcessAsync"/>.
/// One instance is attached per pipe connection via <c>JsonRpc.Attach(stream, target)</c>.
/// </summary>
internal sealed class DaemonRpcTarget : IExcelDaemonRpc
{
    private readonly ExcelMcpService _service;

    public DaemonRpcTarget(ExcelMcpService service)
    {
        _service = service;
    }

    /// <inheritdoc />
    public async Task<ServiceResponse> ProcessCommandAsync(ServiceRequest request)
    {
        _service.RecordActivity();
        return await _service.ProcessAsync(request);
    }
}
