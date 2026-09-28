using System.Collections.Concurrent;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.McpServer.ServiceBridge;
using Sbroenne.ExcelMcp.Service;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

internal sealed class McpOwnedExcelProcesses : IDisposable
{
    private static readonly AsyncLocal<McpOwnedExcelProcesses?> RequestOwner = new();
    private readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _owned = new();

    public McpOwnedExcelProcesses() =>
        SessionManager.ExcelProcessIdentityTracked += OnProcessTracked;

    public IServiceBridgeBackend CreateBackend() => new TrackingBackend(this, new ExcelMcpService());

    private void OnProcessTracked(ExcelProcessIdentity identity)
    {
        if (ReferenceEquals(RequestOwner.Value, this))
        {
            _owned.TryAdd(identity, 0);
        }
    }

    public async Task AssertExitedAsync(ITestOutputHelper output, bool openedSession)
    {
        Assert.False(openedSession && _owned.IsEmpty,
            "A session opened without capturing its owned Excel PID/start-time identity.");
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
        ExcelProcessIdentity[] remaining;
        do
        {
            remaining = _owned.Keys.Where(identity => !OwnedProcessGuard.TryConfirmExited(identity)).ToArray();
            if (remaining.Length == 0)
            {
                return;
            }

            await Task.Delay(250);
        }
        while (DateTime.UtcNow < deadline);

        foreach (var identity in remaining)
        {
            var exited = OwnedProcessGuard.TryTerminate(identity, TimeSpan.Zero, TimeSpan.FromSeconds(5), out var terminated);
            output.WriteLine($"Owned Excel cleanup: PID {identity.ProcessId}, start {identity.StartedAtUtcFileTime}, terminated={terminated}, exited={exited}.");
        }

        throw new XunitException($"Owned Excel processes did not exit after MCP shutdown: {string.Join(", ", remaining)}.");
    }

    public void Dispose() =>
        SessionManager.ExcelProcessIdentityTracked -= OnProcessTracked;

    private sealed class TrackingBackend(McpOwnedExcelProcesses owner, ExcelMcpService service) : IServiceBridgeBackend
    {
        public async Task<ServiceResponse> ProcessAsync(ServiceRequest request)
        {
            var previous = RequestOwner.Value;
            RequestOwner.Value = owner;
            try
            {
                return await service.ProcessAsync(request);
            }
            finally
            {
                RequestOwner.Value = previous;
            }
        }

        public bool ForceCloseSession(string sessionId) =>
            service.SessionManager.CloseSession(sessionId, save: false, force: true);

        public void Dispose() => service.Dispose();
    }
}
