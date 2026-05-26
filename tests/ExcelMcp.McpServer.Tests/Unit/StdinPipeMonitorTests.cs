using System.Runtime.InteropServices;
using Microsoft.Extensions.Hosting;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "StdinPipeMonitor")]
[Trait("Speed", "Fast")]
public sealed class StdinPipeMonitorTests
{
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CreatePipe(
        out IntPtr hReadPipe, out IntPtr hWritePipe,
        IntPtr lpPipeAttributes, uint nSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);

    [Fact]
    public void Start_WhenStdinIsPipe_ReturnsTimer()
    {
        // dotnet test spawns the test process with piped stdin,
        // so Start() should detect the pipe and return a timer.
        var lifetime = new FakeHostApplicationLifetime();

        var timer = StdinPipeMonitor.Start(lifetime);

        try
        {
            Assert.NotNull(timer);
        }
        finally
        {
            timer?.Dispose();
        }
    }

    [Fact]
    public async Task StartCore_WhenPipeAlreadyBroken_CallsStopApplication()
    {
        Assert.True(CreatePipe(out var readHandle, out var writeHandle, IntPtr.Zero, 0));
        CloseHandle(writeHandle);

        var lifetime = new FakeHostApplicationLifetime();

        using var timer = StdinPipeMonitor.StartCore(lifetime, readHandle,
            pollInterval: TimeSpan.FromMilliseconds(50));

        try
        {
            await WaitForConditionAsync(
                () => lifetime.StopCallCount > 0,
                timeout: TimeSpan.FromSeconds(2));

            Assert.True(lifetime.StopCallCount > 0,
                "StopApplication should be called when the pipe is already broken at start");
        }
        finally
        {
            CloseHandle(readHandle);
        }
    }

    [Fact]
    public void StartCore_WithValidPipeHandle_ReturnsTimer()
    {
        Assert.True(CreatePipe(out var readHandle, out var writeHandle, IntPtr.Zero, 0));
        try
        {
            var lifetime = new FakeHostApplicationLifetime();

            using var timer = StdinPipeMonitor.StartCore(lifetime, readHandle,
                pollInterval: TimeSpan.FromMilliseconds(50));

            Assert.NotNull(timer);
        }
        finally
        {
            CloseHandle(readHandle);
            CloseHandle(writeHandle);
        }
    }

    [Fact]
    public async Task StartCore_WhenWriteEndCloses_CallsStopApplication()
    {
        Assert.True(CreatePipe(out var readHandle, out var writeHandle, IntPtr.Zero, 0));
        var lifetime = new FakeHostApplicationLifetime();

        using var timer = StdinPipeMonitor.StartCore(lifetime, readHandle,
            pollInterval: TimeSpan.FromMilliseconds(50));

        CloseHandle(writeHandle);

        try
        {
            await WaitForConditionAsync(
                () => lifetime.StopCallCount > 0,
                timeout: TimeSpan.FromSeconds(2));

            Assert.True(lifetime.StopCallCount > 0,
                "StopApplication should be called when the pipe breaks");
        }
        finally
        {
            CloseHandle(readHandle);
        }
    }

    [Fact]
    public async Task StartCore_WhenPipeIsHealthy_DoesNotCallStopApplication()
    {
        Assert.True(CreatePipe(out var readHandle, out var writeHandle, IntPtr.Zero, 0));
        var lifetime = new FakeHostApplicationLifetime();

        using var timer = StdinPipeMonitor.StartCore(lifetime, readHandle,
            pollInterval: TimeSpan.FromMilliseconds(50));

        try
        {
            await Task.Delay(300);
            Assert.Equal(0, lifetime.StopCallCount);
        }
        finally
        {
            CloseHandle(readHandle);
            CloseHandle(writeHandle);
        }
    }

    [Fact]
    public async Task StartCore_AfterDispose_DoesNotCallStopOnPipeBreak()
    {
        Assert.True(CreatePipe(out var readHandle, out var writeHandle, IntPtr.Zero, 0));
        var lifetime = new FakeHostApplicationLifetime();

        var timer = StdinPipeMonitor.StartCore(lifetime, readHandle,
            pollInterval: TimeSpan.FromMilliseconds(50));

        timer.Dispose();
        await Task.Delay(100);

        CloseHandle(writeHandle);

        try
        {
            await Task.Delay(300);
            Assert.Equal(0, lifetime.StopCallCount);
        }
        finally
        {
            CloseHandle(readHandle);
        }
    }

    private static async Task WaitForConditionAsync(Func<bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (!condition() && DateTime.UtcNow < deadline)
            await Task.Delay(25);
    }

    private sealed class FakeHostApplicationLifetime : IHostApplicationLifetime
    {
        private int _stopCallCount;

        public int StopCallCount => Volatile.Read(ref _stopCallCount);

        public CancellationToken ApplicationStarted => CancellationToken.None;
        public CancellationToken ApplicationStopping => CancellationToken.None;
        public CancellationToken ApplicationStopped => CancellationToken.None;

        public void StopApplication() => Interlocked.Increment(ref _stopCallCount);
    }
}
