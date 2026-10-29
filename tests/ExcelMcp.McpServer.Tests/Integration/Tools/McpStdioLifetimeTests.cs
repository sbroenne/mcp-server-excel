using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("Sequential")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Integration")]
[Trait("Feature", "StdioLifetime")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Slow")]
public sealed class McpStdioLifetimeTests(ITestOutputHelper output)
{
    private const string Initialize = """{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"idle-regression","version":"1"}}}""";
    private const string Initialized = """{"jsonrpc":"2.0","method":"notifications/initialized"}""";
    private const string Ping = """{"jsonrpc":"2.0","id":2,"method":"ping"}""";
    private static readonly TimeSpan ResponseTimeout = TimeSpan.FromSeconds(15);

    [Theory(Timeout = 270000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Idle_ThreadsAndHandlesStayBounded_AndConnectionRemainsUsable(bool useDotnet)
    {
        await RunServerAsync(useDotnet, "idle", async (server, measurements) =>
        {
            await InitializeAsync(server.StandardInput, server.StandardOutput);
            measurements.Phase = "warmup";
            await Task.Delay(TimeSpan.FromSeconds(30));
            measurements.Phase = "idle";
            var baseline = measurements.Sample();
            await Task.Delay(TimeSpan.FromSeconds(180));
            var end = measurements.Sample();
            var idle = measurements.Samples.Where(sample => sample.Phase == "idle").ToArray();
            output.WriteLine(
                $"Idle: threads {baseline.Threads}->{end.Threads}, peak {idle.Max(s => s.Threads)}; " +
                $"handles {baseline.Handles}->{end.Handles}, peak {idle.Max(s => s.Handles)}.");

            measurements.Phase = "ping";
            await server.StandardInput.WriteLineAsync(Ping);
            using var pong = await ReadResponseAsync(server.StandardOutput, 2);
            Assert.Equal(JsonValueKind.Object, pong.RootElement.GetProperty("result").ValueKind);
            Assert.Empty(pong.RootElement.GetProperty("result").EnumerateObject());
            measurements.Phase = "disconnect";
            server.StandardInput.Close();
            await AssertCleanExitAsync(server);

            Assert.True(idle.Max(s => s.Threads) - baseline.Threads <= 5,
                "Idle threads grew by more than five; see the timestamped resource measurements.");
            Assert.True(idle.Max(s => s.Handles) - baseline.Handles <= 32,
                "Idle handles grew by more than 32; see the timestamped resource measurements.");
        });
    }

    [Theory(Timeout = 45000)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task InputCloses_ServerExitsSuccessfully(bool useDotnet, bool initialize)
    {
        await RunServerAsync(useDotnet, $"close-{initialize}", async (server, measurements) =>
        {
            if (initialize)
                await InitializeAsync(server.StandardInput, server.StandardOutput);
            measurements.Phase = "connected";
            await Task.Delay(TimeSpan.FromSeconds(6));
            measurements.Phase = "disconnect";
            server.StandardInput.Close();
            await AssertCleanExitAsync(server);
        });
    }

    [Theory(Timeout = 60000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ClientCrashes_IdleServerExitsSuccessfully(bool useDotnet)
    {
        const string script = """
            $ErrorActionPreference = 'Stop'
            $start = [Diagnostics.ProcessStartInfo]::new($env:EXCELMCP_TEST_SERVER)
            if ($env:EXCELMCP_TEST_SERVER_DLL) { $start.ArgumentList.Add($env:EXCELMCP_TEST_SERVER_DLL) }
            $start.UseShellExecute = $false
            $start.RedirectStandardInput = $true
            $start.RedirectStandardOutput = $true
            $start.RedirectStandardError = $true
            $server = [Diagnostics.Process]::Start($start)
            $errors = $server.StandardError.ReadToEndAsync()
            [Console]::WriteLine("$($server.Id),$($server.StartTime.ToUniversalTime().Ticks)")
            while ($null -ne ($line = [Console]::ReadLine())) {
                $server.StandardInput.WriteLine($line)
                $server.StandardInput.Flush()
                if (($line | ConvertFrom-Json).id) {
                    [Console]::WriteLine($server.StandardOutput.ReadLine())
                }
            }
            $server.StandardInput.Close()
            if (-not $server.WaitForExit(10000)) { throw 'Server did not exit after client input closed.' }
            [Console]::Error.Write($errors.GetAwaiter().GetResult())
            """;
        var product = CreateStartInfo(useDotnet);
        var start = new ProcessStartInfo("pwsh")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            ArgumentList = { "-NoProfile", "-NonInteractive", "-Command", script }
        };
        start.Environment["EXCELMCP_TELEMETRY_MODE"] = "off";
        start.Environment["EXCELMCP_TEST_SERVER"] = product.FileName;
        start.Environment["EXCELMCP_TEST_SERVER_DLL"] = useDotnet ? product.ArgumentList[0] : "";
        using var client = Process.Start(start)!;
        _ = client.Handle;
        var clientErrors = client.StandardError.ReadToEndAsync();
        Process? server = null;
        ResourceMeasurements? measurements = null;
        Exception? failure = null;
        try
        {
            var identity = (await client.StandardOutput.ReadLineAsync().WaitAsync(ResponseTimeout))!.Split(',');
            server = AcquireServer(
                int.Parse(identity[0], CultureInfo.InvariantCulture),
                long.Parse(identity[1], CultureInfo.InvariantCulture));
            measurements = new ResourceMeasurements(server, output, $"crash-{useDotnet}");
            await InitializeAsync(client.StandardInput, client.StandardOutput);
            measurements.Phase = "connected";
            await Task.Delay(TimeSpan.FromSeconds(12));
            measurements.Phase = "client-crash";
            client.Kill();
            await client.WaitForExitAsync().WaitAsync(ResponseTimeout);
            await server.WaitForExitAsync().WaitAsync(ResponseTimeout);
            Assert.Equal(0, server.ExitCode);
            Assert.Empty(await client.StandardOutput.ReadToEndAsync().WaitAsync(ResponseTimeout));
            Assert.Empty(await clientErrors.WaitAsync(ResponseTimeout));
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            var failures = new List<Exception>();
            if (failure is not null)
                failures.Add(failure);
            try
            {
                client.StandardInput.Close();
                await StopOwnedServerAsync(client);
            }
            catch (Exception cleanup)
            {
                failures.Add(cleanup);
            }
            try
            {
                if (server is not null)
                    await StopOwnedServerAsync(server);
            }
            catch (Exception cleanup)
            {
                failures.Add(cleanup);
            }
            try
            {
                _ = await clientErrors.WaitAsync(ResponseTimeout);
            }
            catch (Exception cleanup)
            {
                failures.Add(cleanup);
            }
            try
            {
                if (measurements is not null)
                    await measurements.DisposeAsync();
            }
            catch (Exception cleanup)
            {
                failures.Add(cleanup);
            }
            finally
            {
                server?.Dispose();
            }
            failure = failures.Count switch
            {
                0 => null,
                1 => failures[0],
                _ => new AggregateException(failures)
            };
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private async Task RunServerAsync(bool useDotnet, string scenario,
        Func<Process, ResourceMeasurements, Task> test)
    {
        using var server = Process.Start(CreateStartInfo(useDotnet))!;
        _ = server.Handle;
        var errors = server.StandardError.ReadToEndAsync();
        ResourceMeasurements? measurements = null;
        Exception? failure = null;
        try
        {
            measurements = new ResourceMeasurements(server, output, $"{scenario}-{useDotnet}");
            await test(server, measurements);
            Assert.Empty(await server.StandardOutput.ReadToEndAsync().WaitAsync(ResponseTimeout));
            Assert.Empty(await errors.WaitAsync(ResponseTimeout));
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                server.StandardInput.Close();
                await StopOwnedServerAsync(server);
                output.WriteLine(await errors.WaitAsync(ResponseTimeout));
            }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
            }
            try
            {
                if (measurements is not null)
                    await measurements.DisposeAsync();
            }
            catch (Exception cleanup)
            {
                failure = failure is null ? cleanup : new AggregateException(failure, cleanup);
            }
        }
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    private static ProcessStartInfo CreateStartInfo(bool useDotnet)
    {
        var executable = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\src\ExcelMcp.McpServer\bin\Release\net10.0-windows\Sbroenne.ExcelMcp.McpServer.exe"));
        Assert.True(File.Exists(executable), "Build the MCP server in Release before running process tests.");
        var start = new ProcessStartInfo(useDotnet ? "dotnet" : executable)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        if (useDotnet)
            start.ArgumentList.Add(Path.ChangeExtension(executable, ".dll"));
        start.Environment["EXCELMCP_TELEMETRY_MODE"] = "off";
        return start;
    }

    private static Process AcquireServer(int processId, long startedUtcTicks)
    {
        var candidate = Process.GetProcessById(processId);
        try
        {
            _ = candidate.Handle;
            Assert.Equal(startedUtcTicks, candidate.StartTime.ToUniversalTime().Ticks);
            return candidate;
        }
        catch
        {
            candidate.Dispose();
            throw;
        }
    }

    private static async Task InitializeAsync(StreamWriter input, StreamReader outputStream)
    {
        await input.WriteLineAsync(Initialize);
        using var response = await ReadResponseAsync(outputStream, 1);
        Assert.Equal("excel-mcp", response.RootElement.GetProperty("result").GetProperty("serverInfo").GetProperty("name").GetString());
        await input.WriteLineAsync(Initialized);
    }

    private static async Task<JsonDocument> ReadResponseAsync(StreamReader stream, int id)
    {
        var line = await stream.ReadLineAsync().WaitAsync(ResponseTimeout);
        Assert.NotNull(line);
        var response = JsonDocument.Parse(line);
        Assert.Equal("2.0", response.RootElement.GetProperty("jsonrpc").GetString());
        Assert.Equal(id, response.RootElement.GetProperty("id").GetInt32());
        Assert.False(response.RootElement.TryGetProperty("error", out _), line);
        return response;
    }

    private static async Task AssertCleanExitAsync(Process server)
    {
        await server.WaitForExitAsync().WaitAsync(ResponseTimeout);
        Assert.Equal(0, server.ExitCode);
    }

    private static async Task StopOwnedServerAsync(Process server)
    {
        try
        {
            await server.WaitForExitAsync().WaitAsync(ResponseTimeout);
        }
        catch (TimeoutException timeout)
        {
            // These tests send no Excel operations; only this retained, owned process is stopped.
            try
            {
                server.Kill();
                await server.WaitForExitAsync().WaitAsync(ResponseTimeout);
            }
            catch (Exception cleanup)
            {
                throw new AggregateException(timeout, cleanup);
            }
            throw;
        }
    }

    private sealed record ResourceSample(string Phase, int Threads, int Handles);

    private sealed class ResourceMeasurements : IAsyncDisposable
    {
        private readonly Process _server;
        private readonly ITestOutputHelper _output;
        private readonly StreamWriter _writer;
        private readonly CancellationTokenSource _stop = new();
        private readonly Task _monitor;
        private readonly Stopwatch _elapsed = Stopwatch.StartNew();
        private readonly List<ResourceSample> _samples = [];
        private readonly DateTime _startedUtc;
        private string _phase = "startup";

        public ResourceMeasurements(Process server, ITestOutputHelper output, string scenario)
        {
            _server = server;
            _output = output;
            _startedUtc = server.StartTime.ToUniversalTime();
            var directory = Environment.GetEnvironmentVariable("EXCELMCP_TEST_RESOURCE_LOG_DIR")
                ?? Path.Combine(Path.GetTempPath(), "excelmcp-resource-measurements");
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, $"{scenario}-{server.Id}-{Guid.NewGuid():N}.csv");
            _writer = new StreamWriter(path) { AutoFlush = true };
            _writer.WriteLine("utc,elapsedSeconds,pid,startedUtc,phase,threads,handles");
            output.WriteLine($"Resource measurements: {path}");
            Sample();
            _monitor = MonitorAsync();
        }

        public string Phase
        {
            set
            {
                Volatile.Write(ref _phase, value);
                SampleIfRunning();
            }
        }

        public ResourceSample[] Samples
        {
            get { lock (_samples) return [.. _samples]; }
        }

        public ResourceSample Sample()
        {
            lock (_samples)
            {
                _server.Refresh();
                var sample = new ResourceSample(Volatile.Read(ref _phase), _server.Threads.Count, _server.HandleCount);
                var line = FormattableString.Invariant(
                    $"{DateTime.UtcNow:O},{_elapsed.Elapsed.TotalSeconds:F3},{_server.Id},{_startedUtc:O},{sample.Phase},{sample.Threads},{sample.Handles}");
                _writer.WriteLine(line);
                _output.WriteLine(line);
                _samples.Add(sample);
                return sample;
            }
        }

        private async Task MonitorAsync()
        {
            using var timer = new PeriodicTimer(TimeSpan.FromSeconds(5));
            while (await timer.WaitForNextTickAsync(_stop.Token))
            {
                if (_server.HasExited)
                    return;
                SampleIfRunning();
            }
        }

        private void SampleIfRunning()
        {
            if (_server.HasExited)
                return;
            try
            {
                Sample();
            }
            catch (InvalidOperationException) when (_server.HasExited)
            {
                // Exit can race the resource read; the test separately checks successful exit.
            }
        }

        public async ValueTask DisposeAsync()
        {
            await _stop.CancelAsync();
            try
            {
                await _monitor;
            }
            catch (OperationCanceledException) when (_stop.IsCancellationRequested)
            {
                // Cancelling the measurement timer ends sampling, not the server.
            }
            finally
            {
                _output.WriteLine($"Server {_server.Id} exited: {_server.HasExited}");
                if (_server.HasExited)
                    _writer.WriteLine(FormattableString.Invariant(
                        $"{DateTime.UtcNow:O},{_elapsed.Elapsed.TotalSeconds:F3},{_server.Id},{_startedUtc:O},exited,0,0"));
                _stop.Dispose();
                await _writer.DisposeAsync();
            }
        }
    }
}
