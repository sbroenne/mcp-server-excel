using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Service.Rpc;
using Sbroenne.ExcelMcp.Service.Workers;
using StreamJsonRpc;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("Sequential")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceDaemon")]
[Trait("Speed", "Medium")]
[Trait("RequiresExcel", "false")]
public sealed class ExcelControllerTests
{
    private static ProductProcessLauncher Launcher => new(Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);

    [Theory(Timeout = 30000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_CompletionPublishesNewerReleaseBeforeAcknowledgingIt(bool recordingFails)
    {
        ExcelWorkerClient? started = null;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            started = worker;
        });
        Assert.True((await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"initialize-completion-record"}"""
        })).Success);
        var worker = Assert.IsType<ExcelWorkerClient>(started);
        var older = worker.Snapshot! with { Revision = long.MaxValue - 2 };
        var identity = new ExcelProcessIdentity(WorkerProcess.Current.ProcessId, WorkerProcess.Current.StartedAtUtcFileTime);
        var newer = older with
        {
            Registration = older.Registration with { ReleaseConfirmed = true, ReleasedExcel = identity },
            Releasing = true,
            Revision = older.Revision + 1
        };
        worker.RecordSnapshot(older);
        worker.RecordSnapshot(newer);
        var observed = new List<ExcelProcessIdentity>();
        var failure = new IOException("Test-only release publication failure.");
        void Record(ExcelProcessIdentity released)
        {
            observed.Add(released);
            if (recordingFails)
                throw failure;
        }
        controller.ExcelProcessReleased += Record;
        try
        {
            var apply = typeof(ExcelController).GetMethod("ApplyWorkerSnapshot",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
            if (recordingFails)
                Assert.Same(failure, Assert.Throws<System.Reflection.TargetInvocationException>(
                    () => apply.Invoke(controller, [worker])).InnerException);
            else
                Assert.Same(newer, apply.Invoke(controller, [worker]));
            Assert.Equal(identity, Assert.Single(observed));
            Assert.Same(newer, worker.Snapshot);
        }
        finally
        {
            controller.ExcelProcessReleased -= Record;
        }
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_QueryDistinctUrls_AreReservedSeparately()
    {
        var firstStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirst = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var starts = 0;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            if (Interlocked.Increment(ref starts) == 1)
            {
                firstStarted.SetResult();
                await releaseFirst.Task;
            }
        });
        const string endpoint = "https://example.com/export.xlsx";
        Task<ServiceResponse> Open(string query) => controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = endpoint + query, timeoutSeconds = 1 })
        });
        var first = Open("?report=one&token=test-first");
        ServiceResponse firstResponse;
        try
        {
            await firstStarted.Task.WaitAsync(TimeSpan.FromSeconds(10));
            var duplicate = await Open("?report=one&token=test-first");
            Assert.False(duplicate.Success);
            Assert.Equal("Conflict", duplicate.ErrorCategory);
            var distinct = await Open("?report=two&token=test-second");
            Assert.False(distinct.Success);
            Assert.Equal("ArgumentOutOfRangeException", distinct.ExceptionType);
            Assert.DoesNotContain("test-first", JsonSerializer.Serialize(duplicate));
            Assert.DoesNotContain("test-second", JsonSerializer.Serialize(distinct));
        }
        finally
        {
            releaseFirst.TrySetResult();
            firstResponse = await first;
        }
        Assert.False(firstResponse.Success);
        Assert.Equal("ArgumentOutOfRangeException", firstResponse.ExceptionType);
    }

    [Theory(Timeout = 30000)]
    [InlineData("")]
    [InlineData("?report=one&token=test-private")]
    public async Task Controller_TracksObservedPersonalOneDriveLocationAndPreservesItsReservation(string query)
    {
        ExcelWorkerClient? started = null;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            started = worker;
        });
        Assert.True((await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"initialize-cloud-status-owner"}"""
        })).Success);
        var worker = Assert.IsType<ExcelWorkerClient>(started);
        var session = Guid.NewGuid().ToString("N");
        const string location = "https://d.docs.live.net/123/Documents/Test%20File.xlsx";
        worker.RecordSnapshot(new WorkerSnapshot(worker.Registration,
            [new SessionDescriptor(session, location)],
            new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    sessions = new[] { new { sessionId = session, filePath = location,
                        filePaths = new[] { location }, operationPending = false } }
                })
            }, null, null,
            new Dictionary<string, IReadOnlyList<string>>
            {
                [session] = [WorkbookLocation.GetOwnershipKey(location + query)]
            }, Revision: long.MaxValue));
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (controller.Sessions.SingleOrDefault()?.FilePath != location)
            await Task.Delay(50, timeout.Token);

        var response = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            SessionId = session,
            Args = """{"message":"retain-observed-cloud-location"}"""
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.Null(worker.StatusError);
        Assert.Equal(location, Assert.Single(controller.Sessions).FilePath);
        var rejected = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = location + query })
        });
        Assert.False(rejected.Success);
        Assert.Contains("already", rejected.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(session, Assert.Single(controller.Sessions).SessionId);
        var listed = await controller.ProcessAsync(new ServiceRequest { Command = "session.list" });
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.DoesNotContain("test-private", listed.Result);
        Assert.DoesNotContain("ownershipKeys", listed.Result, StringComparison.OrdinalIgnoreCase);
        if (query.Length > 0)
        {
            var other = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = location + "?report=two", timeoutSeconds = 1 })
            });
            Assert.False(other.Success);
            Assert.Equal("ArgumentOutOfRangeException", other.ExceptionType);
        }
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_TracksEveryReportedWorkbookPathAcrossCompletionAndRename()
    {
        ExcelWorkerClient? started = null;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            started = worker;
        });
        Assert.True((await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"initialize-status-owner"}"""
        })).Success);
        var worker = Assert.IsType<ExcelWorkerClient>(started);
        var session = Guid.NewGuid().ToString("N");
        var source = Path.Join(Path.GetTempPath(), $"source-{Guid.NewGuid():N}.xlsx");
        var target = Path.Join(Path.GetTempPath(), $"target-{Guid.NewGuid():N}.xlsx");
        var renamed = Path.Join(Path.GetTempPath(), $"renamed-{Guid.NewGuid():N}.xlsx");
        Publish(source, long.MaxValue - 1);
        await WaitForPathAsync(source);
        await AssertReservedAsync(source);
        await AssertReservedAsync(target);
        var control = new ServiceRequest
        {
            Command = "diag.echo",
            SessionId = session,
            Args = """{"message":"retain-both-paths"}"""
        };
        var listed = await controller.ProcessAsync(control);
        Assert.True(listed.Success, listed.ErrorMessage);
        await AssertReservedAsync(target);
        Publish(renamed, long.MaxValue);
        Assert.True((await controller.ProcessAsync(control)).Success);
        await WaitForPathAsync(renamed);
        await AssertReservedAsync(renamed);
        await AssertReservedAsync(target);
        var available = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = source, timeoutSeconds = 1 })
        });
        Assert.False(available.Success);
        Assert.Equal("ArgumentOutOfRangeException", available.ExceptionType);

        void Publish(string primary, long revision) =>
            worker.RecordSnapshot(new WorkerSnapshot(worker.Registration,
                [new SessionDescriptor(session, primary)],
                new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new
                    {
                        sessions = new[] { new { sessionId = session, filePath = primary,
                            filePaths = new[] { primary, target }, operationPending = false } }
                    })
                }, null, null,
                new Dictionary<string, IReadOnlyList<string>> { [session] = [primary, target] },
                Revision: revision));

        async Task WaitForPathAsync(string expected)
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            while (controller.Sessions.SingleOrDefault()?.FilePath != expected)
                await Task.Delay(50, timeout.Token);
        }

        async Task AssertReservedAsync(string path)
        {
            foreach (var command in new[] { "session.open", "session.test", "sheet.copy-to-file", "sheet.move-to-file" })
            {
                var other = Path.Join(Path.GetTempPath(), $"other-{Guid.NewGuid():N}.xlsx");
                var response = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    Args = command.StartsWith("sheet.", StringComparison.Ordinal)
                        ? JsonSerializer.Serialize(new
                        {
                            sourceFile = command == "sheet.copy-to-file" ? path : other,
                            targetFile = command == "sheet.move-to-file" ? path : other,
                            sourceSheet = "Sheet1"
                        })
                        : JsonSerializer.Serialize(new { filePath = path })
                });
                Assert.False(response.Success);
                Assert.Equal("Conflict", response.ErrorCategory);
                Assert.Equal("This workbook is already reserved by another session or pending request.", response.ErrorMessage);
                if (command == "sheet.move-to-file")
                {
                    var retry = await controller.ProcessAsync(new ServiceRequest
                    {
                        Command = "session.open",
                        Args = JsonSerializer.Serialize(new { filePath = other, timeoutSeconds = 1 })
                    });
                    Assert.False(retry.Success);
                    Assert.Equal("ArgumentOutOfRangeException", retry.ExceptionType);
                }
            }
        }
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_ValidationReleasesItsReservationWhenNoWorkbookRemainsOpen()
    {
        using var controller = new ExcelController(launcher: Launcher);
        var path = Path.Join(Path.GetTempPath(), $"validation-{Guid.NewGuid():N}.xlsx");
        var validated = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.test",
            Args = JsonSerializer.Serialize(new { filePath = path })
        });
        Assert.True(validated.Success, validated.ErrorMessage);
        using var result = JsonDocument.Parse(validated.Result!);
        Assert.False(result.RootElement.GetProperty("isValid").GetBoolean());
        var opened = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 1 })
        });
        Assert.False(opened.Success);
        Assert.Equal("ArgumentOutOfRangeException", opened.ExceptionType);
    }

    [Theory(Timeout = 30000)]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Controller_AbandonedExitedWorkerKeepsReservationsForSurvivingExcel(bool snapshotHasExcel)
    {
        ExcelWorkerClient? abandoned = null;
        var session = Guid.NewGuid().ToString("N");
        var path = Path.Join(Path.GetTempPath(), $"surviving-{Guid.NewGuid():N}.xlsx");
        var surviving = WorkerProcess.Current;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            worker.RecordSnapshot(new WorkerSnapshot(
                worker.Registration with
                {
                    Excel = snapshotHasExcel ? new ExcelProcessIdentity(surviving.ProcessId, surviving.StartedAtUtcFileTime) : null
                },
                [new SessionDescriptor(session, path)],
                new ServiceResponse
                {
                    Success = true,
                    Result = JsonSerializer.Serialize(new
                    {
                        sessions = new[] { new { sessionId = session, filePath = path, filePaths = new[] { path }, operationPending = false } }
                    })
                }, null, null,
                new Dictionary<string, IReadOnlyList<string>> { [session] = [path] },
                Revision: long.MaxValue));
            abandoned = worker;
            throw new TimeoutException("The startup reply was not confirmed.");
        });
        var started = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path })
        });
        Assert.False(started.Success);
        Assert.Equal("Timeout", started.ErrorCategory);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        while (controller.Sessions.Count == 0)
            await Task.Delay(50, timeout.Token);
        var worker = Assert.IsType<ExcelWorkerClient>(abandoned);
        using var process = System.Diagnostics.Process.GetProcessById(worker.Registration.Worker.ProcessId);
        _ = process.Handle;
        Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
        Assert.True((await worker.ReleaseAsync(CancellationToken.None)).Success);
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        for (var index = 0; index < 25; index++)
        {
            Assert.Equal(session, Assert.Single(controller.Sessions).SessionId);
            var retry = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.False(retry.Success);
            Assert.Equal("Conflict", retry.ErrorCategory);
            await Task.Delay(100, timeout.Token);
        }
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_ExitedWorkerWithoutSnapshotKeepsDispatchedRequestPath()
    {
        System.Diagnostics.Process? process = null;
        var initialized = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: async worker =>
        {
            await worker.InitializeAsync();
            if (process is null)
            {
                Assert.Null(worker.Registration.Excel);
                Assert.Empty(worker.Snapshot!.Sessions);
                process = System.Diagnostics.Process.GetProcessById(worker.Registration.Worker.ProcessId);
                _ = process.Handle;
                Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
                Assert.Equal(0, NtSuspendProcess(process.Handle));
                initialized.SetResult();
            }
        });
        using var cancellation = new CancellationTokenSource();
        var path = Path.Join(Path.GetTempPath(), $"unconfirmed-{Guid.NewGuid():N}.xlsx");
        var opening = controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path }),
            CallerCancellation = cancellation.Token
        });
        var failure = await Record.ExceptionAsync(async () =>
        {
            await initialized.Task.WaitAsync(TimeSpan.FromSeconds(5));
            await Task.Delay(100);
            cancellation.Cancel();
            var response = await opening.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(response.Success);
            using var partial = JsonDocument.Parse(response.PartialState!);
            Assert.Equal("unknown", partial.RootElement.GetProperty("outcome").GetString());
            Assert.Empty(controller.Sessions);
            Assert.NotNull(process);
            process.Kill();
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            for (var index = 0; index < 25; index++)
            {
                var retry = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = path })
                });
                Assert.False(retry.Success);
                Assert.Equal("Conflict", retry.ErrorCategory);
                await Task.Delay(100);
            }
        });
        var cleanup = await Record.ExceptionAsync(async () =>
        {
            cancellation.Cancel();
            if (process is { HasExited: false })
            {
                process.Kill();
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
            }
            await opening.WaitAsync(TimeSpan.FromSeconds(5));
            process?.Dispose();
        });
        if (failure is not null && cleanup is not null)
            throw new AggregateException(failure, cleanup);
        if (failure is not null || cleanup is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture((failure ?? cleanup)!).Throw();
    }

    [DllImport("ntdll.dll")]
    private static extern int NtSuspendProcess(IntPtr processHandle);

    [Fact(Timeout = 30000)]
    public async Task Controller_ExitedWorkerAndExcelReleasePathsWithoutClaimingConfirmedClosure()
    {
        using var exited = new ProductProcessLauncher("dotnet", ["--list-sdks"]).Start([]);
        _ = exited.Handle;
        var identity = WorkerProcess.From(exited);
        await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        using var controller = new ExcelController(launcher: Launcher);
        var worker = await controller.StartWorkerAsync(null, CancellationToken.None);
        var session = Guid.NewGuid().ToString("N");
        var path = Path.Join(Path.GetTempPath(), $"exited-{Guid.NewGuid():N}.xlsx");
        worker.RecordSnapshot(new WorkerSnapshot(
            worker.Registration with { Excel = new ExcelProcessIdentity(identity.ProcessId, identity.StartedAtUtcFileTime) },
            [new SessionDescriptor(session, path)],
            new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(new
                {
                    sessions = new[] { new { sessionId = session, filePath = path, filePaths = new[] { path }, operationPending = false } }
                })
            }, null, null,
            new Dictionary<string, IReadOnlyList<string>> { [session] = [path] },
            Revision: long.MaxValue));
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        while (controller.Sessions.Count == 0)
            await Task.Delay(50, timeout.Token);
        var reserved = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path })
        });
        Assert.False(reserved.Success);
        Assert.Equal("Conflict", reserved.ErrorCategory);
        using var process = System.Diagnostics.Process.GetProcessById(worker.Registration.Worker.ProcessId);
        _ = process.Handle;
        Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
        Assert.True((await worker.ReleaseAsync(CancellationToken.None)).Success);
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        while (controller.Sessions.Count != 0)
            await Task.Delay(50, timeout.Token);
        var closed = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = session });
        Assert.False(closed.Success);
        Assert.Equal("ServiceUnavailable", closed.ErrorCategory);
        using var partial = JsonDocument.Parse(closed.PartialState!);
        Assert.Equal("unknown", partial.RootElement.GetProperty("outcome").GetString());
        var reopened = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 1 })
        });
        Assert.False(reopened.Success);
        Assert.Equal("ArgumentOutOfRangeException", reopened.ExceptionType);
    }

    [Theory(Timeout = 30000)]
    [InlineData("session.save", null, "InvalidInput")]
    [InlineData("session.save", "missing", "SessionNotFound")]
    [InlineData("service.restart", null, "InvalidInput")]
    [InlineData("unknown.action", null, "InvalidInput")]
    public async Task Controller_PreservesExistingValidation(string command, string? session, string expectedCategory)
    {
        var request = new ServiceRequest { Command = command, SessionId = session };
        using var expectedService = new ExcelMcpService();
        var expected = await expectedService.ProcessAsync(request);
        using var controller = new ExcelController(launcher: Launcher);
        var actual = await controller.ProcessAsync(request);
        Assert.False(actual.Success);
        Assert.Equal(expectedCategory, actual.ErrorCategory);
        Assert.Equal(expected.ErrorCategory, actual.ErrorCategory);
        Assert.Equal(expected.ErrorMessage, actual.ErrorMessage);
        Assert.Equal(expected.Command, actual.Command);
        Assert.Equal(expected.SessionId, actual.SessionId);
        Assert.Empty(controller.ExcelProcesses);
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_InvalidCloseDoesNotBreakLaterSessionlessCommands()
    {
        using var controller = new ExcelController(launcher: Launcher);
        var closed = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.close",
            SessionId = "missing"
        });
        Assert.False(closed.Success);
        Assert.Equal("SessionNotFound", closed.ErrorCategory);
        var echoed = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"after-invalid-close"}"""
        });
        Assert.True(echoed.Success, echoed.ErrorMessage);
        using var result = JsonDocument.Parse(echoed.Result!);
        Assert.Equal("after-invalid-close", result.RootElement.GetProperty("message").GetString());
    }

    [Fact(Timeout = 30000)]
    public async Task Controller_FailedInitializationReleasesItsLateEmptyWorker()
    {
        ExcelWorkerClient? launched = null;
        using var controller = new ExcelController(launcher: Launcher, initializeWorker: worker =>
        {
            launched = worker;
            throw new TimeoutException("The startup reply was not confirmed.");
        });
        var response = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "diag.echo",
            Args = """{"message":"startup-failure"}"""
        });
        Assert.False(response.Success);
        Assert.Equal("Timeout", response.ErrorCategory);
        Assert.NotNull(launched);
        using var process = System.Diagnostics.Process.GetProcessById(launched.Registration.Worker.ProcessId);
        _ = process.Handle;
        Assert.Equal(launched.Registration.Worker, WorkerProcess.From(process));
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
        Assert.Empty(controller.Sessions);
        Assert.Empty(controller.ExcelProcesses);
    }

    [Fact(Timeout = 60000)]
    [Trait("Speed", "Slow")]
    public async Task Controller_ExitedStartupWorkersDoNotExhaustCapacity()
    {
        var launcher = new ProductProcessLauncher("dotnet", ["--list-sdks"]);
        using var controller = new ExcelController(launcher: launcher);
        for (var index = 0; index < 33; index++)
        {
            var response = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "diag.echo",
                Args = """{"message":"exited-startup"}"""
            });
            Assert.False(response.Success);
            Assert.DoesNotContain("worker limit", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exited", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact(Timeout = 15000)]
    public async Task Controller_ReleasesEmptyAttachedWorkerForConfirmedExitedProcess()
    {
        using var exited = new ProductProcessLauncher("dotnet", ["--list-sdks"]).Start([]);
        _ = exited.Handle;
        var identity = WorkerProcess.From(exited);
        await exited.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5));
        using var controller = new ExcelController(launcher: Launcher);
        var worker = await controller.StartWorkerAsync(
            new ExcelProcessIdentity(identity.ProcessId, identity.StartedAtUtcFileTime), CancellationToken.None);
        Assert.False(worker.Registration.OwnsExcel);
        Assert.Empty(worker.Snapshot!.Sessions);
        using var process = System.Diagnostics.Process.GetProcessById(worker.Registration.Worker.ProcessId);
        _ = process.Handle;
        Assert.Equal(worker.Registration.Worker, WorkerProcess.From(process));
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(8));
    }

    [Fact(Timeout = 15000)]
    public async Task ControlEndpoint_TracksAcceptedControlsButNotRejectedWorkbookWork()
    {
        var tracked = new System.Collections.Concurrent.ConcurrentBag<string>();
        var pipeName = $"excelmcp-controller-test-{Guid.NewGuid():N}";
        using var controller = new ExcelController(trackRequest: (request, next) =>
        {
            tracked.Add(request.Command);
            return next();
        }, launcher: Launcher);
        var run = controller.RunAsync(pipeName);
        try
        {
            using var control = new ServiceClient(pipeName, useControlEndpoint: true);
            string[] commands = ["service.status", "session.list"];
            foreach (var command in commands)
                Assert.True((await control.SendAsync(new ServiceRequest { Command = command }, TimeSpan.FromSeconds(2))).Success);
            await using var pipe = ServiceSecurity.CreateClient(ServiceClient.GetControlPipeName(pipeName));
            await pipe.ConnectAsync(2000);
            var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
            using var connection = (IDisposable)proxy;
            var rejected = await proxy.ProcessCommandAsync(new ServiceRequest { Command = "diag.echo" });
            Assert.False(rejected.Success);
            Assert.Equal("InvalidInput", rejected.ErrorCategory);
            Assert.Equal(commands.Order(), tracked.Order());
        }
        finally
        {
            await controller.RequestShutdownAsync();
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
    }

    [Theory(Timeout = 30000)]
    [InlineData("session.open")]
    [InlineData("session.create")]
    [InlineData("session.test")]
    [InlineData("sheet.copy-to-file")]
    [InlineData("sheet.move-to-file")]
    public async Task Controller_CancelledStartupReleasesItsEmptyWorkerAndAllowsRetry(string command)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-missing-{Guid.NewGuid():N}.xlsx");
        var registry = new WorkerRegistry();
        var existing = registry.Read().Select(record => record.PipeName).ToHashSet();
        using var controller = new ExcelController(launcher: Launcher);
        using var cancellation = new CancellationTokenSource();
        var opening = controller.ProcessAsync(new ServiceRequest
        {
            Command = command,
            Args = command.StartsWith("sheet.", StringComparison.Ordinal)
                ? JsonSerializer.Serialize(new { sourceFile = path, targetFile = path + ".target.xlsx", sourceSheet = "Sheet1" })
                : JsonSerializer.Serialize(new { filePath = path }),
            CallerCancellation = cancellation.Token
        });
        Assert.False(opening.IsCompleted);
        cancellation.Cancel();
        var cancelled = await opening;
        Assert.False(cancelled.Success);
        Assert.Equal("Cancelled", cancelled.ErrorCategory);
        Assert.Empty(controller.Sessions);
        Assert.False(File.Exists(path));
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(8));
        while (registry.Read().Any(record => record.Controller == WorkerProcess.Current
            && !existing.Contains(record.PipeName) && record.Worker.IsAlive()) && !deadline.IsCancellationRequested)
            await Task.Delay(50);
        Assert.DoesNotContain(registry.Read(), record =>
            record.Controller == WorkerProcess.Current && !existing.Contains(record.PipeName) && record.Worker.IsAlive());
        var retried = await controller.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 1 })
        });
        Assert.False(retried.Success);
        Assert.Equal("ArgumentOutOfRangeException", retried.ExceptionType);
        Assert.Empty(controller.Sessions);
        Assert.False(File.Exists(path));
    }

    [Fact(Timeout = 30000)]
    public async Task ControlEndpoint_AnswersWhenAllWorkConnectionsAreOccupied()
    {
        var pipeName = $"excelmcp-controller-test-{Guid.NewGuid():N}";
        using var controller = new ExcelController(launcher: Launcher);
        var run = controller.RunAsync(pipeName);
        var connections = new List<IDisposable>();
        try
        {
            for (var index = 0; index < DaemonHost.MaximumConnections; index++)
            {
                var pipe = ServiceSecurity.CreateClient(pipeName);
                connections.Add(pipe);
                await pipe.ConnectAsync(2000);
                var proxy = JsonRpc.Attach<IExcelDaemonRpc>(pipe);
                connections.Add((IDisposable)proxy);
                Assert.True((await proxy.ProcessCommandAsync(
                    new ServiceRequest { Command = "service.ping" }).WaitAsync(TimeSpan.FromSeconds(2))).Success);
            }
            using var control = new ServiceClient(pipeName, useControlEndpoint: true);
            foreach (var command in new[] { "service.ping", "service.status", "session.list" })
            {
                var response = await control.SendAsync(new ServiceRequest { Command = command }, TimeSpan.FromSeconds(2));
                Assert.True(response.Success, response.ErrorMessage);
                Assert.Null(response.Command);
                if (command == "service.status")
                {
                    using var status = JsonDocument.Parse(response.Result!);
                    Assert.True(status.RootElement.GetProperty("running").GetBoolean());
                    Assert.Equal(Environment.ProcessId, status.RootElement.GetProperty("processId").GetInt32());
                    Assert.Equal(0, status.RootElement.GetProperty("sessionCount").GetInt32());
                }
                else if (command == "session.list")
                {
                    using var list = JsonDocument.Parse(response.Result!);
                    Assert.Equal(0, list.RootElement.GetProperty("count").GetInt32());
                    Assert.Empty(list.RootElement.GetProperty("sessions").EnumerateArray());
                }
            }
            var stopped = await control.SendAsync(new ServiceRequest { Command = "service.shutdown" }, TimeSpan.FromSeconds(2));
            Assert.True(stopped.Success, stopped.ErrorMessage);
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
        finally
        {
            foreach (var connection in connections)
                connection.Dispose();
            await controller.RequestShutdownAsync();
            await run.WaitAsync(TimeSpan.FromSeconds(4));
        }
    }
}

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "SessionLifecycle")]
[Trait("RequiresExcel", "true")]
public sealed class ExcelControllerWorkbookTests
{
    private static readonly string[] ApprovalActions = ["enable-editing", "enable-connections", "enable-macros"];

    [Fact(Timeout = 60000)]
    public async Task Controller_ReleasesItsWorkerAfterExcelClosesOutsideController()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-external-close-{Guid.NewGuid():N}.xlsx");
        var launcher = new ProductProcessLauncher(Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);
        using var controller = new ExcelController(launcher: launcher);
        System.Diagnostics.Process? worker = null;
        System.Diagnostics.Process? excel = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var created = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path, show = true })
            });
            Assert.True(created.Success, created.ErrorMessage);
            var identity = Assert.Single(controller.ExcelProcesses);
            excel = System.Diagnostics.Process.GetProcessById(identity.ProcessId);
            _ = excel.Handle;
            Assert.Equal(identity.StartedAtUtcFileTime, excel.StartTime.ToUniversalTime().ToFileTimeUtc());
            var registration = Assert.Single(new WorkerRegistry().Read(),
                record => record.Controller == WorkerProcess.Current && record.Excel == identity && record.Worker.IsAlive());
            worker = System.Diagnostics.Process.GetProcessById(registration.Worker.ProcessId);
            _ = worker.Handle;
            Assert.Equal(registration.Worker, WorkerProcess.From(worker));
            var client = new ExcelWorkerClient(registration, WorkerProcess.Current, registration.Generation);
            var closed = await client.ExecuteAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = created.SessionId
            }, CancellationToken.None);
            Assert.True(closed.Success, closed.ErrorMessage);
            await excel.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(8));
            Assert.True(SpinWait.SpinUntil(
                () => controller.Sessions.Count == 0 && controller.ExcelProcesses.Count == 0,
                TimeSpan.FromSeconds(5)), "The controller did not read the completed close from its worker.");
            Assert.Empty(controller.Sessions);
            Assert.Empty(controller.ExcelProcesses);
            var again = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = created.SessionId
            });
            Assert.True(again.Success, again.ErrorMessage);
            Assert.Contains("already closed", again.Result);
        });
        var cleanupFailure = await Record.ExceptionAsync(async () =>
        {
            await controller.RequestShutdownAsync();
            controller.Dispose();
            if (worker is not null)
                await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            if (excel is not null)
                await excel.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            File.Delete(path);
        });
        worker?.Dispose();
        excel?.Dispose();
        var failures = new[] { failure, cleanupFailure }.OfType<Exception>().ToArray();
        if (failures.Length > 0) throw new AggregateException("External Excel closure or cleanup failed.", failures);
    }

    [Theory(Timeout = 90000)]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Controller_ExitedWorkerAllowsExplicitReattachmentWithoutLosingUnsavedValues(bool pendingOpen)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-worker-recovery-{Guid.NewGuid():N}.xlsx");
        var launcher = new ProductProcessLauncher(Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);
        var startupEntered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseStartup = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var startupCancellation = new CancellationTokenSource();
        var blockStartup = false;
        Task<ServiceResponse>? opening = null;
        async Task InitializeWorkerAsync(ExcelWorkerClient worker)
        {
            await worker.InitializeAsync();
            if (blockStartup)
            {
                startupEntered.TrySetResult();
                await releaseStartup.Task.WaitAsync(startupCancellation.Token);
            }
        }
        using var controller = new ExcelController(launcher: launcher, initializeWorker: InitializeWorkerAsync);
        using var recoveryController = new ExcelController(launcher: launcher, initializeWorker: InitializeWorkerAsync);
        System.Diagnostics.Process? excelProcess = null;
        ExcelProcessIdentity? identity = null;
        string? session = null;
        var workerExited = false;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var created = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path, show = true })
            });
            Assert.True(created.Success, created.ErrorMessage);
            session = Assert.IsType<string>(created.SessionId);
            var currentIdentity = Assert.Single(controller.ExcelProcesses);
            identity = currentIdentity;
            excelProcess = System.Diagnostics.Process.GetProcessById(currentIdentity.ProcessId);
            _ = excelProcess.Handle;
            for (var index = 0; index < 2; index++)
            {
                var marker = index == 0 ? "saved-before-worker-exit" : "unsaved-after-worker-exit";
                var written = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "range.set-values",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new
                    {
                        sheetName = "Sheet1",
                        rangeAddress = "A1",
                        values = new[] { new[] { marker } },
                        overwritePolicy = "allow"
                    })
                });
                Assert.True(written.Success, written.ErrorMessage);
                if (marker == "saved-before-worker-exit")
                    Assert.True((await controller.ProcessAsync(new ServiceRequest { Command = "session.save", SessionId = session })).Success);
            }
            var registration = Assert.Single(new WorkerRegistry().Read(),
                record => record.Controller == WorkerProcess.Current && record.Excel == identity && record.Worker.IsAlive());
            using (var worker = System.Diagnostics.Process.GetProcessById(registration.Worker.ProcessId))
            {
                _ = worker.Handle;
                Assert.Equal(registration.Worker.StartedAtUtcFileTime, worker.StartTime.ToUniversalTime().ToFileTimeUtc());
                var client = new ExcelWorkerClient(registration, WorkerProcess.Current, registration.Generation);
                var released = await client.ReleaseAsync(CancellationToken.None);
                Assert.True(released.Success, released.ErrorMessage);
                await worker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                workerExited = true;
            }
            Assert.True(OwnedProcessGuard.MatchesLiveProcess(currentIdentity));
            var registry = new WorkerRegistry();
            registry.Write(registration);
            var earlierControllerRecord = registration with
            {
                PipeName = $"excelmcp-worker-{Guid.NewGuid():N}",
                Controller = registration.Worker,
                Generation = Guid.NewGuid().ToString("N")
            };
            registry.Write(earlierControllerRecord);
            var discovered = await controller.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            Assert.True(discovered.Success, discovered.ErrorMessage);
            using var discovery = JsonDocument.Parse(discovered.Result!);
            Assert.False(discovery.RootElement.GetProperty("complete").GetBoolean());
            Assert.Contains(discovery.RootElement.GetProperty("errors").EnumerateArray(),
                error => error.GetProperty("processId").GetInt32() == currentIdentity.ProcessId
                    && error.GetProperty("message").GetString()!.Contains("no edit was repeated", StringComparison.Ordinal));
            Assert.DoesNotContain(registry.Read(),
                record => record.PipeName == registration.PipeName || record.PipeName == earlierControllerRecord.PipeName);
            var replacement = Assert.Single(registry.Read(), record =>
                record.Excel == currentIdentity && record.Worker.IsAlive());
            using (var replacementProcess = System.Diagnostics.Process.GetProcessById(replacement.Worker.ProcessId))
            {
                _ = replacementProcess.Handle;
                Assert.Equal(replacement.Worker.StartedAtUtcFileTime,
                    replacementProcess.StartTime.ToUniversalTime().ToFileTimeUtc());
                await controller.RequestShutdownAsync();
                await replacementProcess.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            }
            discovered = await recoveryController.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            Assert.True(discovered.Success, discovered.ErrorMessage);
            using var rediscovery = JsonDocument.Parse(discovered.Result!);
            Assert.DoesNotContain(rediscovery.RootElement.GetProperty("errors").EnumerateArray(),
                error => error.GetProperty("message").GetString()!.Contains("no edit was repeated", StringComparison.Ordinal));
            var selected = Assert.Single(rediscovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                item => item.GetProperty("location").GetString() == path);
            Assert.False(selected.GetProperty("saved").GetBoolean());
            var attachRequest = new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new { workbookId = selected.GetProperty("workbookId").GetString() })
            };
            if (pendingOpen)
            {
                blockStartup = true;
                opening = recoveryController.ProcessAsync(new ServiceRequest
                {
                    Command = "session.open",
                    Args = JsonSerializer.Serialize(new { filePath = path }),
                    CallerCancellation = startupCancellation.Token
                });
                await startupEntered.Task.WaitAsync(TimeSpan.FromSeconds(5));
                var refused = await recoveryController.ProcessAsync(attachRequest);
                Assert.False(refused.Success);
                Assert.Equal("Conflict", refused.ErrorCategory);
                blockStartup = false;
                startupCancellation.Cancel();
                releaseStartup.TrySetResult();
                var cancelled = await opening.WaitAsync(TimeSpan.FromSeconds(10));
                Assert.False(cancelled.Success);
                Assert.Equal("Cancelled", cancelled.ErrorCategory);
            }
            var attached = await recoveryController.ProcessAsync(attachRequest);
            Assert.True(attached.Success, attached.ErrorMessage);
            Assert.NotEqual(session, attached.SessionId);
            Assert.Equal(attached.SessionId, Assert.Single(recoveryController.Sessions).SessionId);
            var again = await recoveryController.ProcessAsync(attachRequest);
            Assert.True(again.Success, again.ErrorMessage);
            Assert.Equal(attached.SessionId, again.SessionId);
            Assert.Equal(attached.SessionId, Assert.Single(recoveryController.Sessions).SessionId);
            var read = await recoveryController.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = attached.SessionId,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.True(read.Success, read.ErrorMessage);
            using var result = JsonDocument.Parse(read.Result!);
            Assert.Equal("unsaved-after-worker-exit", result.RootElement.GetProperty("values")[0][0].GetString());
        });
        var cleanupFailure = await Record.ExceptionAsync(async () =>
        {
            blockStartup = false;
            startupCancellation.Cancel();
            releaseStartup.TrySetResult();
            if (opening is not null)
                await opening.WaitAsync(TimeSpan.FromSeconds(10));
            var remaining = new WorkerRegistry().Read().SingleOrDefault(record =>
                record.Controller == WorkerProcess.Current && record.Excel == identity && record.Worker.IsAlive());
            using var remainingWorker = remaining is null ? null
                : System.Diagnostics.Process.GetProcessById(remaining.Worker.ProcessId);
            if (remainingWorker is not null)
            {
                _ = remainingWorker.Handle;
                Assert.Equal(remaining!.Worker.StartedAtUtcFileTime, remainingWorker.StartTime.ToUniversalTime().ToFileTimeUtc());
            }
            if (!workerExited && session is not null)
                Assert.True((await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = session })).Success);
            await controller.RequestShutdownAsync();
            controller.Dispose();
            await recoveryController.RequestShutdownAsync();
            recoveryController.Dispose();
            if (remainingWorker is not null)
                await remainingWorker.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            if (workerExited && identity is { } cleanupIdentity)
                await CloseTestWorkbookAsync(cleanupIdentity, path);
            if (excelProcess is not null)
                await excelProcess.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(15));
            using var verification = new ExcelMcpService();
            var opened = await verification.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            var read = await verification.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = opened.SessionId,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.True(read.Success, read.ErrorMessage);
            using var result = JsonDocument.Parse(read.Result!);
            Assert.Equal("saved-before-worker-exit", result.RootElement.GetProperty("values")[0][0].GetString());
            Assert.True((await verification.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = opened.SessionId
            })).Success);
        });
        excelProcess?.Dispose();
        File.Delete(path);
        var failures = new[] { failure, cleanupFailure }.OfType<Exception>().ToArray();
        if (failures.Length > 0) throw new AggregateException("Worker recovery or test cleanup failed.", failures);
    }

    private static Task CloseTestWorkbookAsync(ExcelProcessIdentity identity, string path)
    {
        var completed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var cleanup = new Thread(() =>
        {
            try
            {
                CloseTestWorkbook(identity, path);
                completed.SetResult();
            }
            catch (Exception ex) { completed.SetException(ex); }
        })
        { IsBackground = true };
        cleanup.SetApartmentState(ApartmentState.STA);
        cleanup.Start();
        return completed.Task.WaitAsync(TimeSpan.FromSeconds(15));
    }

    private static void CloseTestWorkbook(ExcelProcessIdentity identity, string path)
    {
        Excel.Application? application = null;
        Excel.Workbooks? books = null;
        Excel.Workbook? workbook = null;
        try
        {
            application = RunningExcelWindows.Connect(RunningExcelWindows.Find().Windows
                .Where(window => window.Process == identity).DistinctBy(window => window.Process).Single());
            books = application.Workbooks;
            Assert.Equal(1, books.Count);
            workbook = books.Item[Path.GetFileName(path)];
            Assert.Equal(path, workbook.FullName, ignoreCase: true);
            ComUtilities.Release(ref books);
            Assert.True(ExcelShutdownService.CloseAndQuitAndReport(workbook, application, save: false, releaseWorkbook: false));
            application = null;
        }
        finally
        {
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref books);
            ComUtilities.Release(ref application);
        }
    }

    [Fact(Timeout = 90000)]
    public async Task Controller_WorkerPreservesValuesAndReleasesClosedSessions()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-controller-{Guid.NewGuid():N}.xlsx");
        var savedAsPath = Path.ChangeExtension(path, ".saved-as.xlsx");
        var launcher = new ProductProcessLauncher(Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, @"..\..\..\..\..\src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe")), []);
        using var controller = new ExcelController(launcher: launcher);
        string? sessionId = null;
        string? originalSessionId = null;
        var failure = await Record.ExceptionAsync(async () =>
        {
            var created = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path, show = true })
            });
            Assert.True(created.Success, created.ErrorMessage);
            sessionId = Assert.IsType<string>(created.SessionId);
            Assert.Single(controller.Sessions);
            Assert.Single(controller.ExcelProcesses);
            var sheet = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.create",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker"}"""
            });
            Assert.True(sheet.Success, sheet.ErrorMessage);
            var written = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1","values":[["worker-owned-marker"]]}"""
            });
            Assert.True(written.Success, written.ErrorMessage);
            var inspected = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.notices",
                SessionId = sessionId
            });
            Assert.True(inspected.Success, inspected.ErrorMessage);
            using var inspection = JsonDocument.Parse(inspected.Result!);
            var notices = inspection.RootElement.GetProperty("notices");
            Assert.False(notices.GetProperty("protectedView").GetBoolean());
            Assert.False(notices.GetProperty("readOnly").GetBoolean());
            Assert.False(notices.GetProperty("hasVbaProject").GetBoolean());
            Assert.Equal(0, notices.GetProperty("connectionCount").GetInt32());
            Assert.False(notices.GetProperty("complete").GetBoolean());
            foreach (var action in ApprovalActions)
            {
                var refused = await controller.ProcessAsync(new ServiceRequest
                {
                    Command = $"session.{action}",
                    SessionId = sessionId
                });
                Assert.False(refused.Success);
                Assert.Contains("approval", refused.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            var read = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1"}"""
            });
            Assert.True(read.Success, read.ErrorMessage);
            using var result = JsonDocument.Parse(read.Result!);
            Assert.Equal("worker-owned-marker",
                result.RootElement.GetProperty("values")[0][0].GetString());
            var discovered = await controller.ProcessAsync(new ServiceRequest { Command = "session.discover" });
            Assert.True(discovered.Success, discovered.ErrorMessage);
            using var discovery = JsonDocument.Parse(discovered.Result!);
            var selected = Assert.Single(discovery.RootElement.GetProperty("workbooks").EnumerateArray(),
                item => item.GetProperty("location").GetString() == path);
            var attachRequest = new ServiceRequest
            {
                Command = "session.attach",
                Args = JsonSerializer.Serialize(new { workbookId = selected.GetProperty("workbookId").GetString() })
            };
            var attached = await controller.ProcessAsync(attachRequest);
            Assert.True(attached.Success, attached.ErrorMessage);
            Assert.Equal(sessionId, attached.SessionId);
            var savedAs = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "workbook.save-as",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { targetPath = savedAsPath })
            });
            Assert.True(savedAs.Success, savedAs.ErrorMessage);
            Assert.Equal(savedAsPath, Assert.Single(controller.Sessions).FilePath);
            var conflicting = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = savedAsPath })
            });
            Assert.False(conflicting.Success);
            Assert.Equal("Conflict", conflicting.ErrorCategory);
            var original = await controller.ProcessAsync(new ServiceRequest
            {
                Command = File.Exists(path) ? "session.open" : "session.create",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(original.Success, original.ErrorMessage);
            originalSessionId = Assert.IsType<string>(original.SessionId);
            var originalWritten = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.set-values",
                SessionId = original.SessionId,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1","values":[["original-workbook-marker"]]}"""
            });
            Assert.True(originalWritten.Success, originalWritten.ErrorMessage);
            var reattached = await controller.ProcessAsync(attachRequest);
            Assert.True(reattached.Success, reattached.ErrorMessage);
            Assert.Equal(sessionId, reattached.SessionId);
            Assert.Equal(2, controller.Sessions.Count);
            var originalRead = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = original.SessionId,
                Args = """{"sheetName":"Sheet1","rangeAddress":"A1"}"""
            });
            Assert.True(originalRead.Success, originalRead.ErrorMessage);
            using var originalValues = JsonDocument.Parse(originalRead.Result!);
            Assert.Equal("original-workbook-marker", originalValues.RootElement.GetProperty("values")[0][0].GetString());
            var originalClosed = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "session.close",
                SessionId = original.SessionId
            });
            Assert.True(originalClosed.Success, originalClosed.ErrorMessage);
            originalSessionId = null;
            var afterSaveAs = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "range.get-values",
                SessionId = sessionId,
                Args = """{"sheetName":"WorkerMarker","rangeAddress":"A1"}"""
            });
            Assert.True(afterSaveAs.Success, afterSaveAs.ErrorMessage);
            using var afterSaveAsValues = JsonDocument.Parse(afterSaveAs.Result!);
            Assert.Equal("worker-owned-marker", afterSaveAsValues.RootElement.GetProperty("values")[0][0].GetString());
            var traces = await controller.ProcessAsync(new ServiceRequest
            {
                Command = "powerquerytrace.list",
                Args = JsonSerializer.Serialize(new { workbookSessionId = sessionId })
            });
            Assert.True(traces.Success, traces.ErrorMessage);
            using var traceResult = JsonDocument.Parse(traces.Result!);
            Assert.Equal(Assert.Single(controller.ExcelProcesses).ProcessId,
                traceResult.RootElement.GetProperty("processId").GetInt32());
            var closed = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
            Assert.True(closed.Success, closed.ErrorMessage);
            using var closeResult = JsonDocument.Parse(closed.Result!);
            var closeTimings = closeResult.RootElement.GetProperty("closeTimingsMs");
            Assert.Equal(10, closeTimings.EnumerateObject().Count());
            Assert.False(closeTimings.TryGetProperty("save", out _));
            Assert.True(closeTimings.GetProperty("session-close").GetDouble() > 0);
            Assert.True(closeTimings.GetProperty("worker-completion").GetDouble() >= 0);
            Assert.Empty(controller.ExcelProcesses);
            Assert.Empty(controller.Sessions);
            var again = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
            Assert.True(again.Success, again.ErrorMessage);
            Assert.Contains("already closed", again.Result);
            sessionId = null;
            Assert.Null(await controller.RequestShutdownAsync());
        });
        var cleanupFailure = await Record.ExceptionAsync(async () =>
        {
            if (originalSessionId is not null)
                Assert.True((await controller.ProcessAsync(new ServiceRequest
                {
                    Command = "session.close",
                    SessionId = originalSessionId
                })).Success);
            if (sessionId is not null)
            {
                var closed = await controller.ProcessAsync(new ServiceRequest { Command = "session.close", SessionId = sessionId });
                Assert.True(closed.Success, closed.ErrorMessage);
            }
            File.Delete(path);
            File.Delete(savedAsPath);
        });
        var failures = new[] { failure, cleanupFailure }.OfType<Exception>().ToArray();
        if (failures.Length > 0) throw new AggregateException("Workbook state or cleanup failed.", failures);
    }
}
