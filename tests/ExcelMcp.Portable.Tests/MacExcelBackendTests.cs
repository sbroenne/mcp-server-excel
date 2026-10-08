using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac backend state")]
[Trait("RequiresExcel", "false")]
public sealed class MacExcelBackendTests
{
    private static string TestPath(string name) => Path.Combine(Path.GetTempPath(), name);

    [Fact]
    public async Task SharedService_DaemonIdleMonitorCountsMacSessions()
    {
        var backend = new MacExcelBackend((_, _, _) =>
            Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "")));
        var path = TestPath($"idle-session-{Guid.NewGuid():N}.xlsx");
        MacWorkbookTemplate.Copy(path, macroEnabled: false);
        try
        {
            using var service = new ExcelMcpService(backend);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(response.Success, response.ErrorMessage);
            Assert.Equal(1, service.SessionCount);

            var daemon = typeof(ExcelMcpService)
                .GetField("_daemonHost", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(service)!;
            var sessionCount = Assert.IsType<Func<int>>(daemon.GetType()
                .GetField("_sessionCount", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(daemon));
            Assert.Equal(service.SessionCount, sessionCount());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task SharedService_DiagnosticCommandDoesNotDispatchToExcel()
    {
        var backend = new MacExcelBackend((_, _, _) =>
            throw new InvalidOperationException("Diagnostic commands must not reach Excel."));
        using var service = new ExcelMcpService(backend);

        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "diag.ping"
        });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Null(response.ErrorMessage);
        Assert.NotNull(response.Result);
    }

    [Fact]
    public async Task NativeTimeout_RequiresRecoveryWithoutClosingOrDiscardingWorkbook()
    {
        var closeCalls = 0;
        var backend = new MacExcelBackend(async (start, _, cancellationToken) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
            }

            switch (AutomationCommand(start))
            {
                case "sheet.list":
                    return new MacProcessResult(0,
                        """{"success":true,"errorMessage":"","worksheets":[{"name":"Sheet1","index":1,"visible":true}]}""", "");
                case "sheet.rename":
                    await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
                    break;
                case "session.close":
                    closeCalls++;
                    break;
            }
            return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
        });
        var directory = Directory.CreateTempSubdirectory("excelmcp-timeout-recovery-");
        var path = Path.Combine(directory.FullName, "timeout.xlsx");
        MacWorkbookTemplate.Copy(path, macroEnabled: false);
        try
        {
            using var service = new ExcelMcpService(backend);
            var opened = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 10 })
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            using var openedJson = JsonDocument.Parse(opened.Result!);
            var sessionId = openedJson.RootElement.GetProperty("sessionId").GetString()!;

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "sheet.rename",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new { oldName = "Sheet1", newName = "Data" })
            });

            Assert.False(response.Success);
            Assert.True(response.ErrorCategory == "Timeout", response.ErrorMessage);
            Assert.Equal(0, closeCalls);
            var listed = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            using var listJson = JsonDocument.Parse(listed.Result!);
            var recovery = Assert.Single(listJson.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.True(recovery.GetProperty("requiresRecovery").GetBoolean());
            Assert.False(recovery.GetProperty("canClose").GetBoolean());
        }
        finally
        {
            File.Delete(path);
            directory.Delete();
        }
    }

    [Fact]
    public async Task ScenarioEnvironmentVariable_DoesNotBypassProductionCapabilityGate()
    {
        var dispatched = new List<string>();
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.FileName != "/usr/bin/open")
            {
                dispatched.Add(AutomationCommand(start));
            }
            return Task.FromResult(new MacProcessResult(
                0,
                """{"success":true,"errorMessage":""}""",
                ""));
        });
        var directory = Directory.CreateTempSubdirectory("excelmcp-scenario-gate-");
        var path = Path.Combine(directory.FullName, "scenario.xlsx");
        MacWorkbookTemplate.Copy(path, macroEnabled: false);
        var previous = Environment.GetEnvironmentVariable("EXCELMCP_MAC_SCENARIO_E2E");
        try
        {
            Environment.SetEnvironmentVariable("EXCELMCP_MAC_SCENARIO_E2E", "1");
            using var service = new ExcelMcpService(backend);
            var opened = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            using var openedJson = JsonDocument.Parse(opened.Result!);

            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "analysis.list-scenarios",
                SessionId = openedJson.RootElement.GetProperty("sessionId").GetString(),
                Args = JsonSerializer.Serialize(new { sheetName = "Data" })
            });

            Assert.False(response.Success);
            Assert.True(response.ErrorCategory == "PlatformNotSupported", response.ErrorMessage);
            Assert.DoesNotContain("analysis.list-scenarios", dispatched);
        }
        finally
        {
            Environment.SetEnvironmentVariable("EXCELMCP_MAC_SCENARIO_E2E", previous);
            File.Delete(path);
            directory.Delete();
        }
    }

    [Fact]
    public async Task StructuredCommandFailure_CanRemainAResultDto()
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(
                0,
                JsonSerializer.Serialize(new
                {
                    success = false,
                    filePath = TestPath("test.xlsx"),
                    isPythonError = true,
                    errorMessage = "#PYTHON! - Python code raised an error"
                }),
                "")));

        var result = await backend.InvokeAsync(
            "pythoninexcel.get-result",
            new { filePath = TestPath("test.xlsx") },
            TimeSpan.FromSeconds(5),
            allowFailureResult: true);

        Assert.False(result.GetProperty("success").GetBoolean());
        Assert.True(result.GetProperty("isPythonError").GetBoolean());
    }

    [Fact]
    public async Task AutomationFailure_IsNotMistakenForStructuredCommandFailure()
    {
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
            Task.FromResult(new MacProcessResult(
                0,
                """{"success":false,"errorCategory":"ComInterop","errorMessage":"Worksheet does not exist."}""",
                "")));

        await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "pythoninexcel.get-result",
            new { filePath = TestPath("test.xlsx") },
            TimeSpan.FromSeconds(5),
            allowFailureResult: true));
    }

    [Fact]
    public async Task AutomationChild_ReceivesTheRemainingOperationBudget()
    {
        var budgets = new List<TimeSpan>();
        var backend = new MacExcelBackend(async (start, _, cancellationToken) =>
        {
            if (start.FileName != "/usr/bin/open")
            {
                var marker = start.ArgumentList.IndexOf(MacAutomationHost.Marker);
                Assert.True(marker >= 0);
                Assert.Equal(marker + 4, start.ArgumentList.Count);
                Assert.Equal(Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    start.ArgumentList[marker + 2]);
                budgets.Add(TimeSpan.FromTicks(long.Parse(start.ArgumentList[marker + 3],
                    System.Globalization.CultureInfo.InvariantCulture)));
            }
            await Task.Delay(20, cancellationToken);
            return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
        });

        await backend.InvokeAsync("session.open", new { filePath = TestPath("budget.xlsx") },
            TimeSpan.FromSeconds(5));

        Assert.Equal(2, budgets.Count);
        Assert.InRange(budgets[0], TimeSpan.FromTicks(1), TimeSpan.FromSeconds(5));
        Assert.InRange(budgets[1], TimeSpan.FromTicks(1), budgets[0] - TimeSpan.FromMilliseconds(20));
    }

    [Fact]
    public async Task NativeEventTimeout_RemainsATimeoutWithUncertainOutcome()
    {
        var backend = new MacExcelBackend((_, _, _) => Task.FromResult(new MacProcessResult(
            0, """{"success":false,"errorCategory":"Timeout","errorMessage":"Excel Apple Event timed out; its outcome may be uncertain."}""", "")));

        var error = await Assert.ThrowsAsync<TimeoutException>(() => backend.InvokeAsync(
            "session.is-open", new { filePath = TestPath("native-timeout.xlsx") }, TimeSpan.FromSeconds(5)));
        Assert.Contains("uncertain", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task NativeAttachmentTimeout_RetainsTheUncertainOpenOutcome()
    {
        var backend = new MacExcelBackend((start, _, _) => Task.FromResult(new MacProcessResult(
            0, start.FileName != "/usr/bin/open" && AutomationCommand(start) == "session.open"
                ? """{"success":false,"errorCategory":"Timeout","errorMessage":"Native attachment timed out."}"""
                : """{"success":true,"errorMessage":""}""", "")));

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("native-attachment-timeout.xlsx") }, TimeSpan.FromSeconds(5)));
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.Contains("uncertain", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Open_HandsOffExactPathBeforeAttachingWorkbook()
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
        });
        var path = TestPath("workbook with 'quotes' and spaces.xlsx");

        await backend.InvokeAsync("session.open", new { filePath = path, show = false }, TimeSpan.FromSeconds(5));

        Assert.Equal(3, calls.Count);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Equal("/usr/bin/open", calls[1].FileName);
        Assert.Equal(["-g", "-b", "com.microsoft.Excel", path], calls[1].ArgumentList.ToArray());
        Assert.Equal("session.open", AutomationCommand(calls[2]));
    }

    [Theory]
    [InlineData(-1743)]
    [InlineData(-1744)]
    public async Task PermissionFailure_DoesNotLaunchAnyProcess(int status)
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0,
                $$"""{"success":false,"errorMessage":"Automation unavailable (OSStatus {{status}}).","errorCategory":"ComInterop"}""",
                ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("test.xlsx") }, TimeSpan.FromSeconds(5)));

        Assert.Single(calls);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Contains(status.ToString(System.Globalization.CultureInfo.InvariantCulture), error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ExcelNotRunning_LaunchesApplicationBeforeRetryingPermissionPreflight()
    {
        var calls = new List<ProcessStartInfo>();
        var preflightCalls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            if (start.FileName != "/usr/bin/open"
                && AutomationCommand(start) == "session.prepare-open"
                && Interlocked.Increment(ref preflightCalls) == 1)
            {
                return Task.FromResult(new MacProcessResult(
                    0,
                    """{"success":false,"errorMessage":"Automation unavailable (OSStatus -600).","errorCategory":"ComInterop"}""",
                    ""));
            }

            return Task.FromResult(new MacProcessResult(
                0,
                """{"success":true,"errorMessage":""}""",
                ""));
        });
        var path = TestPath("launch-excel.xlsx");

        await backend.InvokeAsync(
            "session.open",
            new { filePath = path },
            TimeSpan.FromSeconds(5));

        Assert.Equal(5, calls.Count);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Equal(["-g", "-a", "Microsoft Excel"], calls[1].ArgumentList.ToArray());
        Assert.Equal("session.prepare-open", AutomationCommand(calls[2]));
        Assert.Equal(["-g", "-b", "com.microsoft.Excel", path], calls[3].ArgumentList.ToArray());
        Assert.Equal("session.open", AutomationCommand(calls[4]));
    }

    [Fact]
    public async Task ColdStartup_RetriesUntilReadyWithoutRepeatingLaunchOrWorkbookHandoff()
    {
        var preflights = 0;
        var launches = 0;
        var handoffs = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                if (start.ArgumentList.Contains("-a")) launches++;
                else handoffs++;
            }
            else if (AutomationCommand(start) == "session.prepare-open" && ++preflights <= 3)
            {
                return Task.FromResult(new MacProcessResult(0,
                    """{"success":false,"errorMessage":"Automation unavailable (OSStatus -600).","errorCategory":"ComInterop"}""", ""));
            }
            return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
        });

        await backend.InvokeAsync("session.open", new { filePath = TestPath("slow-start.xlsx") },
            TimeSpan.FromSeconds(5));

        Assert.Equal(4, preflights);
        Assert.Equal(1, launches);
        Assert.Equal(1, handoffs);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ColdStartup_StopsAtNonStartupFailureOrDeadline(bool neverReady)
    {
        var preflights = 0;
        var launches = 0;
        var handoffs = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.FileName == "/usr/bin/open")
            {
                if (start.ArgumentList.Contains("-a")) launches++;
                else handoffs++;
                return Task.FromResult(new MacProcessResult(0, "", ""));
            }
            preflights++;
            return Task.FromResult(new MacProcessResult(0,
                neverReady || preflights == 1
                    ? """{"success":false,"errorMessage":"Automation unavailable (OSStatus -600).","errorCategory":"ComInterop"}"""
                    : """{"success":false,"errorMessage":"Automation unavailable (OSStatus -1743).","errorCategory":"ComInterop"}""", ""));
        });

        var operation = backend.InvokeAsync("session.open", new { filePath = TestPath("not-ready.xlsx") },
            TimeSpan.FromMilliseconds(500));
        if (neverReady)
        {
            await Assert.ThrowsAsync<TimeoutException>(() => operation);
            Assert.True(preflights >= 2);
        }
        else
        {
            var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => operation);
            Assert.Contains("-1743", error.Message, StringComparison.Ordinal);
            Assert.Equal(2, preflights);
        }
        Assert.Equal(1, launches);
        Assert.Equal(0, handoffs);
    }

    [Theory]
    [InlineData("sheet.list")]
    [InlineData("sheet.create")]
    public async Task WorksheetPath_RejectsAnotherWorkbookBeforeDispatch(string command)
    {
        var sheetDispatches = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.FileName != "/usr/bin/open" && AutomationCommand(start).StartsWith("sheet.", StringComparison.Ordinal))
                sheetDispatches++;
            var result = start.FileName != "/usr/bin/open" && AutomationCommand(start) == "sheet.list"
                ? """{"success":true,"errorMessage":"","worksheets":[{"name":"Sheet1","index":1}]}"""
                : """{"success":true,"errorMessage":""}""";
            return Task.FromResult(new MacProcessResult(0, result, ""));
        });
        var directory = Directory.CreateTempSubdirectory("excelmcp-path-routing-");
        try
        {
            var path = Path.Combine(directory.FullName, "owned.xlsx");
            File.WriteAllText(path, "opaque");
            using var service = new ExcelMcpService(backend);
            var opened = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = JsonSerializer.Serialize(new { filePath = path })
            });
            Assert.True(opened.Success, opened.ErrorMessage);
            using var session = JsonDocument.Parse(Assert.IsType<string>(opened.Result));
            var arguments = new Dictionary<string, object>
            {
                ["filePath"] = Path.Combine(directory.FullName, "other.xlsx")
            };
            if (command == "sheet.create") arguments["sheetName"] = "New";
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = session.RootElement.GetProperty("sessionId").GetString(),
                Args = JsonSerializer.Serialize(arguments)
            });

            Assert.False(response.Success);
            Assert.True(response.ErrorCategory == "PlatformNotSupported", response.ErrorMessage);
            Assert.Contains("session workbook", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, sheetDispatches);

            arguments["filePath"] = Path.Combine(directory.FullName, ".", "owned.xlsx");
            var matchingPath = await service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = session.RootElement.GetProperty("sessionId").GetString(),
                Args = JsonSerializer.Serialize(arguments)
            });
            Assert.True(matchingPath.Success, matchingPath.ErrorMessage);
            Assert.Equal(command == "sheet.create" ? 3 : 1, sheetDispatches);
        }
        finally
        {
            File.Delete(Path.Combine(directory.FullName, "owned.xlsx"));
            directory.Delete();
        }
    }

    [Fact]
    public async Task HandoffFailure_DoesNotAttachWorkbook()
    {
        var count = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            count++;
            return Task.FromResult(start.FileName == "/usr/bin/open"
                ? new MacProcessResult(1, "", "LaunchServices rejected the file")
                : new MacProcessResult(0, """{"success":true}""", ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("test.xlsx") }, TimeSpan.FromSeconds(5)));

        Assert.Equal(2, count);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.Contains("LaunchServices", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("""{"success":false,"errorMessage":"Workbook attachment timed out.","errorCategory":"ComInterop"}""")]
    [InlineData("invalid JSON")]
    [InlineData("{}")]
    [InlineData("""{"success":true,"errorMessage":"Excel rejected the workbook."}""")]
    public async Task Open_UnconfirmedAttachmentRequiresRecoveryInsteadOfAutomaticRetry(string response)
    {
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            return Task.FromResult(new MacProcessResult(
                0, calls == 3 ? response : """{"success":true}""", ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("unconfirmed.xlsx") }, TimeSpan.FromSeconds(5)));

        Assert.Equal(3, calls);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
        Assert.NotNull(error.InnerException);
        Assert.Contains("Do not retry automatically", error.Message, StringComparison.Ordinal);
        Assert.Contains("desktop", error.Message, StringComparison.Ordinal);
        Assert.Contains("dialogs", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Open_TimeoutDuringOrAfterHandoffHasUncertainOutcome(bool duringHandoff)
    {
        var calls = 0;
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            calls++;
            if (calls == (duringHandoff ? 2 : 3))
            {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            }
            return new MacProcessResult(0, """{"success":true}""", "");
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("open-timeout.xlsx") }, TimeSpan.FromMilliseconds(100)));

        Assert.Equal(duringHandoff ? 2 : 3, calls);
        Assert.Equal("RecoveryRequired", error.ErrorCategory);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("create")]
    public async Task SharedService_PreservesRecoveryCategoryAndFileAfterUnconfirmedOpen(string action)
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-service-open-recovery-");
        var path = Path.Combine(directory.FullName, "workbook.xlsx");
        if (action == "open") File.WriteAllBytes(path, [1]);
        var calls = 0;
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls++;
            return Task.FromResult(new MacProcessResult(0, calls == 3
                ? """{"success":false,"errorCategory":"ComInterop","errorMessage":"Attachment not confirmed."}"""
                : """{"success":true}""", ""));
        });
        try
        {
            using var service = new ExcelMcpService(backend);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = $"session.{action}",
                Args = JsonSerializer.Serialize(new { filePath = path, timeoutSeconds = 10 })
            });

            Assert.False(response.Success);
            Assert.Equal("RecoveryRequired", response.ErrorCategory);
            Assert.Equal($"session.{action}", response.Command);
            Assert.Null(response.SessionId);
            Assert.Null(response.Result);
            Assert.Equal(1, service.SessionCount);
            var listed = await service.ProcessAsync(new ServiceRequest { Command = "session.list" });
            Assert.True(listed.Success, listed.ErrorMessage);
            using var inventory = JsonDocument.Parse(Assert.IsType<string>(listed.Result));
            var recovery = Assert.Single(inventory.RootElement.GetProperty("sessions").EnumerateArray());
            Assert.Equal(
                MacPathCanonicalizer.Normalize(path),
                recovery.GetProperty("filePath").GetString());
            Assert.True(recovery.GetProperty("requiresRecovery").GetBoolean());
            Assert.False(recovery.GetProperty("canClose").GetBoolean());
            Assert.Contains("Do not retry automatically", response.ErrorMessage, StringComparison.Ordinal);
            Assert.True(File.Exists(path));
            service.Dispose();
            Assert.Equal(3, calls);
        }
        finally
        {
            File.Delete(path);
            directory.Delete();
        }
    }

    [Fact]
    public async Task DuplicateWorkbookName_DoesNotLaunchHandoff()
    {
        var calls = new List<ProcessStartInfo>();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            calls.Add(start);
            return Task.FromResult(new MacProcessResult(0,
                """{"success":false,"errorMessage":"A workbook with the same name is already open in shared Excel.","errorCategory":"ComInterop"}""",
                ""));
        });

        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("duplicate.xlsx") }, TimeSpan.FromSeconds(5)));

        Assert.Single(calls);
        Assert.Equal("session.prepare-open", AutomationCommand(calls[0]));
        Assert.Contains("same name", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConcurrentOpens_DoNotInterleavePreflightAndHandoff()
    {
        var calls = new List<string>();
        var firstPath = TestPath("first.xlsx");
        var secondPath = TestPath("second.xlsx");
        var firstHandoffStarted = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var releaseFirstHandoff = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            var command = start.FileName == "/usr/bin/open"
                ? $"open:{start.ArgumentList[^1]}"
                : AutomationCommand(start);
            lock (calls)
            {
                calls.Add(command);
            }

            if (command == $"open:{firstPath}")
            {
                firstHandoffStarted.SetResult();
                await releaseFirstHandoff.Task.WaitAsync(cancellationToken);
            }

            return new MacProcessResult(0, """{"success":true,"errorMessage":""}""", "");
        });

        var first = backend.InvokeAsync(
            "session.open", new { filePath = firstPath }, TimeSpan.FromSeconds(5));
        await firstHandoffStarted.Task;
        var second = backend.InvokeAsync(
            "session.open", new { filePath = secondPath }, TimeSpan.FromSeconds(5));
        await Task.Delay(50);

        lock (calls)
        {
            Assert.Equal(["session.prepare-open", $"open:{firstPath}"], calls);
        }

        releaseFirstHandoff.SetResult();
        await Task.WhenAll(first, second);
        lock (calls)
        {
            Assert.Equal(
                [
                    "session.prepare-open",
                    $"open:{firstPath}",
                    "session.open",
                    "session.prepare-open",
                    $"open:{secondPath}",
                    "session.open"
                ],
                calls);
        }
    }

    [Fact]
    public async Task Timeout_BoundsTheWholeOperation()
    {
        var backend = new MacExcelBackend(async (start, input, cancellationToken) =>
        {
            await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            throw new InvalidOperationException("Unreachable.");
        });

        await Assert.ThrowsAsync<TimeoutException>(() => backend.InvokeAsync(
            "session.open", new { filePath = TestPath("test.xlsx") }, TimeSpan.FromMilliseconds(50)));
    }

    private static string AutomationCommand(ProcessStartInfo start)
    {
        var arguments = start.ArgumentList.ToArray();
        var marker = Array.IndexOf(arguments, MacAutomationHost.Marker);
        Assert.True(marker >= 0 && marker + 1 < arguments.Length);
        return arguments[marker + 1];
    }
}
