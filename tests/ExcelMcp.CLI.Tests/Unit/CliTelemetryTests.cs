using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class CliTelemetryTests
{
    [Fact]
    public async Task TrackCommandAsync_TracksTheExecutedCliCommand()
    {
        var request = new ServiceRequest { Command = "range.get-values" };
        string? trackedCommand = null;
        bool? trackedSuccess = null;

        var response = await CliTelemetry.TrackCommandAsync(
            request,
            () => Task.FromResult(new ServiceResponse { Success = true }),
            (command, _, succeeded, _, _) =>
            {
                trackedCommand = command;
                trackedSuccess = succeeded;
            });

        Assert.True(response.Success);
        Assert.Equal("range.get-values", trackedCommand);
        Assert.True(trackedSuccess);
    }

    [Fact]
    public void CreateCommandInvocationTelemetry_IdentifiesCliEntryPoint()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "range.get-values",
                25,
                succeeded: true,
                errorCategory: null);

        Assert.Equal("range/get-values", eventTelemetry.Name);
        Assert.Equal("cli", eventTelemetry.Properties["EntryPoint"]);
        Assert.Equal("range", eventTelemetry.Properties["Tool"]);
        Assert.Equal("get-values", eventTelemetry.Properties["Action"]);
        Assert.Equal("cli", requestTelemetry.Properties["EntryPoint"]);
        Assert.True(requestTelemetry.Success);
    }

    [Fact]
    public void CreateCommandInvocationTelemetry_DoesNotIncludeFailureDetails()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "session.open",
                25,
                succeeded: false,
                errorCategory: "Permissions");

        Assert.Equal("external-dependency", eventTelemetry.Properties["FailureClass"]);
        Assert.Equal("external-dependency", requestTelemetry.Properties["FailureClass"]);
        Assert.DoesNotContain(
            eventTelemetry.Properties.Keys,
            key => key.Contains("error", StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("Timeout", "timeout")]
    [InlineData("Cancelled", "cancellation")]
    public void CreateCommandInvocationTelemetry_DistinguishesTimeoutFromCancellation(
        string errorCategory,
        string expectedFailureCause)
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "vba.run",
                25,
                succeeded: false,
                errorCategory);

        Assert.Equal("timeout-cancellation", eventTelemetry.Properties["FailureClass"]);
        Assert.Equal(expectedFailureCause, eventTelemetry.Properties["FailureCause"]);
        Assert.Equal(expectedFailureCause, requestTelemetry.Properties["FailureCause"]);
    }

    [Fact]
    public void CreateCommandInvocationTelemetry_IdentifiesExpectedNegativeOutcome()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                "session.test",
                25,
                succeeded: true,
                errorCategory: null,
                expectedNegative: true);

        Assert.Equal("expected-negative", eventTelemetry.Properties["Outcome"]);
        Assert.True(requestTelemetry.Success);
        Assert.DoesNotContain("FailureClass", eventTelemetry.Properties.Keys);
    }

    [Fact]
    public async Task TrackCommandAsync_ClassifiesCancellationWhenNoResponseIsReturned()
    {
        var request = new ServiceRequest { Command = "range.get-values" };
        string? trackedCategory = null;
        bool? trackedSuccess = null;

        await Assert.ThrowsAsync<OperationCanceledException>(() => CliTelemetry.TrackCommandAsync(
            request,
            () => Task.FromException<ServiceResponse>(new OperationCanceledException()),
            (_, _, succeeded, errorCategory, _) =>
            {
                trackedCategory = errorCategory;
                trackedSuccess = succeeded;
            }));

        Assert.False(trackedSuccess);
        Assert.Equal("Cancelled", trackedCategory);
        var (eventTelemetry, _) = CliTelemetry.CreateCommandInvocationTelemetry(
            request.Command,
            25,
            succeeded: false,
            errorCategory: trackedCategory);
        Assert.Equal("timeout-cancellation", eventTelemetry.Properties["FailureClass"]);
    }

    [Fact]
    public void CreateCommandInvocationTelemetry_ReplacesCommandsOutsideTheAllowlist()
    {
        var (eventTelemetry, requestTelemetry) =
            CliTelemetry.CreateCommandInvocationTelemetry(
                @"C:\customer\Q3 forecast.xlsx.secret-action",
                25,
                succeeded: true,
                errorCategory: null);

        Assert.Equal("other/other", eventTelemetry.Name);
        Assert.Equal("other", eventTelemetry.Properties["Tool"]);
        Assert.Equal("other", eventTelemetry.Properties["Action"]);
        Assert.DoesNotContain(
            eventTelemetry.Properties.Values.Concat(requestTelemetry.Properties.Values),
            value => value.Contains("forecast", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void TrackCliInvocation_TracksCommandsThatNeverSendARequest()
    {
        string? trackedCommand = null;
        bool? trackedSuccess = null;

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["service", "start"],
            () => 1,
            (command, _, succeeded, _, _) =>
            {
                trackedCommand = command;
                trackedSuccess = succeeded;
            });

        Assert.Equal(1, exitCode);
        Assert.Equal("service.start", trackedCommand);
        Assert.False(trackedSuccess);
    }

    [Fact]
    public void TrackCliInvocation_UsesCanonicalCategoryForCliCommandNames()
    {
        string? trackedCommand = null;

        CliTelemetry.TrackCliInvocation(
            ["calculationmode", "get", @"--file", @"C:\customer\Q3 forecast.xlsx"],
            () => 0,
            (command, _, _, _, _) => trackedCommand = command);

        Assert.Equal("calculationmode.get", trackedCommand);
    }

    [Fact]
    public void TrackCliInvocation_ReportsExpectedNegativeDiagnosticOutcome()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded, bool ExpectedNegative)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["session", "test"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.test" },
                    () => Task.FromResult(new ServiceResponse { Success = true }),
                    (command, _, succeeded, _, expectedNegative) =>
                        trackedInvocations.Add((command, succeeded, expectedNegative))).GetAwaiter().GetResult();
                CliTelemetry.RecordExpectedNegative();
                return 1;
            },
            (command, _, succeeded, _, expectedNegative) =>
                trackedInvocations.Add((command, succeeded, expectedNegative)));

        Assert.Equal(1, exitCode);
        Assert.Equal([("session.test", true, true)], trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_PreservesFailedServiceResponseCategory()
    {
        string? trackedCategory = null;

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["session", "open"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.open" },
                    () => Task.FromResult(new ServiceResponse { Success = false, ErrorCategory = "InvalidInput" }),
                    (_, _, _, errorCategory, _) => trackedCategory = errorCategory).GetAwaiter().GetResult();
                return 1;
            },
            (_, _, _, errorCategory, _) => trackedCategory = errorCategory);

        Assert.Equal(1, exitCode);
        Assert.Equal("InvalidInput", trackedCategory);
        var (eventTelemetry, _) = CliTelemetry.CreateCommandInvocationTelemetry(
            "session.open",
            25,
            succeeded: false,
            errorCategory: trackedCategory);
        Assert.Equal("input-state", eventTelemetry.Properties["FailureClass"]);
    }

    [Fact]
    public void TrackCliInvocation_PreservesCaughtRequestExceptionCategory()
    {
        string? trackedCategory = null;

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["service", "stop"],
            () =>
            {
                try
                {
                    _ = CliTelemetry.TrackCommandAsync(
                        new ServiceRequest { Command = "service.shutdown" },
                        () => Task.FromException<ServiceResponse>(new OperationCanceledException()),
                        (_, _, _, errorCategory, _) => trackedCategory = errorCategory).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException)
                {
                }

                return 1;
            },
            (_, _, _, errorCategory, _) => trackedCategory = errorCategory);

        Assert.Equal(1, exitCode);
        Assert.Equal("Cancelled", trackedCategory);
    }

    [Fact]
    public void TrackCliInvocation_PreservesPostResponseFailureCategory()
    {
        string? trackedCategory = null;

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["session", "list"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.list" },
                    () => Task.FromResult(new ServiceResponse { Success = true }),
                    (_, _, _, errorCategory, _) => trackedCategory = errorCategory).GetAwaiter().GetResult();
                CliTelemetry.RecordFinalFailure("InvalidResponse");
                return 1;
            },
            (_, _, _, errorCategory, _) => trackedCategory = errorCategory);

        Assert.Equal(1, exitCode);
        Assert.Equal("InvalidResponse", trackedCategory);
    }

    [Fact]
    public void TrackCliInvocation_PreservesPerItemBatchTelemetry()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["batch", "--input", "commands.json"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.list" },
                    () => Task.FromResult(new ServiceResponse { Success = true }),
                    (command, _, succeeded, _, _) => trackedInvocations.Add((command, succeeded))).GetAwaiter().GetResult();
                return 0;
            },
            (command, _, succeeded, _, _) => trackedInvocations.Add((command, succeeded)));

        Assert.Equal(0, exitCode);
        Assert.Equal([("session.list", true)], trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_PreservesExpectedNegativeBatchTelemetry()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded, bool ExpectedNegative)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["batch", "--input", "commands.json"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.test" },
                    () => Task.FromResult(new ServiceResponse
                    {
                        Success = true,
                        Result = """{"canOpen":false}"""
                    }),
                    (command, _, succeeded, _, expectedNegative) =>
                        trackedInvocations.Add((command, succeeded, expectedNegative))).GetAwaiter().GetResult();
                return 1;
            },
            (command, _, succeeded, _, expectedNegative) =>
                trackedInvocations.Add((command, succeeded, expectedNegative)));

        Assert.Equal(1, exitCode);
        Assert.Equal([("session.test", true, true)], trackedInvocations);
    }

    [Fact]
    public async Task TrackCommandAsync_RefreshAllFailure_TracksFailureWithResultCategory()
    {
        var tracked = new List<(string Command, bool Succeeded, string? ErrorCategory, bool ExpectedNegative)>();
        var refreshAll = new PowerQueryRefreshAllResult
        {
            Success = false,
            ErrorMessage = "1 of 2 queries failed to refresh: 'Broken'.",
            RefreshedQueries = ["Good"],
            FailedQueries =
            [
                new PowerQueryRefreshFailure
                {
                    QueryName = "Broken",
                    ErrorCategory = "Expression",
                    ErrorMessage = "[Expression.Error] The name 'Missing' wasn't recognized.",
                    ExceptionType = "COMException",
                    HResult = "0x800A03EC"
                }
            ]
        };

        await CliTelemetry.TrackCommandAsync(
            new ServiceRequest { Command = "powerquery.refresh-all" },
            () => Task.FromResult(new ServiceResponse
            {
                Success = true,
                Result = JsonSerializer.Serialize(refreshAll, ServiceProtocol.JsonOptions)
            }),
            (command, _, succeeded, errorCategory, expectedNegative) =>
                tracked.Add((command, succeeded, errorCategory, expectedNegative)));

        Assert.Equal([("powerquery.refresh-all", false, "Expression", false)], tracked);
    }

    [Fact]
    public async Task TrackCommandAsync_FileTestCanOpenFalse_RemainsExpectedNegative()
    {
        var tracked = new List<(string Command, bool Succeeded, string? ErrorCategory, bool ExpectedNegative)>();

        await CliTelemetry.TrackCommandAsync(
            new ServiceRequest { Command = "session.test" },
            () => Task.FromResult(new ServiceResponse
            {
                Success = true,
                Result = """{"success":false,"canOpen":false}"""
            }),
            (command, _, succeeded, errorCategory, expectedNegative) =>
                tracked.Add((command, succeeded, errorCategory, expectedNegative)));

        Assert.Equal([("session.test", true, (string?)null, true)], tracked);
    }

    [Fact]
    public void TrackCliInvocation_PreservesLocallyRejectedBatchItems()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded, string? ErrorCategory)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["batch", "--input", "commands.json"],
            () =>
            {
                CliTelemetry.TrackLocalFailure(
                    "range.set-values",
                    5,
                    "InvalidInput",
                    (command, _, succeeded, errorCategory, _) =>
                        trackedInvocations.Add((command, succeeded, errorCategory)));
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.list" },
                    () => Task.FromResult(new ServiceResponse { Success = true }),
                    (command, _, succeeded, errorCategory, _) =>
                        trackedInvocations.Add((command, succeeded, errorCategory))).GetAwaiter().GetResult();
                return 1;
            },
            (command, _, succeeded, errorCategory, _) =>
                trackedInvocations.Add((command, succeeded, errorCategory)));

        Assert.Equal(1, exitCode);
        Assert.Equal(
            [
                ("range.set-values", false, "InvalidInput"),
                ("session.list", true, null)
            ],
            trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_TracksBatchFailureAfterLocallyRejectedItem()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded, string? ErrorCategory)>();

        Assert.Throws<IOException>(() =>
            CliTelemetry.TrackCliInvocation(
                ["batch", "--input", "commands.json"],
                () =>
                {
                    CliTelemetry.TrackLocalFailure(
                        "range.set-values",
                        5,
                        "InvalidInput",
                        (command, _, succeeded, errorCategory, _) =>
                            trackedInvocations.Add((command, succeeded, errorCategory)));
                    throw new IOException("Connection failed.");
                },
                (command, _, succeeded, errorCategory, _) =>
                    trackedInvocations.Add((command, succeeded, errorCategory))));

        Assert.Equal(
            [
                ("range.set-values", false, "InvalidInput"),
                ("batch.run", false, (string?)null)
            ],
            trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_TracksBatchWhenNoItemsAreTracked()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["batch", "--input", "empty.json"],
            () => 1,
            (command, _, succeeded, _, _) => trackedInvocations.Add((command, succeeded)));

        Assert.Equal(1, exitCode);
        Assert.Equal([("batch.run", false)], trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_UsesSuccessfulFinalOutcomeAfterExpectedRequestFailure()
    {
        var trackedInvocations = new List<(string Command, bool Succeeded)>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["service", "status"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "service.ping" },
                    () => Task.FromResult(new ServiceResponse { Success = false, ErrorCategory = "ServiceUnavailable" }),
                    (command, _, succeeded, _, _) => trackedInvocations.Add((command, succeeded))).GetAwaiter().GetResult();
                return 0;
            },
            (command, _, succeeded, _, _) => trackedInvocations.Add((command, succeeded)));

        Assert.Equal(0, exitCode);
        Assert.Equal([("service.status", true)], trackedInvocations);
    }

    [Fact]
    public void TrackCliInvocation_SkipsHelpRequests()
    {
        var tracked = false;

        CliTelemetry.TrackCliInvocation(
            ["sheet", "--help"],
            () => 0,
            (_, _, _, _, _) => tracked = true);

        Assert.False(tracked);
    }

    [Theory]
    [InlineData("--help")]
    [InlineData("-h")]
    [InlineData("--HELP")]
    [InlineData("-H")]
    [InlineData("sheet --help")]
    [InlineData("--help sheet")]
    [InlineData("range get-values --file book.xlsx -h")]
    public void IsHelpRequest_HelpFlagAnywhere_ReturnsTrue(string commandLine)
    {
        Assert.True(CliTelemetry.IsHelpRequest(commandLine.Split(' ')));
    }

    [Theory]
    [InlineData("diag ping")]
    [InlineData("sheet list --file book.xlsx")]
    [InlineData("service start")]
    [InlineData("range get-values --helpful")]
    public void IsHelpRequest_NoHelpFlag_ReturnsFalse(string commandLine)
    {
        Assert.False(CliTelemetry.IsHelpRequest(commandLine.Split(' ')));
    }

    [Fact]
    public void IsHelpRequest_NoArguments_ReturnsFalse()
    {
        Assert.False(CliTelemetry.IsHelpRequest([]));
    }
}
