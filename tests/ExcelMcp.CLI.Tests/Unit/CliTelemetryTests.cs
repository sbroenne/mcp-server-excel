using Sbroenne.ExcelMcp.CLI.Telemetry;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Telemetry")]
[Trait("Speed", "Fast")]
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
            (command, _, succeeded, _) =>
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

    [Fact]
    public async Task TrackCommandAsync_ClassifiesCancellationWhenNoResponseIsReturned()
    {
        var request = new ServiceRequest { Command = "range.get-values" };
        string? trackedCategory = null;
        bool? trackedSuccess = null;

        await Assert.ThrowsAsync<OperationCanceledException>(() => CliTelemetry.TrackCommandAsync(
            request,
            () => Task.FromException<ServiceResponse>(new OperationCanceledException()),
            (_, _, succeeded, errorCategory) =>
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
            (command, _, succeeded, _) =>
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
            (command, _, _, _) => trackedCommand = command);

        Assert.Equal("calculation.get", trackedCommand);
    }

    [Fact]
    public void TrackCliInvocation_DoesNotDuplicateRequestTelemetry()
    {
        var trackedCommands = new List<string>();

        var exitCode = CliTelemetry.TrackCliInvocation(
            ["session", "list"],
            () =>
            {
                _ = CliTelemetry.TrackCommandAsync(
                    new ServiceRequest { Command = "session.list" },
                    () => Task.FromResult(new ServiceResponse { Success = true }),
                    (command, _, _, _) => trackedCommands.Add(command)).GetAwaiter().GetResult();
                return 0;
            },
            (command, _, _, _) => trackedCommands.Add(command));

        Assert.Equal(0, exitCode);
        Assert.Equal(["session.list"], trackedCommands);
    }

    [Fact]
    public void TrackCliInvocation_SkipsHelpRequests()
    {
        var tracked = false;

        CliTelemetry.TrackCliInvocation(
            ["sheet", "--help"],
            () => 0,
            (_, _, _, _) => tracked = true);

        Assert.False(tracked);
    }
}
