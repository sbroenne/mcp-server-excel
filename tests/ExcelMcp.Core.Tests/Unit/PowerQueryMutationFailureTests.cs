using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PowerQueryMutationFailureTests
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    [Theory]
    [InlineData("Formula.Firewall", "PrivacyFirewall")]
    [InlineData("Please specify a privacy level for each data source.", "Privacy")]
    [InlineData("DataSource.Error: unable to connect", "Connectivity")]
    [InlineData("Unclassified Excel failure", "ComInterop")]
    public void CreateMutationFailure_PreservesCauseAndDetectedState(string message, string category)
    {
#pragma warning disable CA2201 // Synthetic COM diagnostics for pure error mapping.
        var cause = new COMException(message, unchecked((int)0x800A03EC));
#pragma warning restore CA2201
        var error = PowerQueryCommands.CreateMutationFailure(cause, "Query", true, state =>
        {
            state.QueryExists = true;
            state.LoadMode = PowerQueryLoadMode.LoadToTable;
            state.TargetSheet = "ActualSheet";
            state.IsLoadedToDataModel = false;
        });

        Assert.Same(cause, error.InnerException);
        Assert.Equal(category, error.ErrorCategory);
        Assert.Equal(category, OperationFailureClassifier.Classify(error));
        Assert.Equal("0x800A03EC", OperationFailureClassifier.GetComHResult(error));
        Assert.Contains(message, error.Message, StringComparison.Ordinal);
        Assert.Contains("Query 'Query' was created", error.Message, StringComparison.Ordinal);
        Assert.Contains("ActualSheet", error.Message, StringComparison.Ordinal);
        if (category == "PrivacyFirewall")
        {
            Assert.Contains("powerquery set-privacy-settings", error.Message, StringComparison.Ordinal);
        }
        var state = Assert.IsType<PowerQueryPartialState>(error.PartialState);
        Assert.Equal("Query", state.QueryName);
        Assert.True(state.QueryCreated);
        Assert.Null(state.MCodeUpdated);
        Assert.True(state.QueryExists);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, state.LoadMode);
        Assert.Equal("ActualSheet", state.TargetSheet);
        Assert.False(state.IsLoadedToDataModel);
        Assert.Null(state.StateInspectionError);
    }

    [Fact]
    public void CreateMutationFailure_InspectionFails_PreservesPrimaryErrorAndMarksUnknownState()
    {
        var cause = new InvalidOperationException("Primary load failure");
        var error = PowerQueryCommands.CreateMutationFailure(cause, "Query", false, state =>
        {
            state.QueryExists = true;
            throw new InvalidOperationException("Inspection failed");
        });

        Assert.Same(cause, error.InnerException);
        Assert.Null(error.ErrorCategory);
        Assert.Contains(cause.Message, error.Message, StringComparison.Ordinal);
        Assert.Contains("Inspection failed", error.Message, StringComparison.Ordinal);
        var state = Assert.IsType<PowerQueryPartialState>(error.PartialState);
        Assert.Null(state.QueryCreated);
        Assert.True(state.MCodeUpdated);
        Assert.True(state.QueryExists);
        Assert.Null(state.LoadMode);
        Assert.Null(state.TargetSheet);
        Assert.Null(state.IsLoadedToDataModel);
        Assert.Equal("InvalidOperationException: Inspection failed", state.StateInspectionError);
    }

    [Fact]
    public void CreateMutationFailure_QueryAbsent_DoesNotInventLoadState()
    {
        var error = PowerQueryCommands.CreateMutationFailure(
            new InvalidOperationException("Load failed"), "Query", true,
            state => state.QueryExists = false);

        var state = Assert.IsType<PowerQueryPartialState>(error.PartialState);
        Assert.True(state.QueryCreated);
        Assert.False(state.QueryExists);
        Assert.Null(state.LoadMode);
        Assert.Null(state.IsLoadedToDataModel);
        Assert.Contains("not found", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CreateMutationFailure_ValidationText_DoesNotGuessEngineCategory()
    {
        var error = PowerQueryCommands.CreateMutationFailure(
            new InvalidOperationException("Worksheet 'privacy level' already contains data."),
            "Query", true, state => state.QueryExists = true);

        Assert.Null(error.ErrorCategory);
        Assert.NotNull(error.PartialState);
    }

    [Fact]
    public void CreateMutationFailure_TypedCategory_TakesPrecedenceOverMessage()
    {
        var cause = new OperationFailureException(
            OperationFailureCategory.Permissions, "Worksheet 'privacy level' is protected.");
        var error = PowerQueryCommands.CreateMutationFailure(
            cause, "Query", true, state => state.QueryExists = true);

        Assert.Equal("Permissions", error.ErrorCategory);
        Assert.Same(cause, error.InnerException);
    }

    [Theory]
    [InlineData("cancelled")]
    [InlineData("timeout")]
    [InlineData("invalid-com")]
    [InlineData("categorized-cancelled")]
    [InlineData("disconnected")]
    [InlineData("call-failed")]
    [InlineData("server-unavailable")]
    public void CanReportMutationFailure_UnsafeSessionErrors_AreNotWrapped(string scenario)
    {
#pragma warning disable CA2201 // Synthetic native errors for the reporting guard.
        Exception cause = scenario switch
        {
            "cancelled" => new OperationCanceledException("Formula.Firewall"),
            "timeout" => new TimeoutException("privacy level"),
            "invalid-com" => new InvalidComObjectException("Disconnected"),
            "categorized-cancelled" => new OperationFailureException(
                OperationFailureCategory.Cancelled, "Cancelled by Excel"),
            "disconnected" => new COMException("privacy level", unchecked((int)0x80010108)),
            "call-failed" => new COMException("privacy level", unchecked((int)0x800706BE)),
            "server-unavailable" => new COMException("privacy level", unchecked((int)0x800706BA)),
            _ => throw new ArgumentException("Unknown scenario", nameof(scenario))
        };
#pragma warning restore CA2201
        Assert.False(PowerQueryCommands.CanReportMutationFailure(cause));
        Assert.False(PowerQueryCommands.CanReportMutationFailure(
            new InvalidOperationException("Outer context", cause)));
    }

    [Theory]
    [InlineData(PowerQueryLoadMode.ConnectionOnly, "connection-only")]
    [InlineData(PowerQueryLoadMode.LoadToTable, "load-to-table")]
    [InlineData(PowerQueryLoadMode.LoadToDataModel, "load-to-data-model")]
    [InlineData(PowerQueryLoadMode.LoadToBoth, "load-to-both")]
    public void PartialState_SerializesExistingLoadModeNames(PowerQueryLoadMode mode, string expected)
    {
        var state = new PowerQueryPartialState { QueryName = "Query", LoadMode = mode };
        var json = JsonSerializer.SerializeToElement(state, JsonOptions);
        Assert.Equal(expected, json.GetProperty("loadMode").GetString());
        Assert.False(json.TryGetProperty("queryExists", out _));
        Assert.False(json.TryGetProperty("isLoadedToDataModel", out _));
    }
}
