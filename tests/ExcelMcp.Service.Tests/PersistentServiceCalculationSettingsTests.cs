using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "CalculationMode")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceCalculationSettingsTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void Settings_ReadsNativeIterationAndPrecisionState()
    {
        var response = _fixture.Send("calculationmode.get-settings", new { });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal("application", result.RootElement.GetProperty("settingsScope").GetString());
        Assert.False(result.RootElement.GetProperty("precisionAsDisplayed").GetBoolean());
        Assert.True(result.RootElement.GetProperty("maximumIterations").GetInt32() > 0);
        Assert.True(result.RootElement.GetProperty("maximumChange").GetDouble() > 0);
        Assert.Equal(0, result.RootElement.GetProperty("calculationStateValue").GetInt32());
        Assert.Equal("done", result.RootElement.GetProperty("calculationState").GetString());
    }

    [Fact]
    public void Settings_ChangesIterationWithoutChangingModeAndRestoresIt()
    {
        var previous = _fixture.Send("calculationmode.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var failure = Record.Exception(() =>
        {
            _fixture.Send("calculationmode.set-settings", new
            {
                iterationEnabled = true,
                maximumIterations = 37,
                maximumChange = 0.0002
            });
            var response = _fixture.Send("calculationmode.get-settings", new { });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("iterationEnabled").GetBoolean());
            Assert.Equal(37, result.RootElement.GetProperty("maximumIterations").GetInt32());
            Assert.Equal(0.0002, result.RootElement.GetProperty("maximumChange").GetDouble());
            Assert.Equal(original.RootElement.GetProperty("modeValue").GetInt32(),
                result.RootElement.GetProperty("modeValue").GetInt32());
        });
        RestoreSettingsAndThrowFailure(original.RootElement, failure);
    }

    [Theory]
    [InlineData("full")]
    [InlineData("rebuild")]
    public void Calculate_FullModesActuallyRecalculateAndPreserveManualMode(string kind)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var previous = _fixture.Send("calculationmode.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var failure = Record.Exception(() =>
        {
            Assert.True(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual).Success);
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[7d]]));
            RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, sheetName, "B1", [["=A1*3"]]));
            _fixture.Send("calculationmode.calculate", new { scope = "application" });
            AssertValue(21d);
            RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[11d]],
                overwritePolicy: OverwritePolicy.Allow));
            AssertValue(21d);
            _fixture.Send("calculationmode.calculate", new { scope = "application", kind });
            AssertValue(33d);
            var retained = calculation.GetSettings(_fixture.BatchToken);
            Assert.True(retained.Success);
            Assert.Equal((int)CalculationMode.Manual, retained.ModeValue);
        });
        RestoreSettingsAndThrowFailure(original.RootElement, failure);

        void AssertValue(double expected)
        {
            var values = _commands.GetValues(_fixture.BatchToken, sheetName, "B1");
            RequireSuccess(values);
            Assert.Equal(expected, Convert.ToDouble(values.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        }
    }

    [Theory]
    [InlineData(0, 0.001)]
    [InlineData(32768, 0.001)]
    [InlineData(100, 0)]
    [InlineData(100, -1)]
    public async Task Settings_InvalidLimitsRejectBeforeChangingMode(int iterations, double change)
    {
        var previous = _fixture.Send("calculationmode.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var rejected = await _fixture.SendForFailureAsync("calculationmode.set-settings", new
        {
            mode = "manual",
            iterationEnabled = !original.RootElement.GetProperty("iterationEnabled").GetBoolean(),
            maximumIterations = iterations,
            maximumChange = change
        });
        Assert.False(rejected.Success);
        Assert.Equal("InvalidInput", rejected.ErrorCategory);
        Assert.Equal(nameof(ArgumentOutOfRangeException), rejected.ExceptionType);
        Assert.Contains(iterations is < 1 or > 32767 ? "maximumIterations" : "maximumChange",
            rejected.ErrorMessage, StringComparison.Ordinal);
        var retained = _fixture.Send("calculationmode.get-settings", new { });
        using var result = JsonDocument.Parse(retained.Result!);
        Assert.Equal(original.RootElement.GetProperty("modeValue").GetInt32(),
            result.RootElement.GetProperty("modeValue").GetInt32());
        Assert.Equal(original.RootElement.GetProperty("maximumIterations").GetInt32(),
            result.RootElement.GetProperty("maximumIterations").GetInt32());
        Assert.Equal(original.RootElement.GetProperty("iterationEnabled").GetBoolean(),
            result.RootElement.GetProperty("iterationEnabled").GetBoolean());
        Assert.Equal(original.RootElement.GetProperty("maximumChange").GetDouble(),
            result.RootElement.GetProperty("maximumChange").GetDouble());
    }

    [Theory]
    [InlineData("sheet", null, null, "sheetName")]
    [InlineData("sheet", "", null, "sheetName")]
    [InlineData("sheet", " \t ", null, "sheetName")]
    [InlineData("range", null, "A1", "sheetName")]
    [InlineData("range", "", "A1", "sheetName")]
    [InlineData("range", " \t ", "A1", "sheetName")]
    [InlineData("range", "target", null, "rangeAddress")]
    [InlineData("range", "target", "", "rangeAddress")]
    [InlineData("range", "target", " \t ", "rangeAddress")]
    public async Task Calculate_MissingTargetsReturnInvalidInputWithoutCalculation(
        string scope, string? sheetName, string? rangeAddress, string parameter)
    {
        var targetSheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var previous = _fixture.Send("calculationmode.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var failure = await Record.ExceptionAsync(async () =>
        {
            _fixture.Send("calculationmode.set-settings", new { mode = "manual" });
            Assert.True(_commands.SetValues(_fixture.BatchToken, targetSheet, "A1", [[10d]]).Success);
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, targetSheet, "B1", [["=A1+5"]]).Success);
            _fixture.Send("calculationmode.calculate", new { scope = "application" });
            Assert.True(_commands.SetValues(_fixture.BatchToken, targetSheet, "A1", [[99d]],
                overwritePolicy: OverwritePolicy.Allow).Success);
            var beforeValues = _commands.GetValues(_fixture.BatchToken, targetSheet, "B1");
            Assert.True(beforeValues.Success);
            Assert.Equal(15d, Convert.ToDouble(beforeValues.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
            var beforeSettings = _fixture.Send("calculationmode.get-settings", new { });

            var rejected = await _fixture.SendForFailureAsync("calculationmode.calculate", new
            {
                scope,
                sheetName = sheetName == "target" ? targetSheet : sheetName,
                rangeAddress
            });

            Assert.False(rejected.Success);
            Assert.Equal("InvalidInput", rejected.ErrorCategory);
            Assert.Equal(nameof(ArgumentException), rejected.ExceptionType);
            Assert.Equal("calculationmode.calculate", rejected.Command);
            Assert.Null(rejected.Result);
            Assert.Contains(parameter, rejected.ErrorMessage, StringComparison.Ordinal);
            var afterSettings = _fixture.Send("calculationmode.get-settings", new { });
            Assert.Equal(beforeSettings.Result, afterSettings.Result);
            var afterValues = _commands.GetValues(_fixture.BatchToken, targetSheet, "B1");
            Assert.True(afterValues.Success);
            Assert.Equal(15d, Convert.ToDouble(afterValues.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        });
        RestoreSettingsAndThrowFailure(original.RootElement, failure);
    }

    [Fact]
    public async Task Precision_RequiresPermissionAndDisablingDoesNotRecoverDigits()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1.2345]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "A1", "0.00").Success);
        var rejected = await _fixture.SendForFailureAsync("calculationmode.set-precision", new
        {
            precisionAsDisplayed = true
        });
        Assert.False(rejected.Success);
        var unchanged = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(unchanged.Success);
        Assert.Equal(1.2345, Convert.ToDouble(unchanged.Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture));
        var failure = Record.Exception(() =>
        {
            var enabled = _fixture.Send("calculationmode.set-precision", new
            {
                precisionAsDisplayed = true,
                allowPrecisionLoss = true
            });
            using var result = JsonDocument.Parse(enabled.Result!);
            Assert.True(result.RootElement.GetProperty("precisionAsDisplayed").GetBoolean());
            _fixture.Send("calculationmode.set-precision", new { precisionAsDisplayed = false });
            var rounded = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
            Assert.True(rounded.Success);
            Assert.Equal(1.23, Convert.ToDouble(rounded.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        });
        var cleanup = Record.Exception(() =>
            _fixture.Send("calculationmode.set-precision", new { precisionAsDisplayed = false }));
        ThrowCombinedFailure(failure, cleanup);
    }

    [Fact]
    public void Settings_IterationActuallyEvaluatesCircularFormulaAndRestoresSettings()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var previous = _fixture.Send("calculationmode.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var failure = Record.Exception(() =>
        {
            _fixture.Send("calculationmode.set-settings", new
            {
                mode = "manual",
                iterationEnabled = true,
                maximumIterations = 10,
                maximumChange = 0.00001
            });
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1",
                [["=IF(A1<5,A1+1,A1)"]]).Success);
            _fixture.Send("calculationmode.calculate", new { scope = "application", kind = "rebuild" });
            var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
            Assert.True(values.Success);
            Assert.Equal(5d, Convert.ToDouble(values.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        });
        var cleanup = Record.Exception(() =>
            _fixture.Send("range.clear-contents", new { sheetName, rangeAddress = "A1" }));
        if (cleanup is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        RestoreSettingsAndThrowFailure(original.RootElement, failure);
    }

    private void RestoreSettingsAndThrowFailure(JsonElement original, Exception? failure)
    {
        var cleanup = Record.Exception(() =>
            _fixture.Send("calculationmode.set-settings", new
            {
                mode = original.GetProperty("mode").GetString(),
                iterationEnabled = original.GetProperty("iterationEnabled").GetBoolean(),
                maximumIterations = original.GetProperty("maximumIterations").GetInt32(),
                maximumChange = original.GetProperty("maximumChange").GetDouble()
            }));
        ThrowCombinedFailure(failure, cleanup);
    }

    private static void ThrowCombinedFailure(Exception? failure, Exception? cleanup)
    {
        if (cleanup is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }
}
