using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
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
        var response = _fixture.Send("calculation.get-settings", new { });
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
        var previous = _fixture.Send("calculation.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        try
        {
            _fixture.Send("calculation.set-settings", new
            {
                iterationEnabled = true,
                maximumIterations = 37,
                maximumChange = 0.0002
            });
            var response = _fixture.Send("calculation.get-settings", new { });
            using var result = JsonDocument.Parse(response.Result!);
            Assert.True(result.RootElement.GetProperty("iterationEnabled").GetBoolean());
            Assert.Equal(37, result.RootElement.GetProperty("maximumIterations").GetInt32());
            Assert.Equal(0.0002, result.RootElement.GetProperty("maximumChange").GetDouble());
            Assert.Equal(original.RootElement.GetProperty("modeValue").GetInt32(),
                result.RootElement.GetProperty("modeValue").GetInt32());
        }
        finally
        {
            _fixture.Send("calculation.set-settings", new
            {
                iterationEnabled = original.RootElement.GetProperty("iterationEnabled").GetBoolean(),
                maximumIterations = original.RootElement.GetProperty("maximumIterations").GetInt32(),
                maximumChange = original.RootElement.GetProperty("maximumChange").GetDouble()
            });
        }
    }

    [Theory]
    [InlineData("full")]
    [InlineData("rebuild")]
    public void Calculate_FullModesActuallyRecalculateAndPreserveManualMode(string kind)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var calculation = _fixture.CreateCommands<ICalculationModeCommands>();
        var previous = calculation.GetSettings(_fixture.BatchToken);
        Assert.True(previous.Success);
        try
        {
            Assert.True(calculation.SetSettings(_fixture.BatchToken, CalculationMode.Manual).Success);
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1:B1",
                [["=7", "=A1*3"]]).Success);
            _fixture.Send("calculation.calculate", new { scope = "application", kind });
            var values = _commands.GetValues(_fixture.BatchToken, sheetName, "B1");
            Assert.True(values.Success);
            Assert.Equal(21d, Convert.ToDouble(values.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
            var retained = calculation.GetSettings(_fixture.BatchToken);
            Assert.True(retained.Success);
            Assert.Equal((int)CalculationMode.Manual, retained.ModeValue);
        }
        finally
        {
            Assert.True(calculation.SetSettings(_fixture.BatchToken, (CalculationMode)previous.ModeValue).Success);
        }
    }

    [Theory]
    [InlineData(0, 0.001)]
    [InlineData(32768, 0.001)]
    [InlineData(100, 0)]
    [InlineData(100, -1)]
    public async Task Settings_InvalidLimitsRejectBeforeChangingMode(int iterations, double change)
    {
        var previous = _fixture.Send("calculation.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        var rejected = await _fixture.SendForFailureAsync("calculation.set-settings", new
        {
            mode = "manual",
            maximumIterations = iterations,
            maximumChange = change
        });
        Assert.False(rejected.Success);
        var retained = _fixture.Send("calculation.get-settings", new { });
        using var result = JsonDocument.Parse(retained.Result!);
        Assert.Equal(original.RootElement.GetProperty("modeValue").GetInt32(),
            result.RootElement.GetProperty("modeValue").GetInt32());
        Assert.Equal(original.RootElement.GetProperty("maximumIterations").GetInt32(),
            result.RootElement.GetProperty("maximumIterations").GetInt32());
    }

    [Fact]
    public async Task Precision_RequiresPermissionAndDisablingDoesNotRecoverDigits()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[1.2345]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheetName, "A1", "0.00").Success);
        var rejected = await _fixture.SendForFailureAsync("calculation.set-precision", new
        {
            precisionAsDisplayed = true
        });
        Assert.False(rejected.Success);
        var unchanged = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(unchanged.Success);
        Assert.Equal(1.2345, Convert.ToDouble(unchanged.Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture));
        try
        {
            var enabled = _fixture.Send("calculation.set-precision", new
            {
                precisionAsDisplayed = true,
                allowPrecisionLoss = true
            });
            using var result = JsonDocument.Parse(enabled.Result!);
            Assert.True(result.RootElement.GetProperty("precisionAsDisplayed").GetBoolean());
            _fixture.Send("calculation.set-precision", new { precisionAsDisplayed = false });
            var rounded = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
            Assert.True(rounded.Success);
            Assert.Equal(1.23, Convert.ToDouble(rounded.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        }
        finally
        {
            _fixture.Send("calculation.set-precision", new { precisionAsDisplayed = false });
        }
    }

    [Fact]
    public void Settings_IterationActuallyEvaluatesCircularFormulaAndRestoresSettings()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var previous = _fixture.Send("calculation.get-settings", new { });
        using var original = JsonDocument.Parse(previous.Result!);
        try
        {
            _fixture.Send("calculation.set-settings", new
            {
                mode = "manual",
                iterationEnabled = true,
                maximumIterations = 10,
                maximumChange = 0.00001
            });
            Assert.True(_commands.SetFormulas(_fixture.BatchToken, sheetName, "A1",
                [["=IF(A1<5,A1+1,A1)"]]).Success);
            _fixture.Send("calculation.calculate", new { scope = "application", kind = "rebuild" });
            var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
            Assert.True(values.Success);
            Assert.Equal(5d, Convert.ToDouble(values.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        }
        finally
        {
            try
            {
                _fixture.Send("range.clear-contents", new { sheetName, rangeAddress = "A1" });
            }
            finally
            {
                _fixture.Send("calculation.set-settings", new
                {
                    mode = original.RootElement.GetProperty("mode").GetString(),
                    iterationEnabled = original.RootElement.GetProperty("iterationEnabled").GetBoolean(),
                    maximumIterations = original.RootElement.GetProperty("maximumIterations").GetInt32(),
                    maximumChange = original.RootElement.GetProperty("maximumChange").GetDouble()
                });
            }
        }
    }
}
