using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "CalculationMode")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceCalculationTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly ICalculationModeCommands _calculation =
        fixture.CreateCommands<ICalculationModeCommands>();

    [Fact]
    public void GetMode_ReturnsAutomaticByDefault()
    {
        var result = _calculation.GetMode(_fixture.BatchToken);

        Assert.True(result.Success);
        Assert.Equal("automatic", result.Mode);
        Assert.Equal(-4105, result.ModeValue);
    }

    [Theory]
    [InlineData(CalculationMode.Automatic, false)]
    [InlineData(CalculationMode.Manual, false)]
    [InlineData(CalculationMode.SemiAutomatic, false)]
    [InlineData(CalculationMode.Automatic, true)]
    [InlineData(CalculationMode.Manual, true)]
    [InlineData(CalculationMode.SemiAutomatic, true)]
    public void RangeWrites_PreserveModeAndRecalculateOrdinaryDependentsByMode(
        CalculationMode mode,
        bool writeFormula)
    {
        var previous = _calculation.GetMode(_fixture.BatchToken);
        Assert.True(previous.Success, previous.ErrorMessage);
        var range = _fixture.CreateCommands<IRangeCommands>();
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        try
        {
            Assert.True(_calculation.SetMode(_fixture.BatchToken, CalculationMode.Manual).Success);
            var seed = range.SetFormulas(_fixture.BatchToken, sheetName, "A1:B1", [["=2", "=A1*2"]]);
            Assert.True(seed.Success, seed.ErrorMessage);
            Assert.True(_calculation.Calculate(_fixture.BatchToken, CalculationScope.Workbook).Success);
            var baseline = range.GetValues(_fixture.BatchToken, sheetName, "B1");
            Assert.True(baseline.Success, baseline.ErrorMessage);
            Assert.Equal(4d, Convert.ToDouble(baseline.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

            Assert.True(_calculation.SetMode(_fixture.BatchToken, mode).Success);
            var write = writeFormula
                ? range.SetFormulas(_fixture.BatchToken, sheetName, "A1", [["=3"]], overwritePolicy: OverwritePolicy.Allow)
                : range.SetValues(_fixture.BatchToken, sheetName, "A1", [[3]], overwritePolicy: OverwritePolicy.Allow);
            Assert.True(write.Success, write.ErrorMessage);
            var retainedMode = _calculation.GetMode(_fixture.BatchToken);
            Assert.True(retainedMode.Success, retainedMode.ErrorMessage);
            Assert.Equal((int)mode, retainedMode.ModeValue);
            var beforeExplicitCalculation = range.GetValues(_fixture.BatchToken, sheetName, "B1");
            Assert.True(beforeExplicitCalculation.Success, beforeExplicitCalculation.ErrorMessage);
            Assert.Equal(mode == CalculationMode.Manual ? 4d : 6d,
                Convert.ToDouble(beforeExplicitCalculation.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

            Assert.True(_calculation.Calculate(_fixture.BatchToken, CalculationScope.Workbook).Success);
            var calculated = range.GetValues(_fixture.BatchToken, sheetName, "B1");
            Assert.True(calculated.Success, calculated.ErrorMessage);
            Assert.Equal(6d, Convert.ToDouble(calculated.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        }
        finally
        {
            var restored = _calculation.SetMode(_fixture.BatchToken, (CalculationMode)previous.ModeValue);
            Assert.True(restored.Success, restored.ErrorMessage);
        }
    }

    [Fact]
    public void SetMode_ToManual_Succeeds()
    {
        AssertModeRoundTrip(CalculationMode.Manual, "manual", -4135);
    }

    [Fact]
    public void SetMode_ToSemiAutomatic_Succeeds()
    {
        AssertModeRoundTrip(CalculationMode.SemiAutomatic, "semi-automatic", 2);
    }

    [Theory]
    [InlineData(CalculationScope.Workbook, 9d, 6d, 9d)]
    [InlineData(CalculationScope.Sheet, 9d, 4d, 6d)]
    [InlineData(CalculationScope.Range, 6d, 4d, 6d)]
    public void Calculate_UpdatesOnlyTheRequestedIndependentScope(
        CalculationScope scope, double targetOtherCell, double otherSheetCell, double otherSheetOtherCell)
    {
        var previous = _calculation.GetMode(_fixture.BatchToken);
        Assert.True(previous.Success, previous.ErrorMessage);
        var range = _fixture.CreateCommands<IRangeCommands>();
        var failure = Record.Exception(() =>
        {
            var manual = _calculation.SetMode(_fixture.BatchToken, CalculationMode.Manual);
            Assert.True(manual.Success, manual.ErrorMessage);
            var target = _fixture.CreateTestSheet(_fixture.BatchToken);
            var other = _fixture.CreateTestSheet(_fixture.BatchToken);
            foreach (var sheet in new[] { target, other })
            {
                var input = range.SetValues(_fixture.BatchToken, sheet, "A1", [[2]]);
                Assert.True(input.Success, input.ErrorMessage);
                var first = range.SetFormulas(_fixture.BatchToken, sheet, "B1", [["=A1*2"]]);
                Assert.True(first.Success, first.ErrorMessage);
                var second = range.SetFormulas(_fixture.BatchToken, sheet, "D1", [["=A1*3"]]);
                Assert.True(second.Success, second.ErrorMessage);
            }
            var baseline = _calculation.Calculate(_fixture.BatchToken, CalculationScope.Workbook);
            Assert.True(baseline.Success, baseline.ErrorMessage);
            foreach (var sheet in new[] { target, other })
            {
                AssertCells(sheet, 4d, 6d);
                var change = range.SetValues(_fixture.BatchToken, sheet, "A1", [[3]], overwritePolicy: OverwritePolicy.Allow);
                Assert.True(change.Success, change.ErrorMessage);
                AssertCells(sheet, 4d, 6d);
            }

            var calculated = _calculation.Calculate(_fixture.BatchToken, scope,
                scope == CalculationScope.Workbook ? null : target,
                scope == CalculationScope.Range ? "B1" : null);
            Assert.True(calculated.Success, calculated.ErrorMessage);
            AssertCells(target, 6d, targetOtherCell);
            AssertCells(other, otherSheetCell, otherSheetOtherCell);

            void AssertCells(string sheet, double first, double second)
            {
                var values = range.GetValues(_fixture.BatchToken, sheet, "B1:D1");
                Assert.True(values.Success, values.ErrorMessage);
                Assert.Equal(first, Convert.ToDouble(values.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
                Assert.Equal(second, Convert.ToDouble(values.Values[0][2], System.Globalization.CultureInfo.InvariantCulture));
            }
        });
        var restoration = Record.Exception(() =>
        {
            var restored = _calculation.SetMode(_fixture.BatchToken, (CalculationMode)previous.ModeValue);
            Assert.True(restored.Success, restored.ErrorMessage);
        });
        if (restoration != null)
            failure = PersistentServiceCleanupFailures.Combine(failure, restoration);
        if (failure != null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [Fact]
    public void Calculate_SheetScope_RequiresSheetName()
    {
        var result = _calculation.Calculate(
            _fixture.BatchToken,
            CalculationScope.Sheet,
            null);

        Assert.False(result.Success);
        Assert.Contains("sheetName is required", result.ErrorMessage ?? "");
    }

    [Fact]
    public void Calculate_RangeScope_RequiresBothSheetAndRange()
    {
        var result = _calculation.Calculate(
            _fixture.BatchToken,
            CalculationScope.Range,
            "Sheet1",
            null);

        Assert.False(result.Success);
        Assert.Contains("rangeAddress are required", result.ErrorMessage ?? "");
    }

    [Fact]
    public void Calculate_MissingSheet_PropagatesUsefulPublicError()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _calculation.Calculate(
                _fixture.BatchToken,
                CalculationScope.Sheet,
                "MissingSheet"));

        Assert.Contains("ComInterop/", exception.Message);
        Assert.DoesNotContain(
            "Calculation failed",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private void AssertModeRoundTrip(
        CalculationMode mode,
        string expectedName,
        int expectedValue)
    {
        try
        {
            var setResult = _calculation.SetMode(_fixture.BatchToken, mode);
            Assert.True(setResult.Success, setResult.ErrorMessage);

            var getResult = _calculation.GetMode(_fixture.BatchToken);
            Assert.Equal(expectedName, getResult.Mode);
            Assert.Equal(expectedValue, getResult.ModeValue);
        }
        finally
        {
            _calculation.SetMode(_fixture.BatchToken, CalculationMode.Automatic);
        }
    }
}
