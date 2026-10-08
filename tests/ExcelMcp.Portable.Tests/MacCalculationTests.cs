using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacCalculationTests
{
    [Theory]
    [InlineData(CalculationScope.Application, CalculationKind.Normal)]
    [InlineData(CalculationScope.Application, CalculationKind.Full)]
    [InlineData(CalculationScope.Application, CalculationKind.Rebuild)]
    public void ApplicationCalculationFailsBeforeAnyDispatch(CalculationScope scope, CalculationKind kind)
    {
        var dispatched = 0;
        using var batch = CreateBatch(() => dispatched++);
        var error = Assert.Throws<PlatformNotSupportedException>(() =>
        batch.Invoke<OperationResult>("calculationmode.calculate", new { scope, kind }));
        Assert.Contains("no calculation was attempted", error.Message, StringComparison.Ordinal);
        Assert.Equal(0, dispatched);
    }

    [Fact]
    public void AcceptedScopedCalculationIsAvailableWithoutHelper()
    {
        var capability = MacCommandCapabilities.Get("calculationmode.calculate");
        Assert.True(capability.IsAvailable);
        Assert.Equal(MacCapabilityTier.Native, capability.RequiredTier);
        Assert.Equal("Implemented", capability.ImplementationStatus);
        Assert.Empty(capability.UnavailableMessage);
    }

    [Theory]
    [InlineData(CalculationScope.Sheet, null, null, CalculationKind.Normal)]
    [InlineData(CalculationScope.Range, "Sheet1", null, CalculationKind.Normal)]
    [InlineData(CalculationScope.Sheet, "Sheet1", "A1", CalculationKind.Normal)]
    [InlineData(CalculationScope.Application, "Sheet1", null, CalculationKind.Normal)]
    [InlineData(CalculationScope.Range, "Sheet1", "A1", CalculationKind.Full)]
    [InlineData((CalculationScope)99, null, null, CalculationKind.Normal)]
    [InlineData(CalculationScope.Sheet, "Sheet1", null, (CalculationKind)99)]
    public void InvalidScopeInputsUseSharedValidation(
        CalculationScope scope, string? sheetName, string? rangeAddress, CalculationKind kind)
    {
        Assert.ThrowsAny<ArgumentException>(() =>
            CalculationCommandValidation.Validate(scope, sheetName, rangeAddress, kind));
    }

    private static MacExcelBatch CreateBatch(Action dispatched) => new(new MacExcelBackend((_, _, _) =>
    {
        dispatched();
        return Task.FromResult(new MacProcessResult(0, """{"success":true,"errorMessage":""}""", ""));
    }), new MacExcelSession
    {
        SessionId = "calculation-contract",
        FilePath = Path.GetFullPath("opaque.xlsx"),
        IsVisible = false,
        OperationTimeout = TimeSpan.FromSeconds(10)
    });
}
