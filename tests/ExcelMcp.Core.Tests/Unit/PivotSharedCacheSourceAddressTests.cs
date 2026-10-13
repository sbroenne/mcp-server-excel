using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PivotTables")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class PivotSharedCacheSourceAddressTests
{
    [Theory]
    [InlineData("R1C1")]
    [InlineData("R1C1:R6C4")]
    [InlineData("r1c1:r6c4")]
    [InlineData("C1:C4")]
    [InlineData("R1:R6")]
    [InlineData("c1:c4")]
    [InlineData("r1:r6")]
    [InlineData("C3")]
    [InlineData("R6")]
    [InlineData("C16384")]
    [InlineData("R1048576")]
    public void NativeR1C1SourceAddress_RequiresConversionBeforeA1Resolution(string address) =>
        Assert.True(PivotTableCommands.IsNativeR1C1SourceAddress(address));

    [Theory]
    [InlineData("A1:D6")]
    [InlineData("$C$1:$C$4")]
    [InlineData("$R$1:$R$6")]
    [InlineData("A:D")]
    [InlineData("$A:$D")]
    [InlineData("1:6")]
    [InlineData("$1:$6")]
    [InlineData("SharedSales")]
    [InlineData("SharedSales[#All]")]
    [InlineData("R[1]C[1]:R[6]C[4]")]
    [InlineData("R1:C4")]
    [InlineData("")]
    public void NonNativeR1C1SourceAddress_DoesNotUseNativeConversion(string address) =>
        Assert.False(PivotTableCommands.IsNativeR1C1SourceAddress(address));
}
