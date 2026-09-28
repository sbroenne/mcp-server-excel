using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public void ListCalculatedMembers_NonOlapPivotTable_ReturnsError()
    {
        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            _salesSheetName,
            "F2",
            "RegularPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var result = _pivotCommands.ListCalculatedMembers(
            batch,
            "RegularPivot");

        Assert.False(result.Success);
        Assert.Contains("not an OLAP PivotTable", result.ErrorMessage);
        Assert.Contains("create-calculated-field", result.ErrorMessage);
    }

    [Fact]
    public void CreateCalculatedMember_NonOlapPivotTable_ReturnsError()
    {
        var batch = _fixture.BatchToken;
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            _salesSheetName,
            "F2",
            "RegularPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var result = _pivotCommands.CreateCalculatedMember(
            batch,
            "RegularPivot",
            "[Measures].[ShouldFail]",
            "[Measures].[Something]",
            CalculatedMemberType.Measure);

        Assert.False(result.Success);
        Assert.Contains("not an OLAP PivotTable", result.ErrorMessage);
    }
}
