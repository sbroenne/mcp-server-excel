using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelCommandsTests
{
    [Fact]
    public void CreateMeasure_WithMixedCaseWholeNumber_PreservesFormat()
    {
        var measureName = $"Test_{nameof(CreateMeasure_WithMixedCaseWholeNumber_PreservesFormat)}_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var createResult = _dataModelCommands.CreateMeasure(
            batch,
            "SalesTable",
            measureName,
            "SUM(SalesTable[Amount])",
            formatType: "wHoLeNuMbEr");
        var readResult = _dataModelCommands.Read(batch, measureName);

        Assert.True(createResult.Success, $"CreateMeasure failed: {createResult.ErrorMessage}");
        Assert.True(readResult.Success, $"Read failed: {readResult.ErrorMessage}");
        Assert.NotNull(readResult.FormatInfo);
        Assert.Equal("WholeNumber", readResult.FormatInfo.Type);
        Assert.Equal(0, readResult.FormatInfo.DecimalPlaces);
    }

    [Fact]
    public void UpdateMeasure_WithMixedCaseWholeNumber_PreservesFormat()
    {
        var measureName = $"Test_{nameof(UpdateMeasure_WithMixedCaseWholeNumber_PreservesFormat)}_{Guid.NewGuid():N}";

        using var batch = ExcelSession.BeginBatch(_dataModelFile);

        var createResult = _dataModelCommands.CreateMeasure(
            batch,
            "SalesTable",
            measureName,
            "SUM(SalesTable[Amount])",
            formatType: "Decimal");
        var updateResult = _dataModelCommands.UpdateMeasure(
            batch,
            measureName,
            formatType: "WHOLEnumber");
        var readResult = _dataModelCommands.Read(batch, measureName);

        Assert.True(createResult.Success, $"CreateMeasure failed: {createResult.ErrorMessage}");
        Assert.True(updateResult.Success, $"UpdateMeasure failed: {updateResult.ErrorMessage}");
        Assert.True(readResult.Success, $"Read failed: {readResult.ErrorMessage}");
        Assert.NotNull(readResult.FormatInfo);
        Assert.Equal("WholeNumber", readResult.FormatInfo.Type);
        Assert.Equal(0, readResult.FormatInfo.DecimalPlaces);
    }
}
