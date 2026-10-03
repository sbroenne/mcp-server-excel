using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Theory]
    [InlineData("General")]
    [InlineData("Currency")]
    [InlineData("Decimal")]
    [InlineData("Percentage")]
    [InlineData("WholeNumber")]
    public void CreateMeasure_ExplicitSupportedFormat_ReadsBackRequestedType(string format)
    {
        var name = $"RequestedFormat_{Guid.NewGuid():N}";
        Assert.True(CreateMeasure("SalesTable", name, "SUM(SalesTable[Amount])", format).Success);
        var read = _dataModelCommands.Read(_fixture.BatchToken, name);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.NotNull(read.FormatInfo);
        Assert.Equal(format, read.FormatInfo.Type);
    }

    [Fact]
    public void ReadMeasure_CurrencyWithEmptySymbol_PreservesCurrencyType()
    {
        var name = $"EmptySymbol_{Guid.NewGuid():N}";
        Assert.True(CreateMeasure("SalesTable", name, "SUM(SalesTable[Amount])", "Currency").Success);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Model? model = null;
            Excel.ModelMeasures? measures = null;
            Excel.ModelMeasure? measure = null;
            object? format = null;
            try
            {
                model = ctx.Book.Model;
                measures = model.ModelMeasures;
                for (int i = 1; i <= measures.Count; i++)
                {
                    measure = measures.Item(i);
                    if (measure.Name == name)
                    {
                        format = measure.FormatInformation;
                        var currency = (Excel.ModelFormatCurrency)format;
                        currency.Symbol = "";
                        measure.FormatInformation = format;
                        return;
                    }
                    ComUtilities.Release(ref measure);
                }
                Assert.Fail("Created measure was not found in Excel.");
            }
            finally
            {
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref measures);
                ComUtilities.Release(ref model);
            }
        });
        var read = _dataModelCommands.Read(_fixture.BatchToken, name);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.NotNull(read.FormatInfo);
        Assert.Equal("Currency", read.FormatInfo.Type);
        Assert.Equal("", read.FormatInfo.Symbol);
    }

    [Fact]
    public void MeasureFormat_OmittedCreateDefaultsAndEmptyUpdatePreservesExistingFormat()
    {
        var general = $"DefaultFormat_{Guid.NewGuid():N}";
        Assert.True(CreateMeasure("SalesTable", general, "SUM(SalesTable[Amount])").Success);
        var generalRead = _dataModelCommands.Read(_fixture.BatchToken, general);
        Assert.True(generalRead.Success, generalRead.ErrorMessage);
        Assert.NotNull(generalRead.FormatInfo);
        Assert.Equal("General", generalRead.FormatInfo.Type);

        var currency = $"PreservedFormat_{Guid.NewGuid():N}";
        Assert.True(CreateMeasure("SalesTable", currency, "SUM(SalesTable[Amount])", "Currency").Success);
        Assert.True(_dataModelCommands.UpdateMeasure(
            _fixture.BatchToken, currency, formatType: "", description: "Updated description").Success);
        var updated = _dataModelCommands.Read(_fixture.BatchToken, currency);
        Assert.True(updated.Success, updated.ErrorMessage);
        Assert.NotNull(updated.FormatInfo);
        Assert.Equal("Currency", updated.FormatInfo.Type);
        Assert.Equal("Updated description", updated.Description);
    }

    [Fact]
    public void CreateMeasure_WithMixedCaseWholeNumber_PreservesFormat()
    {
        var measureName = $"Test_{nameof(CreateMeasure_WithMixedCaseWholeNumber_PreservesFormat)}_{Guid.NewGuid():N}";

        var batch = _fixture.BatchToken;

        var createResult = CreateMeasure(
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

        var batch = _fixture.BatchToken;

        var createResult = CreateMeasure(
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
