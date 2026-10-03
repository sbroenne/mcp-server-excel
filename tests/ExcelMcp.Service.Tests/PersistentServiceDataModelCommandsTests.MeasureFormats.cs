using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Theory]
    [InlineData("General", false)]
    [InlineData("Currency", false)]
    [InlineData("Decimal", false)]
    [InlineData("Percentage", false)]
    [InlineData("WholeNumber", false)]
    [InlineData("General", true)]
    [InlineData("Currency", true)]
    [InlineData("Decimal", true)]
    [InlineData("Percentage", true)]
    [InlineData("WholeNumber", true)]
    public void WriteMeasure_FormatType_ReturnsExactNativeMetadata(string formatType, bool update)
    {
        var name = $"Test_NativeFormat_{Guid.NewGuid():N}";
        var initialFormat = formatType == "General" ? "Percentage" : "General";
        RequireSuccess(CreateMeasure("SalesTable", name, "1234.5",
            formatType: update ? initialFormat : formatType));
        if (update)
        {
            var before = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name));
            Assert.Equal(initialFormat, Assert.IsType<Sbroenne.ExcelMcp.Core.Models.MeasureFormatInfo>(before.FormatInfo).Type);
            RequireSuccess(_dataModelCommands.UpdateMeasure(
                _fixture.BatchToken, name, formatType: formatType));
        }

        var read = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name));
        Assert.Equal(name, read.MeasureName);
        Assert.Equal("SalesTable", read.TableName);
        Assert.Equal("1234.5", read.DaxFormula);
        var format = Assert.IsType<Sbroenne.ExcelMcp.Core.Models.MeasureFormatInfo>(read.FormatInfo);
        Assert.Equal(formatType, format.Type);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Model? model = null;
            Excel.ModelMeasures? measures = null;
            Excel.ModelMeasure? measure = null;
            object? nativeFormat = null;
            try
            {
                model = context.Book.Model;
                measures = model.ModelMeasures;
                measure = measures.Item(name);
                nativeFormat = measure.FormatInformation;
                switch (formatType)
                {
                    case "General":
                        Assert.True(nativeFormat is Excel.ModelFormatGeneral);
                        Assert.Null(format.DecimalPlaces);
                        Assert.Null(format.UseThousandSeparator);
                        Assert.Null(format.Symbol);
                        break;
                    case "Currency":
                        var currency = nativeFormat as Excel.ModelFormatCurrency;
                        Assert.NotNull(currency);
                        Assert.Equal(currency.Symbol, format.Symbol);
                        Assert.Equal(currency.DecimalPlaces, format.DecimalPlaces);
                        Assert.Null(format.UseThousandSeparator);
                        break;
                    case "Decimal":
                        var number = nativeFormat as Excel.ModelFormatDecimalNumber;
                        Assert.NotNull(number);
                        Assert.Equal(number.DecimalPlaces, format.DecimalPlaces);
                        Assert.Equal(number.UseThousandSeparator, format.UseThousandSeparator);
                        Assert.Null(format.Symbol);
                        break;
                    case "Percentage":
                        var percentage = nativeFormat as Excel.ModelFormatPercentageNumber;
                        Assert.NotNull(percentage);
                        Assert.Equal(percentage.DecimalPlaces, format.DecimalPlaces);
                        Assert.Equal(percentage.UseThousandSeparator, format.UseThousandSeparator);
                        Assert.Null(format.Symbol);
                        break;
                    case "WholeNumber":
                        var whole = nativeFormat as Excel.ModelFormatWholeNumber;
                        Assert.NotNull(whole);
                        Assert.Equal(0, format.DecimalPlaces);
                        Assert.Equal(whole.UseThousandSeparator, format.UseThousandSeparator);
                        Assert.Null(format.Symbol);
                        break;
                }
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref nativeFormat);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref measure);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref measures);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref model);
            }
        });
        var evaluated = RequireSuccess(_dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Result\", [{name}])"));
        Assert.Equal(1234.5, Convert.ToDouble(Assert.Single(Assert.Single(evaluated.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("General")]
    [InlineData("Currency")]
    [InlineData("Decimal")]
    [InlineData("Percentage")]
    [InlineData("WholeNumber")]
    public void CreateMeasure_ExplicitSupportedFormat_ReadsBackRequestedType(string format)
    {
        var name = $"RequestedFormat_{Guid.NewGuid():N}";
        RequireSuccess(CreateMeasure("SalesTable", name, "SUM(SalesTable[Amount])", format));
        var read = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name));
        Assert.NotNull(read.FormatInfo);
        Assert.Equal(format, read.FormatInfo.Type);
    }

    [Fact]
    public void ReadMeasure_CurrencyWithEmptySymbol_PreservesCurrencyType()
    {
        var name = $"EmptySymbol_{Guid.NewGuid():N}";
        RequireSuccess(CreateMeasure("SalesTable", name, "SUM(SalesTable[Amount])", "Currency"));
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
                measure = measures.Item(name);
                format = measure.FormatInformation;
                var currency = (Excel.ModelFormatCurrency)format;
                currency.Symbol = "";
                measure.FormatInformation = format;
            }
            finally
            {
                ComUtilities.Release(ref format);
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref measures);
                ComUtilities.Release(ref model);
            }
        });
        var read = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name));
        Assert.NotNull(read.FormatInfo);
        Assert.Equal("Currency", read.FormatInfo.Type);
        Assert.Equal("", read.FormatInfo.Symbol);
    }

    [Fact]
    public void MeasureFormat_OmittedCreateDefaultsAndEmptyUpdatePreservesExistingFormat()
    {
        var general = $"DefaultFormat_{Guid.NewGuid():N}";
        RequireSuccess(CreateMeasure("SalesTable", general, "SUM(SalesTable[Amount])"));
        var generalRead = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, general));
        Assert.NotNull(generalRead.FormatInfo);
        Assert.Equal("General", generalRead.FormatInfo.Type);

        var currency = $"PreservedFormat_{Guid.NewGuid():N}";
        RequireSuccess(CreateMeasure("SalesTable", currency, "SUM(SalesTable[Amount])", "Currency"));
        RequireSuccess(_dataModelCommands.UpdateMeasure(
            _fixture.BatchToken, currency, formatType: "", description: "Updated description"));
        var updated = RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, currency));
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
        RequireSuccess(createResult);
        var readResult = RequireSuccess(_dataModelCommands.Read(batch, measureName));
        Assert.Equal("SUM(SalesTable[Amount])", readResult.DaxFormula);
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
        RequireSuccess(createResult);
        var before = RequireSuccess(_dataModelCommands.Read(batch, measureName));
        Assert.Equal("Decimal", Assert.IsType<Sbroenne.ExcelMcp.Core.Models.MeasureFormatInfo>(before.FormatInfo).Type);
        RequireSuccess(_dataModelCommands.UpdateMeasure(
            batch,
            measureName,
            formatType: "WHOLEnumber"));
        var readResult = RequireSuccess(_dataModelCommands.Read(batch, measureName));
        Assert.Equal(before.DaxFormula, readResult.DaxFormula);
        Assert.Equal(before.Description, readResult.Description);
        Assert.NotNull(readResult.FormatInfo);
        Assert.Equal("WholeNumber", readResult.FormatInfo.Type);
        Assert.Equal(0, readResult.FormatInfo.DecimalPlaces);
    }
}
