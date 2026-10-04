using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for native DAX syntax, formula preservation, and evaluation.
/// </summary>
/// <remarks>
/// When Windows uses a comma as the decimal mark, Excel reads a comma that touches a number
/// as a decimal point, so ExcelMcp stores a space between them and says so in Message (#978).
/// </remarks>
public partial class PersistentServiceDataModelCommandsTests
{
    private const string DecimalCommaSpacingNote = "Spaces were added next to commas that touch a number";

    private static bool WindowsUsesDecimalComma =>
        new CultureInfo(CultureInfo.CurrentCulture.Name, useUserOverride: true)
            .NumberFormat.NumberDecimalSeparator == ",";

    private static string ExpectedStoredFormula(string formula, string decimalCommaFormula) =>
        WindowsUsesDecimalComma ? decimalCommaFormula : formula;

    private static void AssertSpacingNote(OperationResult result, string formula, string decimalCommaFormula)
    {
        RequireSuccess(result);
        if (WindowsUsesDecimalComma && formula != decimalCommaFormula)
        {
            Assert.NotNull(result.Message);
            Assert.StartsWith(DecimalCommaSpacingNote, result.Message, StringComparison.Ordinal);
        }
        else
        {
            Assert.Null(result.Message);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteMeasure_NumberBeforeComma_StoresWorkingFormulaAndReportsSpacing(bool update)
    {
        var batch = _fixture.BatchToken;
        var measureName = $"Test_NumericComma_{Guid.NewGuid():N}";
        const string formula = "IF(1=1, ROUND(1.25,1), 0)";
        const string decimalCommaFormula = "IF(1=1 , ROUND(1.25 , 1), 0)";

        var created = CreateMeasure("SalesTable", measureName, update ? "0" : formula);
        Assert.True(created.Success, created.ErrorMessage);
        var written = created;
        if (update)
        {
            Assert.Null(created.Message);
            written = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: formula);
        }

        AssertSpacingNote(written, formula, decimalCommaFormula);
        var read = _dataModelCommands.Read(batch, measureName);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), read.DaxFormula);

        var evaluated = _dataModelCommands.Evaluate(batch, $"EVALUATE ROW(\"Result\", [{measureName}])");
        Assert.True(evaluated.Success, evaluated.ErrorMessage);
        Assert.Equal(1.3m, Convert.ToDecimal(Assert.Single(Assert.Single(evaluated.Rows)),
            CultureInfo.InvariantCulture));
    }

    [Fact]
    public async Task Read_AfterSaveAndReopen_ReturnsStoredDaxWithDecimalPoint()
    {
        var measureName = $"Test_ReopenDecimal_{Guid.NewGuid():N}";
        const string formula = "IF(1=1, 1.5, 0)";
        const string decimalCommaFormula = "IF(1=1 , 1.5 , 0)";
        var created = CreateMeasure("SalesTable", measureName, formula);
        AssertSpacingNote(created, formula, decimalCommaFormula);

        await _fixture.SaveAndReopenAsync();

        var read = _dataModelCommands.Read(_fixture.BatchToken, measureName);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), read.DaxFormula);

        var listed = _dataModelCommands.ListMeasures(_fixture.BatchToken, "SalesTable");
        Assert.True(listed.Success, listed.ErrorMessage);
        var info = Assert.Single(listed.Measures, m => m.Name == measureName);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), info.FormulaPreview);

        var evaluated = _dataModelCommands.Evaluate(_fixture.BatchToken,
            $"EVALUATE ROW(\"Result\", [{measureName}])");
        Assert.True(evaluated.Success, evaluated.ErrorMessage);
        Assert.Equal(1.5m, Convert.ToDecimal(Assert.Single(Assert.Single(evaluated.Rows)),
            CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false, "IF(TRUE(), -1E3, 0)", "IF(TRUE(), -1E3 , 0)", -1000)]
    [InlineData(true, "IF(TRUE(), -1E3, 0)", "IF(TRUE(), -1E3 , 0)", -1000)]
    [InlineData(false, "MAX(1e3,2E+2)", "MAX(1e3 , 2E+2)", 1000)]
    [InlineData(true, "MAX(1e3,2E+2)", "MAX(1e3 , 2E+2)", 1000)]
    [InlineData(false, "ROUND(1.25e-3,4)", "ROUND(1.25e-3 , 4)", 0.0013)]
    [InlineData(true, "ROUND(1.25e-3,4)", "ROUND(1.25e-3 , 4)", 0.0013)]
    public void WriteMeasure_ScientificNotation_PreservesFormulaAndEvaluates(
        bool update, string formula, string decimalCommaFormula, double expected)
    {
        var batch = _fixture.BatchToken;
        var measureName = $"Test_ExponentComma_{Guid.NewGuid():N}";
        var created = CreateMeasure("SalesTable", measureName, update ? "0" : formula);
        Assert.True(created.Success, created.ErrorMessage);
        var written = created;
        if (update)
        {
            Assert.Null(created.Message);
            written = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: formula);
        }

        AssertSpacingNote(written, formula, decimalCommaFormula);
        var read = _dataModelCommands.Read(batch, measureName);
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), read.DaxFormula);

        var listed = _dataModelCommands.ListMeasures(batch, "SalesTable");
        Assert.True(listed.Success, listed.ErrorMessage);
        var info = Assert.Single(listed.Measures, m => m.Name == measureName);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), info.FormulaPreview);

        var evaluated = _dataModelCommands.Evaluate(batch, $"EVALUATE ROW(\"Result\", [{measureName}])");
        Assert.True(evaluated.Success, evaluated.ErrorMessage);
        Assert.Equal(expected, Convert.ToDouble(Assert.Single(Assert.Single(evaluated.Rows)),
            CultureInfo.InvariantCulture), precision: 10);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Evaluate_InvalidStoredFormula_PreservesMeasureDefinitions(bool update)
    {
        var batch = _fixture.BatchToken;
        var name = $"Test_RetainedSyntax_{Guid.NewGuid():N}";
        const string formula = "IF(TRUE(), 1.5, 0)";
        RequireSuccess(CreateMeasure("SalesTable", name, formula,
            formatType: "Decimal", description: "Retained measure"));
        var rejectedName = $"Test_InvalidSyntax_{Guid.NewGuid():N}";
        const string invalidFormula = "NOT_A_DAX_FUNCTION(1.5, 0)";
        var expectedInvalidFormula = ExpectedStoredFormula(invalidFormula, "NOT_A_DAX_FUNCTION(1.5 , 0)");
        RequireSuccess(CreateMeasure("SalesTable", rejectedName, update ? "0" : invalidFormula));
        if (update)
        {
            RequireSuccess(_dataModelCommands.UpdateMeasure(batch, rejectedName, daxFormula: invalidFormula));
        }
        Assert.Equal(expectedInvalidFormula, RequireSuccess(_dataModelCommands.Read(batch, rejectedName)).DaxFormula);
        var before = RequireSuccess(_dataModelCommands.Read(batch, name));
        var listed = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_dataModelCommands.ListMeasures(batch)).Measures);
        var failure = Assert.Throws<InvalidOperationException>(() => _dataModelCommands.Evaluate(
            batch, $"EVALUATE ROW(\"Rejected\", [{rejectedName}])"));
        Assert.Contains("DAX evaluation failed", failure.Message);
        var after = RequireSuccess(_dataModelCommands.Read(batch, name));
        Assert.Equal(before.DaxFormula, after.DaxFormula);
        Assert.Equal(before.Description, after.Description);
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before.FormatInfo),
            System.Text.Json.JsonSerializer.Serialize(after.FormatInfo));
        Assert.Equal(before.MeasureName, after.MeasureName);
        Assert.Equal(before.TableName, after.TableName);
        Assert.Equal(before.CharacterCount, after.CharacterCount);
        Assert.Equal(listed, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_dataModelCommands.ListMeasures(batch)).Measures));
        Assert.Equal(expectedInvalidFormula, RequireSuccess(_dataModelCommands.Read(batch, rejectedName)).DaxFormula);
        var evaluated = RequireSuccess(_dataModelCommands.Evaluate(batch,
            $"EVALUATE ROW(\"Result\", [{name}])"));
        Assert.Equal(1.5, Convert.ToDouble(Assert.Single(Assert.Single(evaluated.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("SUMX({1.5, 2.5}, [Value])", "SUMX({1.5 , 2.5}, [Value])", 4)]
    [InlineData("VAR rate = 1.5 /* literal 9.9, */ RETURN IF(rate > 0, rate, 0)",
        "VAR rate = 1.5 /* literal 9.9, */ RETURN IF(rate > 0 , rate, 0)", 1.5)]
    [InlineData("VAR rate = 1.5 // literal 9.9,\nRETURN IF(rate > 0, rate, 0)",
        "VAR rate = 1.5 // literal 9.9,\nRETURN IF(rate > 0 , rate, 0)", 1.5)]
    [InlineData("IF(TRUE(), 1.5E-2, 0)", "IF(TRUE(), 1.5E-2 , 0)", 0.015)]
    public void WriteMeasure_NumericTokensAndComments_PreservesAndEvaluates(
        string formula, string decimalCommaFormula, double expected)
    {
        var name = $"Test_Syntax_{Guid.NewGuid():N}";
        AssertSpacingNote(CreateMeasure("SalesTable", name, formula), formula, decimalCommaFormula);
        var expectedFormula = ExpectedStoredFormula(formula, decimalCommaFormula);
        Assert.Equal(expectedFormula, RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name)).DaxFormula);
        RequireSuccess(_dataModelCommands.UpdateMeasure(_fixture.BatchToken, name, daxFormula: "0"));
        Assert.Equal("0", RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name)).DaxFormula);
        AssertSpacingNote(_dataModelCommands.UpdateMeasure(_fixture.BatchToken, name, daxFormula: formula),
            formula, decimalCommaFormula);
        Assert.Equal(expectedFormula, RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name)).DaxFormula);
        var result = RequireSuccess(_dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Result\", [{name}])"));
        Assert.Equal(expected, Convert.ToDouble(Assert.Single(Assert.Single(result.Rows)),
            System.Globalization.CultureInfo.InvariantCulture), precision: 12);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteMeasure_NativeSpacedSyntax_PreservesAndEvaluates(bool quoted)
    {
        var name = $"Test_NativeCulture_{Guid.NewGuid():N}";
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Model? model = null;
            Microsoft.Office.Interop.Excel.ModelTables? tables = null;
            Microsoft.Office.Interop.Excel.ModelTable? table = null;
            Microsoft.Office.Interop.Excel.ModelMeasures? measures = null;
            Microsoft.Office.Interop.Excel.ModelMeasure? measure = null;
            Microsoft.Office.Interop.Excel.ModelFormatGeneral? format = null;
            try
            {
                var condition = quoted ? "\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\"" : "1=1";
                var formula = $"IF({condition} , 1.5 , 0)";
                model = context.Book.Model;
                tables = model.ModelTables;
                table = tables.Item("SalesTable");
                measures = model.ModelMeasures;
                format = model.ModelFormatGeneral;
                measure = measures.Add(name, table, formula, format, Type.Missing);
                _fixture.RegisterDataModelMeasureForCleanup(name);
                Assert.Equal(formula, measure.Formula);
                measure.Formula = "0";
                measure.Formula = formula;
                Assert.Equal(formula, measure.Formula);
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref measure);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref format);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref measures);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref table);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref tables);
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref model);
            }
        });
        var expectedCondition = quoted ? "\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\"" : "1=1";
        Assert.Equal($"IF({expectedCondition} , 1.5 , 0)",
            RequireSuccess(_dataModelCommands.Read(_fixture.BatchToken, name)).DaxFormula);
        var evaluated = RequireSuccess(_dataModelCommands.Evaluate(
            _fixture.BatchToken, $"EVALUATE ROW(\"Result\", [{name}])"));
        Assert.Equal(1.5, Convert.ToDouble(Assert.Single(Assert.Single(evaluated.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteMeasure_CommaArguments_PreservesFormulaAndEvaluates(bool update)
    {
        var batch = _fixture.BatchToken;
        var baseline = RequireSuccess(_dataModelCommands.Evaluate(batch,
            "EVALUATE ROW(\"Total\", SUM(SalesTable[Amount]))"));
        var total = Convert.ToDecimal(Assert.Single(Assert.Single(baseline.Rows)),
            System.Globalization.CultureInfo.InvariantCulture);

        var measureName = $"Test_CommaArguments_{Guid.NewGuid():N}";
        const string formula = "DIVIDE(SUM(SalesTable[Amount]), 1000)";
        var created = CreateMeasure("SalesTable", measureName,
            update ? "SUM(SalesTable[Amount])" : formula);
        RequireSuccess(created);
        Assert.Null(created.Message);
        if (update)
        {
            var updated = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: formula);
            RequireSuccess(updated);
            Assert.Null(updated.Message);
        }

        var read = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(read);
        Assert.Equal(formula, read.DaxFormula);

        var evaluated = _dataModelCommands.Evaluate(batch,
            $"EVALUATE ROW(\"Result\", [{measureName}])");
        RequireSuccess(evaluated);
        Assert.Equal(total / 1000m,
            Convert.ToDecimal(Assert.Single(Assert.Single(evaluated.Rows)),
                System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WriteMeasure_QuotedCommasEscapedQuotesAndDecimalPoint_PreservesAndEvaluates(bool update)
    {
        var batch = _fixture.BatchToken;
        var measureName = $"Test_Literals_{Guid.NewGuid():N}";
        const string formula = "IF(\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\", 1.5, 0)";
        const string decimalCommaFormula = "IF(\"North, \"\"South\"\"\" = \"North, \"\"South\"\"\", 1.5 , 0)";
        var baseline = _dataModelCommands.Evaluate(batch, $"EVALUATE ROW(\"Expected\", {formula})");
        RequireSuccess(baseline);
        Assert.Equal(1.5m, Convert.ToDecimal(Assert.Single(Assert.Single(baseline.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
        var created = CreateMeasure("SalesTable", measureName, update ? "0" : formula);
        RequireSuccess(created);
        var written = created;
        if (update)
        {
            written = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: formula);
        }
        AssertSpacingNote(written, formula, decimalCommaFormula);
        var read = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(read);
        Assert.Equal(ExpectedStoredFormula(formula, decimalCommaFormula), read.DaxFormula);
        var evaluated = _dataModelCommands.Evaluate(batch, $"EVALUATE ROW(\"Result\", [{measureName}])");
        RequireSuccess(evaluated);
        Assert.Equal(1.5m, Convert.ToDecimal(Assert.Single(Assert.Single(evaluated.Rows)),
            System.Globalization.CultureInfo.InvariantCulture));
    }

    #region Native DAX Tests

    /// <summary>
    /// Tests that DAX formulas with function argument separators (commas) are handled correctly.
    /// This is the exact formula from the user's bug report where DATEADD arguments were corrupted.
    /// LLM use case: "create a measure with DATEADD function"
    /// </summary>
    [Fact]
    public void CreateMeasure_DateAddFormula_CreatesSuccessfully()
    {
        // This is the formula that was failing - comma was becoming period on European locales
        var measureName = $"Test_DATEADD_{Guid.NewGuid():N}";
        var daxFormula = "CALCULATE([Total Sales], DATEADD(SalesTable[Date], -1, MONTH))";
        const string decimalCommaFormula = "CALCULATE([Total Sales], DATEADD(SalesTable[Date], -1 , MONTH))";

        var batch = _fixture.BatchToken;

        var created = CreateMeasure("SalesTable", measureName, daxFormula);
        AssertSpacingNote(created, daxFormula, decimalCommaFormula);

        // Verify measure was created
        var listResult = RequireSuccess(_dataModelCommands.ListMeasures(batch));
        Assert.Contains(listResult.Measures, m => m.Name == measureName);

        var readResult = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(readResult);
        Assert.NotNull(readResult.DaxFormula);
        Assert.Contains("DATEADD", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("CALCULATE", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(ExpectedStoredFormula(daxFormula, decimalCommaFormula), readResult.DaxFormula);
    }

    /// <summary>
    /// Tests nested function calls with multiple comma separators.
    /// LLM use case: "create a complex DAX measure with nested functions"
    /// </summary>
    [Fact]
    public void CreateMeasure_NestedFunctions_CreatesSuccessfully()
    {
        var measureName = $"Test_Nested_{Guid.NewGuid():N}";
        // Complex formula with multiple nested functions and comma separators
        var daxFormula = "CALCULATE(SUM(SalesTable[Amount]), FILTER(ALL(SalesTable), SalesTable[Amount] > 100))";

        var batch = _fixture.BatchToken;
        RequireSuccess(CreateMeasure("SalesTable", measureName, daxFormula));

        // Verify measure was created and formula is valid
        var readResult = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(readResult);
        Assert.Equal(daxFormula, readResult.DaxFormula);
    }

    /// <summary>
    /// Tests updating a measure with a complex DAX formula containing function separators.
    /// LLM use case: "update this measure's formula to use DATESINPERIOD"
    /// </summary>
    [Fact]
    public void UpdateMeasure_ComplexDaxFormula_UpdatesSuccessfully()
    {
        var measureName = $"Test_Update_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        // Rolling 3-month formula with multiple comma separators
        var updatedFormula = "AVERAGEX(DATESINPERIOD(SalesTable[Date], MAX(SalesTable[Date]), -3, MONTH), SalesTable[Amount])";
        const string decimalCommaFormula = "AVERAGEX(DATESINPERIOD(SalesTable[Date], MAX(SalesTable[Date]), -3 , MONTH), SalesTable[Amount])";

        var batch = _fixture.BatchToken;

        // Create measure with simple formula
        var created = CreateMeasure("SalesTable", measureName, originalFormula);
        RequireSuccess(created);
        Assert.Equal(originalFormula, RequireSuccess(_dataModelCommands.Read(batch, measureName)).DaxFormula);

        var updated = _dataModelCommands.UpdateMeasure(batch, measureName, daxFormula: updatedFormula);
        AssertSpacingNote(updated, updatedFormula, decimalCommaFormula);

        // Verify the formula was updated
        var readResult = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(readResult);
        Assert.Contains("AVERAGEX", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("DATESINPERIOD", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(ExpectedStoredFormula(updatedFormula, decimalCommaFormula), readResult.DaxFormula);
    }

    /// <summary>
    /// Tests DAX formula with string literals containing commas - commas inside strings should NOT be translated.
    /// LLM use case: "create a measure that checks for a specific text value"
    /// </summary>
    [Fact]
    public void CreateMeasure_StringLiteralWithComma_PreservesStringContent()
    {
        var measureName = $"Test_String_{Guid.NewGuid():N}";
        // Formula with comma inside a string literal - this comma should NOT be translated
        var daxFormula = "IF(MAX(SalesTable[Region]) = \"North, South\", 1, 0)";
        const string decimalCommaFormula = "IF(MAX(SalesTable[Region]) = \"North, South\", 1 , 0)";

        var batch = _fixture.BatchToken;
        var created = CreateMeasure("SalesTable", measureName, daxFormula);
        AssertSpacingNote(created, daxFormula, decimalCommaFormula);

        // Verify measure was created
        var readResult = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(readResult);
        Assert.Contains("IF", readResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(ExpectedStoredFormula(daxFormula, decimalCommaFormula), readResult.DaxFormula);
    }

    /// <summary>
    /// Tests simple DAX formula without function separators - should work unchanged.
    /// LLM use case: "create a simple SUM measure"
    /// </summary>
    [Fact]
    public void CreateMeasure_SimpleFormula_CreatesSuccessfully()
    {
        var measureName = $"Test_Simple_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        var batch = _fixture.BatchToken;
        RequireSuccess(CreateMeasure("SalesTable", measureName, daxFormula));

        var readResult = _dataModelCommands.Read(batch, measureName);
        RequireSuccess(readResult);
        Assert.Equal(daxFormula, readResult.DaxFormula);
    }

    #endregion
}
