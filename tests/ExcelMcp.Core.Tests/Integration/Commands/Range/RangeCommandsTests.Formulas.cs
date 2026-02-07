using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range formulas operations
/// </summary>
public partial class RangeCommandsTests
{
    // === FORMULA OPERATIONS TESTS ===

    [Fact]
    public void GetFormulas_ReturnsFormulasAndValues()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set values and formulas
        _commands.SetValues(batch, sheetName, "A1:A3",
        [
            [10],
            [20],
            [30]
        ]);

        _commands.SetFormulas(batch, sheetName, "B1",
        [
            ["=SUM(A1:A3)"]
        ]);

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "B1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal("=SUM(A1:A3)", result.Formulas[0][0]);
        Assert.Equal(
            60.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetFormulas_WritesFormulasToRange()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "A1:A3",
        [
            [5],
            [10],
            [15]
        ]);

        var formulas = new List<List<string>>
        {
            new() { "=A1*2", "=A2*2", "=A3*2" }
        };

        // Act
        var result = _commands.SetFormulas(batch, sheetName, "B1:D1", formulas);
        // Assert
        Assert.True(result.Success);

        // Verify values
        var readResult = _commands.GetValues(batch, sheetName, "B1:D1");
        Assert.Equal(
            10.0,
            Convert.ToDouble(readResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            20.0,
            Convert.ToDouble(readResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            30.0,
            Convert.ToDouble(readResult.Values[0][2], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetFormulas_WithJsonElementFormulas_WritesFormulasCorrectly()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up source data
        _commands.SetValues(batch, sheetName, "A1:A3",
        [
            [100],
            [200],
            [300]
        ]);

        // Simulate MCP framework JSON deserialization
        // MCP receives: {"formulas": [["=SUM(A1:A3)", "=AVERAGE(A1:A3)"]]}
        // Framework deserializes to List<List<string>> where each string is JsonElement
        string json = """[["=SUM(A1:A3)", "=AVERAGE(A1:A3)"]]""";
        var jsonDoc = System.Text.Json.JsonDocument.Parse(json);

        var testFormulas = new List<List<string>>();
        foreach (var rowElement in jsonDoc.RootElement.EnumerateArray())
        {
            var row = new List<string>();
            foreach (var cellElement in rowElement.EnumerateArray())
            {
                // This is JsonElement, not primitive string
                row.Add(cellElement.GetString() ?? "");
            }
            testFormulas.Add(row);
        }

        // Act - Should handle JsonElement conversion internally
        var result = _commands.SetFormulas(batch, sheetName, "B1:C1", testFormulas);
        // Assert
        Assert.True(result.Success, $"SetFormulas failed: {result.ErrorMessage}");

        // Verify formulas were written correctly
        var formulaResult = _commands.GetFormulas(batch, sheetName, "B1:C1");
        Assert.True(formulaResult.Success);
        Assert.Equal("=SUM(A1:A3)", formulaResult.Formulas[0][0]);
        Assert.Equal("=AVERAGE(A1:A3)", formulaResult.Formulas[0][1]);

        // Verify calculated values
        Assert.Equal(
            600.0,
            Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // SUM
        Assert.Equal(
            200.0,
            Convert.ToDouble(formulaResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture)); // AVERAGE
    }

    [Fact]
    public void ComplexFormulas_RealisticBusinessScenario_CalculatesCorrectly()
    {
        // Arrange - Create a realistic sales report with complex formulas
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Step 1: Set up headers
        _commands.SetValues(batch, sheetName, "A1:G1",
        [
            ["Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total Sales", "Performance"]
        ]);

        // Step 2: Set up product sales data (4 products, 4 quarters each)
        _commands.SetValues(batch, sheetName, "A2:E5",
        [
            ["Widget A", 15000, 18000, 22000, 25000],
            ["Widget B", 12000, 14000, 16000, 18000],
            ["Widget C", 8000, 9000, 11000, 13000],
            ["Widget D", 20000, 22000, 24000, 26000]
        ]);

        // Step 3: Add formulas for Total Sales (column F)
        // Using SUM function for each row
        var totalFormulas = new List<List<string>>
        {
            new() { "=SUM(B2:E2)" },
            new() { "=SUM(B3:E3)" },
            new() { "=SUM(B4:E4)" },
            new() { "=SUM(B5:E5)" }
        };
        var totalResult = _commands.SetFormulas(batch, sheetName, "F2:F5", totalFormulas);
        Assert.True(totalResult.Success, $"Failed to set total formulas: {totalResult.ErrorMessage}");

        // Step 4: Add formulas for Performance Rating (column G)
        // Using IF and AVERAGE functions
        var performanceFormulas = new List<List<string>>
        {
            new() { """=IF(AVERAGE(B2:E2)>20000,"Excellent",IF(AVERAGE(B2:E2)>15000,"Good","Average"))""" },
            new() { """=IF(AVERAGE(B3:E3)>20000,"Excellent",IF(AVERAGE(B3:E3)>15000,"Good","Average"))""" },
            new() { """=IF(AVERAGE(B4:E4)>20000,"Excellent",IF(AVERAGE(B4:E4)>15000,"Good","Average"))""" },
            new() { """=IF(AVERAGE(B5:E5)>20000,"Excellent",IF(AVERAGE(B5:E5)>15000,"Good","Average"))""" }
        };
        var perfResult = _commands.SetFormulas(batch, sheetName, "G2:G5", performanceFormulas);
        Assert.True(perfResult.Success, $"Failed to set performance formulas: {perfResult.ErrorMessage}");

        // Step 5: Add summary statistics row with complex formulas
        _commands.SetValues(batch, sheetName, "A7", [["TOTALS"]]);

        var summaryFormulas = new List<List<string>>
        {
            new()
            {
                "=SUM(B2:B5)",  // Q1 Total
                "=SUM(C2:C5)",  // Q2 Total
                "=SUM(D2:D5)",  // Q3 Total
                "=SUM(E2:E5)",  // Q4 Total
                "=SUM(F2:F5)",  // Grand Total
                "=CONCATENATE(\"Avg: \",TEXT(AVERAGE(F2:F5),\"$#,##0\"))"  // Average with formatting
            }
        };
        var summaryResult = _commands.SetFormulas(batch, sheetName, "B7:G7", summaryFormulas);
        Assert.True(summaryResult.Success, $"Failed to set summary formulas: {summaryResult.ErrorMessage}");

        // Step 6: Add growth rate calculation (comparing Q4 to Q1)
        _commands.SetValues(batch, sheetName, "H1", [["Growth Rate"]]);
        var growthFormulas = new List<List<string>>
        {
            new() { "=TEXT((E2-B2)/B2,\"0.0%\")" },
            new() { "=TEXT((E3-B3)/B3,\"0.0%\")" },
            new() { "=TEXT((E4-B4)/B4,\"0.0%\")" },
            new() { "=TEXT((E5-B5)/B5,\"0.0%\")" }
        };
        var growthResult = _commands.SetFormulas(batch, sheetName, "H2:H5", growthFormulas);
        Assert.True(growthResult.Success, $"Failed to set growth formulas: {growthResult.ErrorMessage}");

        // Act - Retrieve and verify all calculated values
        var totalsResult = _commands.GetFormulas(batch, sheetName, "F2:F5");
        var performanceResult = _commands.GetFormulas(batch, sheetName, "G2:G5");
        var summaryTotalsResult = _commands.GetFormulas(batch, sheetName, "B7:G7");
        var growthRatesResult = _commands.GetFormulas(batch, sheetName, "H2:H5");

        // Assert - Verify formula calculations
        Assert.True(totalsResult.Success);
        Assert.True(performanceResult.Success);
        Assert.True(summaryTotalsResult.Success);
        Assert.True(growthRatesResult.Success);

        // Verify Total Sales calculations
        Assert.Equal(
            80000.0,
            Convert.ToDouble(totalsResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            60000.0,
            Convert.ToDouble(totalsResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            41000.0,
            Convert.ToDouble(totalsResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            92000.0,
            Convert.ToDouble(totalsResult.Values[3][0], System.Globalization.CultureInfo.InvariantCulture));

        // Verify Performance Ratings
        Assert.Equal("Good", performanceResult.Values[0][0]);
        Assert.Equal("Average", performanceResult.Values[1][0]);
        Assert.Equal("Average", performanceResult.Values[2][0]);
        Assert.Equal("Excellent", performanceResult.Values[3][0]);

        // Verify Summary Row Calculations
        Assert.Equal(
            55000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            63000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            73000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            82000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][3], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            273000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][4], System.Globalization.CultureInfo.InvariantCulture));
        var avgText = summaryTotalsResult.Values[0][5]?.ToString() ?? string.Empty;
        Assert.Contains("68", avgText);
        Assert.Contains("250", avgText);

        // Verify Growth Rate Calculations
        Assert.Contains("%", growthRatesResult.Values[0][0]?.ToString() ?? string.Empty);
        Assert.Contains("%", growthRatesResult.Values[1][0]?.ToString() ?? string.Empty);
        Assert.Contains("%", growthRatesResult.Values[2][0]?.ToString() ?? string.Empty);
        Assert.Contains("%", growthRatesResult.Values[3][0]?.ToString() ?? string.Empty);

        // Verify formulas are preserved correctly
        Assert.Contains("SUM", totalsResult.Formulas[0][0]);
        Assert.Contains("IF", performanceResult.Formulas[0][0]);
        Assert.Contains("AVERAGE", performanceResult.Formulas[0][0]);
        Assert.Contains("CONCATENATE", summaryTotalsResult.Formulas[0][5]);
        Assert.Contains("TEXT", growthRatesResult.Formulas[0][0]);
    }

    // === EDGE CASE TESTS ===

    [Fact]
    public void SetFormulas_CrossSheetReferences_CalculatesCorrectly()
    {
        // Arrange - Test that our API correctly handles cross-sheet formula references
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create second sheet (add after the test sheet)
        string dataSheetName = $"Data_{Guid.NewGuid():N}"[..31]; // Excel sheet name max 31 chars
        batch.Execute((ctx, ct) =>
        {
            dynamic sheets = ctx.Book.Worksheets;
            dynamic sheet2 = sheets.Add();
            sheet2.Name = dataSheetName;
            return 0;
        });

        // Set up source data on data sheet
        _commands.SetValues(batch, dataSheetName, "A1:A3",
        [
            [100],
            [200],
            [300]
        ]);

        // Act - Set formulas on test sheet that reference data sheet
        var formulas = new List<List<string>>
        {
            new() { $"='{dataSheetName}'!A1", $"='{dataSheetName}'!A2", $"='{dataSheetName}'!A3" },
            new() { $"=SUM('{dataSheetName}'!A1:A3)", $"=AVERAGE('{dataSheetName}'!A1:A3)", $"=MAX('{dataSheetName}'!A1:A3)" }
        };
        var result = _commands.SetFormulas(batch, sheetName, "A1:C2", formulas);

        // Assert
        Assert.True(result.Success, $"SetFormulas with cross-sheet references failed: {result.ErrorMessage}");

        // Verify formulas are preserved with sheet references
        var formulaResult = _commands.GetFormulas(batch, sheetName, "A1:C2");
        Assert.True(formulaResult.Success);

        Assert.Contains(dataSheetName, formulaResult.Formulas[0][0]);
        Assert.Contains(dataSheetName, formulaResult.Formulas[0][1]);
        Assert.Contains(dataSheetName, formulaResult.Formulas[0][2]);
        Assert.Contains(dataSheetName, formulaResult.Formulas[1][0]);

        // Verify calculated values from cross-sheet references
        Assert.Equal(
            100.0,
            Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            200.0,
            Convert.ToDouble(formulaResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            300.0,
            Convert.ToDouble(formulaResult.Values[0][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            600.0,
            Convert.ToDouble(formulaResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            200.0,
            Convert.ToDouble(formulaResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            300.0,
            Convert.ToDouble(formulaResult.Values[1][2], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetFormulas_AbsoluteAndRelativeReferences_PreservesReferenceTypes()
    {
        // Arrange - Test that our API preserves absolute ($A$1) vs relative (A1) reference types
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up source data
        _commands.SetValues(batch, sheetName, "A1:A3",
        [
            [10],
            [20],
            [30]
        ]);

        // Act - Set formulas with different reference types
        var formulas = new List<List<string>>
        {
            new() { "=$A$1",      "=A1",       "=$A1",      "=A$1" },
            new() { "=$A$1*2",    "=A1*2",     "=$A1*2",    "=A$1*2" },
            new() { "=SUM($A$1:$A$3)", "=SUM(A1:A3)", "=SUM($A1:A3)", "=SUM(A$1:A$3)" }
        };
        var result = _commands.SetFormulas(batch, sheetName, "B1:E3", formulas);

        // Assert
        Assert.True(result.Success, $"SetFormulas with reference types failed: {result.ErrorMessage}");

        var formulaResult = _commands.GetFormulas(batch, sheetName, "B1:E3");
        Assert.True(formulaResult.Success);

        Assert.Equal("=$A$1", formulaResult.Formulas[0][0]);
        Assert.Equal("=A1", formulaResult.Formulas[0][1]);
        Assert.Equal("=$A1", formulaResult.Formulas[0][2]);
        Assert.Equal("=A$1", formulaResult.Formulas[0][3]);

        Assert.Contains("$A$1", formulaResult.Formulas[1][0]);
        Assert.Contains("A1", formulaResult.Formulas[1][1]);
        Assert.Contains("$A1", formulaResult.Formulas[1][2]);
        Assert.Contains("A$1", formulaResult.Formulas[1][3]);

        Assert.Contains("$A$1:$A$3", formulaResult.Formulas[2][0]);
        Assert.Contains("A1:A3", formulaResult.Formulas[2][1]);
        Assert.Contains("$A1:A3", formulaResult.Formulas[2][2]);
        Assert.Contains("A$1:A$3", formulaResult.Formulas[2][3]);

        Assert.Equal(
            10.0,
            Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            10.0,
            Convert.ToDouble(formulaResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            10.0,
            Convert.ToDouble(formulaResult.Values[0][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            10.0,
            Convert.ToDouble(formulaResult.Values[0][3], System.Globalization.CultureInfo.InvariantCulture));

        Assert.Equal(
            20.0,
            Convert.ToDouble(formulaResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            20.0,
            Convert.ToDouble(formulaResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            20.0,
            Convert.ToDouble(formulaResult.Values[1][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            20.0,
            Convert.ToDouble(formulaResult.Values[1][3], System.Globalization.CultureInfo.InvariantCulture));

        Assert.Equal(
            60.0,
            Convert.ToDouble(formulaResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            60.0,
            Convert.ToDouble(formulaResult.Values[2][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            60.0,
            Convert.ToDouble(formulaResult.Values[2][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            60.0,
            Convert.ToDouble(formulaResult.Values[2][3], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetFormulas_LargeFormulaSet_HandlesEfficientlyInBulk()
    {
        // Arrange - Test that our batch API handles large formula sets efficiently
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        const int rowCount = 1000;
        var sourceValues = new List<List<object?>>();
        for (int i = 1; i <= rowCount; i++)
        {
            sourceValues.Add([i, i * 2, i * 3]);
        }
        _commands.SetValues(batch, sheetName, $"A1:C{rowCount}", sourceValues);

        var formulas = new List<List<string>>();
        for (int i = 1; i <= rowCount; i++)
        {
            formulas.Add([$"=A{i}+B{i}+C{i}"]);
        }

        var startTime = DateTime.UtcNow;
        var result = _commands.SetFormulas(batch, sheetName, $"D1:D{rowCount}", formulas);
        var duration = DateTime.UtcNow - startTime;

        Assert.True(result.Success, $"SetFormulas for large set failed: {result.ErrorMessage}");
        Assert.True(duration.TotalSeconds < 10,
            $"Large formula set took too long: {duration.TotalSeconds:F2} seconds (expected < 10s)");

        var sampleResult = _commands.GetFormulas(batch, sheetName, "D1");
        Assert.Equal("=A1+B1+C1", sampleResult.Formulas[0][0]);
        Assert.Equal(
            6.0,
            Convert.ToDouble(sampleResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

        var middleResult = _commands.GetFormulas(batch, sheetName, "D500");
        Assert.Equal("=A500+B500+C500", middleResult.Formulas[0][0]);
        Assert.Equal(
            3000.0,
            Convert.ToDouble(middleResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

        var lastResult = _commands.GetFormulas(batch, sheetName, $"D{rowCount}");
        Assert.Equal($"=A{rowCount}+B{rowCount}+C{rowCount}", lastResult.Formulas[0][0]);
        Assert.Equal(
            6000.0,
            Convert.ToDouble(lastResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));

        startTime = DateTime.UtcNow;
        var bulkResult = _commands.GetFormulas(batch, sheetName, $"D1:D{rowCount}");
        duration = DateTime.UtcNow - startTime;

        Assert.True(bulkResult.Success);
        Assert.Equal(rowCount, bulkResult.Formulas.Count);
        Assert.True(duration.TotalSeconds < 5,
            $"Bulk formula read took too long: {duration.TotalSeconds:F2} seconds (expected < 5s)");
    }

    [Fact]
    public void SetFormulas_WideHorizontalRange_NoOutOfMemoryError()
    {
        // Regression test for bug where 0-based arrays caused "out of memory" error
        // User reported: Setting formulas to A2:P2 (16 columns) failed with E_OUTOFMEMORY
        // Root cause: Excel COM requires 1-based arrays, we were passing 0-based C# arrays

        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up source data in row 1
        _commands.SetValues(batch, sheetName, "A1:P1",
        [
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
        ]);

        // Create 16 formulas referencing the cells above (simulating user's table header formulas)
        var formulas = new List<List<string>>
        {
            new()
            {
                "=A1*2", "=B1*2", "=C1*2", "=D1*2",
                "=E1*2", "=F1*2", "=G1*2", "=H1*2",
                "=I1*2", "=J1*2", "=K1*2", "=L1*2",
                "=M1*2", "=N1*2", "=O1*2", "=P1*2"
            }
        };

        // Act - Write 16 formulas to A2:P2 (single row, 16 columns)
        var result = _commands.SetFormulas(batch, sheetName, "A2:P2", formulas);

        // Assert - Should succeed without "out of memory" error
        Assert.True(result.Success, $"SetFormulas failed: {result.ErrorMessage}");

        // Verify formulas were written correctly
        var readResult = _commands.GetFormulas(batch, sheetName, "A2:P2");
        Assert.True(readResult.Success);
        Assert.Single(readResult.Formulas); // One row
        Assert.Equal(16, readResult.Formulas[0].Count); // 16 columns

        // Verify first, middle, and last formulas
        Assert.Equal("=A1*2", readResult.Formulas[0][0]);
        Assert.Equal("=H1*2", readResult.Formulas[0][7]);
        Assert.Equal("=P1*2", readResult.Formulas[0][15]);

        // Verify calculated values
        Assert.Equal(2.0, Convert.ToDouble(readResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(16.0, Convert.ToDouble(readResult.Values[0][7], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(32.0, Convert.ToDouble(readResult.Values[0][15], System.Globalization.CultureInfo.InvariantCulture));
    }
}




