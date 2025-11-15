using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range formulas operations
/// </summary>
public partial class RangeCommandsTests
{
    /// <inheritdoc/>
    // === FORMULA OPERATIONS TESTS ===

    [Fact]
    public async Task GetFormulas_ReturnsFormulasAndValues()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(GetFormulas_ReturnsFormulasAndValues), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set values and formulas
        await _commands.SetValues(batch, "Sheet1", "A1:A3",
        [
            [10],
            [20],
            [30]
        ]);

        await _commands.SetFormulas(batch, "Sheet1", "B1",
        [
            ["=SUM(A1:A3)"]
        ]);

        // Act
        var result = _commands.GetFormulas(batch, "Sheet1", "B1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal("=SUM(A1:A3)", result.Formulas[0][0]);
        Assert.Equal(
            60.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetFormulas_WritesFormulasToRange()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(SetFormulas_WritesFormulasToRange), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        await _commands.SetValues(batch, "Sheet1", "A1:A3",
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
        var result = _commands.SetFormulas(batch, "Sheet1", "B1:D1", formulas);
        // Assert
        Assert.True(result.Success);

        // Verify values
        var readResult = _commands.GetValues(batch, "Sheet1", "B1:D1");
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
    /// <inheritdoc/>

    [Fact]
    public async Task SetFormulas_WithJsonElementFormulas_WritesFormulasCorrectly()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFile(nameof(RangeCommandsTests), nameof(SetFormulas_WithJsonElementFormulas_WritesFormulasCorrectly), _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up source data
        await _commands.SetValues(batch, "Sheet1", "A1:A3",
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
        var result = _commands.SetFormulas(batch, "Sheet1", "B1:C1", testFormulas);
        // Assert
        Assert.True(result.Success, $"SetFormulasAsync failed: {result.ErrorMessage}");

        // Verify formulas were written correctly
        var formulaResult = _commands.GetFormulas(batch, "Sheet1", "B1:C1");
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
    /// <inheritdoc/>

    [Fact]
    public async Task ComplexFormulas_RealisticBusinessScenario_CalculatesCorrectly()
    {
        // Arrange - Create a realistic sales report with complex formulas
        string testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(ComplexFormulas_RealisticBusinessScenario_CalculatesCorrectly),
            _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Step 1: Set up headers
        await _commands.SetValues(batch, "Sheet1", "A1:G1",
        [
            ["Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total Sales", "Performance"]
        ]);

        // Step 2: Set up product sales data (4 products, 4 quarters each)
        await _commands.SetValues(batch, "Sheet1", "A2:E5",
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
        var totalResult = _commands.SetFormulas(batch, "Sheet1", "F2:F5", totalFormulas);
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
        var perfResult = _commands.SetFormulas(batch, "Sheet1", "G2:G5", performanceFormulas);
        Assert.True(perfResult.Success, $"Failed to set performance formulas: {perfResult.ErrorMessage}");

        // Step 5: Add summary statistics row with complex formulas
        await _commands.SetValues(batch, "Sheet1", "A7", [["TOTALS"]]);

        var summaryFormulas = new List<List<string>>
        {
            new() {
                "=SUM(B2:B5)",  // Q1 Total
                "=SUM(C2:C5)",  // Q2 Total
                "=SUM(D2:D5)",  // Q3 Total
                "=SUM(E2:E5)",  // Q4 Total
                "=SUM(F2:F5)",  // Grand Total
                "=CONCATENATE(\"Avg: \",TEXT(AVERAGE(F2:F5),\"$#,##0\"))"  // Average with formatting
            }
        };
        var summaryResult = _commands.SetFormulas(batch, "Sheet1", "B7:G7", summaryFormulas);
        Assert.True(summaryResult.Success, $"Failed to set summary formulas: {summaryResult.ErrorMessage}");

        // Step 6: Add growth rate calculation (comparing Q4 to Q1)
        await _commands.SetValues(batch, "Sheet1", "H1", [["Growth Rate"]]);
        var growthFormulas = new List<List<string>>
        {
            new() { "=TEXT((E2-B2)/B2,\"0.0%\")" },
            new() { "=TEXT((E3-B3)/B3,\"0.0%\")" },
            new() { "=TEXT((E4-B4)/B4,\"0.0%\")" },
            new() { "=TEXT((E5-B5)/B5,\"0.0%\")" }
        };
        var growthResult = _commands.SetFormulas(batch, "Sheet1", "H2:H5", growthFormulas);
        Assert.True(growthResult.Success, $"Failed to set growth formulas: {growthResult.ErrorMessage}");

        // Act - Retrieve and verify all calculated values
        var totalsResult = _commands.GetFormulas(batch, "Sheet1", "F2:F5");
        var performanceResult = _commands.GetFormulas(batch, "Sheet1", "G2:G5");
        var summaryTotalsResult = _commands.GetFormulas(batch, "Sheet1", "B7:G7");
        var growthRatesResult = _commands.GetFormulas(batch, "Sheet1", "H2:H5");

        // Assert - Verify formula calculations
        Assert.True(totalsResult.Success);
        Assert.True(performanceResult.Success);
        Assert.True(summaryTotalsResult.Success);
        Assert.True(growthRatesResult.Success);

        // Verify Total Sales calculations
        Assert.Equal(
            80000.0,
            Convert.ToDouble(totalsResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // Widget A: 15000+18000+22000+25000
        Assert.Equal(
            60000.0,
            Convert.ToDouble(totalsResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture)); // Widget B: 12000+14000+16000+18000
        Assert.Equal(
            41000.0,
            Convert.ToDouble(totalsResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture)); // Widget C: 8000+9000+11000+13000
        Assert.Equal(
            92000.0,
            Convert.ToDouble(totalsResult.Values[3][0], System.Globalization.CultureInfo.InvariantCulture)); // Widget D: 20000+22000+24000+26000

        // Verify Performance Ratings (IF/AVERAGE logic)
        // Formula: IF(AVERAGE>20000,"Excellent",IF(AVERAGE>15000,"Good","Average"))
        // Widget A avg: (15000+18000+22000+25000)/4 = 20000 (20000 > 15000 = TRUE, so "Good")
        // Widget B avg: (12000+14000+16000+18000)/4 = 15000 (15000 NOT > 15000, so "Average")
        // Widget C avg: (8000+9000+11000+13000)/4 = 10250 (NOT > 15000, so "Average")
        // Widget D avg: (20000+22000+24000+26000)/4 = 23000 (23000 > 20000, so "Excellent")
        Assert.Equal("Good", performanceResult.Values[0][0]);
        Assert.Equal("Average", performanceResult.Values[1][0]);
        Assert.Equal("Average", performanceResult.Values[2][0]);
        Assert.Equal("Excellent", performanceResult.Values[3][0]);

        // Verify Summary Row Calculations
        Assert.Equal(
            55000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // Q1 Total: 15000+12000+8000+20000
        Assert.Equal(
            63000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture)); // Q2 Total
        Assert.Equal(
            73000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][2], System.Globalization.CultureInfo.InvariantCulture)); // Q3 Total
        Assert.Equal(
            82000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][3], System.Globalization.CultureInfo.InvariantCulture)); // Q4 Total
        Assert.Equal(
            273000.0,
            Convert.ToDouble(summaryTotalsResult.Values[0][4], System.Globalization.CultureInfo.InvariantCulture)); // Grand Total
        // Note: TEXT formatting includes thousands separator, locale-dependent
        var avgText = summaryTotalsResult.Values[0][5]?.ToString() ?? "";
        Assert.Contains("68", avgText); // CONCATENATE + TEXT formatting: "Avg: $68,250"
        Assert.Contains("250", avgText); // Verify both parts (thousands separator may be present)

        // Verify Growth Rate Calculations (TEXT formatted percentages)
        // Note: TEXT formatting rounds, so exact format may vary by locale
        Assert.Contains("%", growthRatesResult.Values[0][0]?.ToString() ?? ""); // Widget A: (25000-15000)/15000
        Assert.Contains("%", growthRatesResult.Values[1][0]?.ToString() ?? ""); // Widget B: (18000-12000)/12000
        Assert.Contains("%", growthRatesResult.Values[2][0]?.ToString() ?? ""); // Widget C: (13000-8000)/8000
        Assert.Contains("%", growthRatesResult.Values[3][0]?.ToString() ?? ""); // Widget D: (26000-20000)/20000

        // Verify formulas are preserved correctly
        Assert.Contains("SUM", totalsResult.Formulas[0][0]);
        Assert.Contains("IF", performanceResult.Formulas[0][0]);
        Assert.Contains("AVERAGE", performanceResult.Formulas[0][0]);
        Assert.Contains("CONCATENATE", summaryTotalsResult.Formulas[0][5]);
        Assert.Contains("TEXT", growthRatesResult.Formulas[0][0]);
    }
    /// <inheritdoc/>

    // === EDGE CASE TESTS ===

    [Fact]
    public async Task SetFormulas_CrossSheetReferences_CalculatesCorrectly()
    {
        // Arrange - Test that our API correctly handles cross-sheet formula references
        string testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetFormulas_CrossSheetReferences_CalculatesCorrectly),
            _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create second sheet (add after Sheet1 to avoid reordering)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheets = ctx.Book.Worksheets;
            dynamic sheet1 = sheets.Item(1);
            dynamic sheet2 = sheets.Add(After: sheet1);
            sheet2.Name = "Data";
            return 0;
        });

        // Set up source data on "Data" sheet
        await _commands.SetValues(batch, "Data", "A1:A3",
        [
            [100],
            [200],
            [300]
        ]);

        // Act - Set formulas on Sheet1 that reference Data sheet
        var formulas = new List<List<string>>
        {
            new() { "=Data!A1", "=Data!A2", "=Data!A3" },
            new() { "=SUM(Data!A1:A3)", "=AVERAGE(Data!A1:A3)", "=MAX(Data!A1:A3)" }
        };
        var result = _commands.SetFormulas(batch, "Sheet1", "A1:C2", formulas);

        // Assert
        Assert.True(result.Success, $"SetFormulas with cross-sheet references failed: {result.ErrorMessage}");

        // Verify formulas are preserved with sheet references
        var formulaResult = _commands.GetFormulas(batch, "Sheet1", "A1:C2");
        Assert.True(formulaResult.Success);

        // Verify formula strings contain sheet references
        Assert.Contains("Data!", formulaResult.Formulas[0][0]);
        Assert.Contains("Data!", formulaResult.Formulas[0][1]);
        Assert.Contains("Data!", formulaResult.Formulas[0][2]);
        Assert.Contains("Data!", formulaResult.Formulas[1][0]);

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
            Convert.ToDouble(formulaResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture)); // SUM
        Assert.Equal(
            200.0,
            Convert.ToDouble(formulaResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture)); // AVERAGE
        Assert.Equal(
            300.0,
            Convert.ToDouble(formulaResult.Values[1][2], System.Globalization.CultureInfo.InvariantCulture)); // MAX
    }
    /// <inheritdoc/>

    [Fact]
    public async Task SetFormulas_AbsoluteAndRelativeReferences_PreservesReferenceTypes()
    {
        // Arrange - Test that our API preserves absolute ($A$1) vs relative (A1) reference types
        string testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetFormulas_AbsoluteAndRelativeReferences_PreservesReferenceTypes),
            _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up source data
        await _commands.SetValues(batch, "Sheet1", "A1:A3",
        [
            [10],
            [20],
            [30]
        ]);

        // Act - Set formulas with different reference types
        var formulas = new List<List<string>>
        {
            new() { "=$A$1",      "=A1",       "=$A1",      "=A$1" },  // Row 1: Different reference types
            new() { "=$A$1*2",    "=A1*2",     "=$A1*2",    "=A$1*2" }, // Row 2: Reference types in expressions
            new() { "=SUM($A$1:$A$3)", "=SUM(A1:A3)", "=SUM($A1:A3)", "=SUM(A$1:A$3)" } // Row 3: Range references
        };
        var result = _commands.SetFormulas(batch, "Sheet1", "B1:E3", formulas);

        // Assert
        Assert.True(result.Success, $"SetFormulas with reference types failed: {result.ErrorMessage}");

        // Verify formula strings are preserved exactly as set (round-trip test)
        var formulaResult = _commands.GetFormulas(batch, "Sheet1", "B1:E3");
        Assert.True(formulaResult.Success);

        // Row 1 - Reference type preservation
        Assert.Equal("=$A$1", formulaResult.Formulas[0][0]);     // Absolute
        Assert.Equal("=A1", formulaResult.Formulas[0][1]);       // Relative
        Assert.Equal("=$A1", formulaResult.Formulas[0][2]);      // Mixed (column absolute)
        Assert.Equal("=A$1", formulaResult.Formulas[0][3]);      // Mixed (row absolute)

        // Row 2 - Reference types in expressions
        Assert.Contains("$A$1", formulaResult.Formulas[1][0]);   // Absolute in expression
        Assert.Contains("A1", formulaResult.Formulas[1][1]);     // Relative in expression
        Assert.Contains("$A1", formulaResult.Formulas[1][2]);    // Mixed in expression
        Assert.Contains("A$1", formulaResult.Formulas[1][3]);    // Mixed in expression

        // Row 3 - Range references
        Assert.Contains("$A$1:$A$3", formulaResult.Formulas[2][0]); // Absolute range
        Assert.Contains("A1:A3", formulaResult.Formulas[2][1]);     // Relative range
        Assert.Contains("$A1:A3", formulaResult.Formulas[2][2]);    // Mixed range
        Assert.Contains("A$1:A$3", formulaResult.Formulas[2][3]);   // Mixed range

        // Verify all formulas calculate correctly (value should be same regardless of reference type)
        Assert.Equal(
            10.0,
            Convert.ToDouble(formulaResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // All reference A1
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
            Convert.ToDouble(formulaResult.Values[1][0], System.Globalization.CultureInfo.InvariantCulture)); // All multiply by 2
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
            Convert.ToDouble(formulaResult.Values[2][0], System.Globalization.CultureInfo.InvariantCulture)); // All SUM A1:A3
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
    /// <inheritdoc/>

    [Fact]
    public async Task SetFormulas_LargeFormulaSet_HandlesEfficientlyInBulk()
    {
        // Arrange - Test that our batch API handles large formula sets efficiently
        string testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(RangeCommandsTests),
            nameof(SetFormulas_LargeFormulaSet_HandlesEfficientlyInBulk),
            _tempDir);
        using var batch = ExcelSession.BeginBatch(testFile);

        // Set up source data (1000 rows)
        const int rowCount = 1000;
        var sourceValues = new List<List<object?>>();
        for (int i = 1; i <= rowCount; i++)
        {
            sourceValues.Add([i, i * 2, i * 3]); // Columns A, B, C
        }
        await _commands.SetValues(batch, "Sheet1", $"A1:C{rowCount}", sourceValues);

        // Generate 1000 formulas for column D (sum of A, B, C)
        var formulas = new List<List<string>>();
        for (int i = 1; i <= rowCount; i++)
        {
            formulas.Add([$"=A{i}+B{i}+C{i}"]); // Each row sums A, B, C
        }

        // Act - Set all 1000 formulas in one batch operation
        var startTime = DateTime.UtcNow;
        var result = _commands.SetFormulas(batch, "Sheet1", $"D1:D{rowCount}", formulas);
        var duration = DateTime.UtcNow - startTime;

        // Assert
        Assert.True(result.Success, $"SetFormulas for large set failed: {result.ErrorMessage}");

        // Verify performance (should complete in reasonable time - under 10 seconds for 1000 formulas)
        Assert.True(duration.TotalSeconds < 10,
            $"Large formula set took too long: {duration.TotalSeconds:F2} seconds (expected < 10s)");

        // Sample verification - check first, middle, and last formulas
        var sampleResult = _commands.GetFormulas(batch, "Sheet1", "D1");
        Assert.Equal("=A1+B1+C1", sampleResult.Formulas[0][0]);
        Assert.Equal(
            6.0,
            Convert.ToDouble(sampleResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // 1+2+3

        var middleResult = _commands.GetFormulas(batch, "Sheet1", "D500");
        Assert.Equal("=A500+B500+C500", middleResult.Formulas[0][0]);
        Assert.Equal(
            3000.0,
            Convert.ToDouble(middleResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // 500+1000+1500

        var lastResult = _commands.GetFormulas(batch, "Sheet1", $"D{rowCount}");
        Assert.Equal($"=A{rowCount}+B{rowCount}+C{rowCount}", lastResult.Formulas[0][0]);
        Assert.Equal(
            6000.0,
            Convert.ToDouble(lastResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture)); // 1000+2000+3000

        // Verify bulk read performance - retrieve all 1000 formulas at once
        startTime = DateTime.UtcNow;
        var bulkResult = _commands.GetFormulas(batch, "Sheet1", $"D1:D{rowCount}");
        duration = DateTime.UtcNow - startTime;

        Assert.True(bulkResult.Success);
        Assert.Equal(rowCount, bulkResult.Formulas.Count);
        Assert.True(duration.TotalSeconds < 5,
            $"Bulk formula read took too long: {duration.TotalSeconds:F2} seconds (expected < 5s)");
    }
}
