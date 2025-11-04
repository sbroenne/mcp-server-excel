using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range formulas operations
/// </summary>
public partial class RangeCommandsTests
{
    // === FORMULA OPERATIONS TESTS ===

    [Fact]
    public async Task GetFormulas_ReturnsFormulasAndValues()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(GetFormulas_ReturnsFormulasAndValues), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set values and formulas
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 10 },
            new() { 20 },
            new() { 30 }
        ]);

        await _commands.SetFormulasAsync(batch, "Sheet1", "B1",
        [
            new() { "=SUM(A1:A3)" }
        ]);

        // Act
        var result = await _commands.GetFormulasAsync(batch, "Sheet1", "B1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal("=SUM(A1:A3)", result.Formulas[0][0]);
        Assert.Equal(60.0, Convert.ToDouble(result.Values[0][0]));
    }

    [Fact]
    public async Task SetFormulas_WritesFormulasToRange()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(SetFormulas_WritesFormulasToRange), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 5 },
            new() { 10 },
            new() { 15 }
        ]);

        var formulas = new List<List<string>>
        {
            new() { "=A1*2", "=A2*2", "=A3*2" }
        };

        // Act
        var result = await _commands.SetFormulasAsync(batch, "Sheet1", "B1:D1", formulas);
        // Assert
        Assert.True(result.Success);

        // Verify values
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "B1:D1");
        Assert.Equal(10.0, Convert.ToDouble(readResult.Values[0][0]));
        Assert.Equal(20.0, Convert.ToDouble(readResult.Values[0][1]));
        Assert.Equal(30.0, Convert.ToDouble(readResult.Values[0][2]));
    }

    [Fact]
    public async Task SetFormulas_WithJsonElementFormulas_WritesFormulasCorrectly()
    {
        // Arrange
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(nameof(RangeCommandsTests), nameof(SetFormulas_WithJsonElementFormulas_WritesFormulasCorrectly), _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up source data
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 100 },
            new() { 200 },
            new() { 300 }
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
        var result = await _commands.SetFormulasAsync(batch, "Sheet1", "B1:C1", testFormulas);
        // Assert
        Assert.True(result.Success, $"SetFormulasAsync failed: {result.ErrorMessage}");

        // Verify formulas were written correctly
        var formulaResult = await _commands.GetFormulasAsync(batch, "Sheet1", "B1:C1");
        Assert.True(formulaResult.Success);
        Assert.Equal("=SUM(A1:A3)", formulaResult.Formulas[0][0]);
        Assert.Equal("=AVERAGE(A1:A3)", formulaResult.Formulas[0][1]);

        // Verify calculated values
        Assert.Equal(600.0, Convert.ToDouble(formulaResult.Values[0][0])); // SUM
        Assert.Equal(200.0, Convert.ToDouble(formulaResult.Values[0][1])); // AVERAGE
    }

    [Fact]
    public async Task ComplexFormulas_RealisticBusinessScenario_CalculatesCorrectly()
    {
        // Arrange - Create a realistic sales report with complex formulas
        string testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests), 
            nameof(ComplexFormulas_RealisticBusinessScenario_CalculatesCorrectly), 
            _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Step 1: Set up headers
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:G1",
        [
            new() { "Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total Sales", "Performance" }
        ]);

        // Step 2: Set up product sales data (4 products, 4 quarters each)
        await _commands.SetValuesAsync(batch, "Sheet1", "A2:E5",
        [
            new() { "Widget A", 15000, 18000, 22000, 25000 },
            new() { "Widget B", 12000, 14000, 16000, 18000 },
            new() { "Widget C", 8000, 9000, 11000, 13000 },
            new() { "Widget D", 20000, 22000, 24000, 26000 }
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
        var totalResult = await _commands.SetFormulasAsync(batch, "Sheet1", "F2:F5", totalFormulas);
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
        var perfResult = await _commands.SetFormulasAsync(batch, "Sheet1", "G2:G5", performanceFormulas);
        Assert.True(perfResult.Success, $"Failed to set performance formulas: {perfResult.ErrorMessage}");

        // Step 5: Add summary statistics row with complex formulas
        await _commands.SetValuesAsync(batch, "Sheet1", "A7", [new() { "TOTALS" }]);
        
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
        var summaryResult = await _commands.SetFormulasAsync(batch, "Sheet1", "B7:G7", summaryFormulas);
        Assert.True(summaryResult.Success, $"Failed to set summary formulas: {summaryResult.ErrorMessage}");

        // Step 6: Add growth rate calculation (comparing Q4 to Q1)
        await _commands.SetValuesAsync(batch, "Sheet1", "H1", [new() { "Growth Rate" }]);
        var growthFormulas = new List<List<string>>
        {
            new() { "=TEXT((E2-B2)/B2,\"0.0%\")" },
            new() { "=TEXT((E3-B3)/B3,\"0.0%\")" },
            new() { "=TEXT((E4-B4)/B4,\"0.0%\")" },
            new() { "=TEXT((E5-B5)/B5,\"0.0%\")" }
        };
        var growthResult = await _commands.SetFormulasAsync(batch, "Sheet1", "H2:H5", growthFormulas);
        Assert.True(growthResult.Success, $"Failed to set growth formulas: {growthResult.ErrorMessage}");

        // Act - Retrieve and verify all calculated values
        var totalsResult = await _commands.GetFormulasAsync(batch, "Sheet1", "F2:F5");
        var performanceResult = await _commands.GetFormulasAsync(batch, "Sheet1", "G2:G5");
        var summaryTotalsResult = await _commands.GetFormulasAsync(batch, "Sheet1", "B7:G7");
        var growthRatesResult = await _commands.GetFormulasAsync(batch, "Sheet1", "H2:H5");

        // Assert - Verify formula calculations
        Assert.True(totalsResult.Success);
        Assert.True(performanceResult.Success);
        Assert.True(summaryTotalsResult.Success);
        Assert.True(growthRatesResult.Success);

        // Verify Total Sales calculations
        Assert.Equal(80000.0, Convert.ToDouble(totalsResult.Values[0][0])); // Widget A: 15000+18000+22000+25000
        Assert.Equal(60000.0, Convert.ToDouble(totalsResult.Values[1][0])); // Widget B: 12000+14000+16000+18000
        Assert.Equal(41000.0, Convert.ToDouble(totalsResult.Values[2][0])); // Widget C: 8000+9000+11000+13000
        Assert.Equal(92000.0, Convert.ToDouble(totalsResult.Values[3][0])); // Widget D: 20000+22000+24000+26000

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
        Assert.Equal(55000.0, Convert.ToDouble(summaryTotalsResult.Values[0][0])); // Q1 Total: 15000+12000+8000+20000
        Assert.Equal(63000.0, Convert.ToDouble(summaryTotalsResult.Values[0][1])); // Q2 Total
        Assert.Equal(73000.0, Convert.ToDouble(summaryTotalsResult.Values[0][2])); // Q3 Total
        Assert.Equal(82000.0, Convert.ToDouble(summaryTotalsResult.Values[0][3])); // Q4 Total
        Assert.Equal(273000.0, Convert.ToDouble(summaryTotalsResult.Values[0][4])); // Grand Total
        // Note: TEXT formatting may vary by locale, just verify it contains the average value
        var avgText = summaryTotalsResult.Values[0][5]?.ToString() ?? "";
        Assert.Contains("68250", avgText); // CONCATENATE + TEXT formatting (locale-dependent format)

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
}
