using System;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Tests for verifying the Data Model template asset exists.
/// To regenerate template, run: dotnet script BuildDataModelTemplate.csx
/// </summary>
[Trait("Category", "Asset")]
public class DataModelAssetBuilderTests
{
    private readonly ITestOutputHelper _output;

    public DataModelAssetBuilderTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Verifies the Data Model template exists.
    /// This test runs in CI to ensure template is present.
    /// Template is stored in git - regenerate only if structure changes.
    /// </summary>
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Fast")]
    [Trait("RequiresExcel", "true")]
    public async Task DataModelTemplate_Exists()
    {
        var solutionRoot = Path.GetFullPath(Path.Join(AppContext.BaseDirectory, "../../../../.."));
        var templatePath = Path.Join(solutionRoot, "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");

        // Verify file exists
        Assert.True(File.Exists(templatePath),
            $"Data Model template not found at: {templatePath}\n" +
            $"This file should be in git. If missing, see tests/ExcelMcp.Core.Tests/docs/DATA-MODEL-SETUP.md");

        // Verify it can be opened (basic sanity check)
        await using var batch = await ExcelSession.BeginBatchAsync(templatePath);
        var result = await batch.Execute<int>((ctx, ct) =>
        {
            return ctx.Book.Worksheets.Count;
        });

        Assert.True(result > 0, "Template should have at least one worksheet");
    }

    /// <summary>
    /// Verifies the Data Model template has expected structure.
    /// </summary>
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Fast")]
    [Trait("RequiresExcel", "true")]
    public async Task DataModelTemplate_HasExpectedStructure()
    {
        var fixture = new DataModelReadTestsFixture();
        await fixture.InitializeAsync();

        try
        {
            var commands = new Core.Commands.DataModelCommands();
            await using var batch = await ExcelSession.BeginBatchAsync(fixture.TestFilePath);

            // Verify tables
            var tablesResult = await commands.ListTablesAsync(batch);
            Assert.True(tablesResult.Success, $"Failed to list tables: {tablesResult.ErrorMessage}");
            Assert.Equal(3, tablesResult.Tables.Count);
            Assert.Contains(tablesResult.Tables, t => t.Name == "SalesTable");
            Assert.Contains(tablesResult.Tables, t => t.Name == "CustomersTable");
            Assert.Contains(tablesResult.Tables, t => t.Name == "ProductsTable");

            // Verify relationships
            var relationshipsResult = await commands.ListRelationshipsAsync(batch);
            Assert.True(relationshipsResult.Success, $"Failed to list relationships: {relationshipsResult.ErrorMessage}");
            Assert.Equal(2, relationshipsResult.Relationships.Count);

            // Verify measures
            var measuresResult = await commands.ListMeasuresAsync(batch);
            Assert.True(measuresResult.Success, $"Failed to list measures: {measuresResult.ErrorMessage}");
            Assert.Equal(3, measuresResult.Measures.Count);
            Assert.Contains("Total Sales", measuresResult.Measures.Select(m => m.Name));
            Assert.Contains("Average Sale", measuresResult.Measures.Select(m => m.Name));
            Assert.Contains("Total Customers", measuresResult.Measures.Select(m => m.Name));
        }
        finally
        {
            await fixture.DisposeAsync();
        }
    }
}
