using System;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.TestAssets;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Tests for building and verifying the Data Model template asset.
/// Run these manually to generate/regenerate the template file.
/// </summary>
[Trait("Category", "Asset")]
[Trait("RunType", "Manual")]
public class DataModelAssetBuilderTests
{
    private readonly ITestOutputHelper _output;

    public DataModelAssetBuilderTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Builds the Data Model template file.
    /// Run this when:
    /// - First time setting up tests
    /// - Data Model schema changes
    /// - Template needs regeneration
    /// 
    /// Command: dotnet test --filter "FullyQualifiedName~BuildDataModelAsset_GeneratesTemplate"
    /// </summary>
    [Fact(Skip = "Manual - run explicitly to generate template")]
    public async Task BuildDataModelAsset_GeneratesTemplate()
    {
        var solutionRoot = Path.GetFullPath(Path.Join(AppContext.BaseDirectory, "../../../../.."));
        var targetPath = Path.Join(solutionRoot, "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");

        _output.WriteLine($"Generating Data Model template...");
        _output.WriteLine($"Target: {targetPath}");
        _output.WriteLine($"Version: {DataModelAssetBuilder.ASSET_VERSION}");
        _output.WriteLine("");

        var result = await DataModelAssetBuilder.CreateDataModelAssetAsync(targetPath);

        Assert.True(File.Exists(result), $"Template file should exist at {result}");
        _output.WriteLine($"✅ Template generated successfully");
        _output.WriteLine($"   Commit this file to source control:");
        _output.WriteLine($"   git add tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");
        _output.WriteLine($"   git commit -m \"test: Regenerate Data Model template (v{DataModelAssetBuilder.ASSET_VERSION})\"");
    }

    /// <summary>
    /// Verifies the Data Model template exists and has correct version.
    /// This test runs in CI to ensure template is present and current.
    /// </summary>
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Fast")]
    [Trait("RequiresExcel", "true")]
    public async Task DataModelTemplate_ExistsAndIsCurrentVersion()
    {
        var solutionRoot = Path.GetFullPath(Path.Join(AppContext.BaseDirectory, "../../../../.."));
        var templatePath = Path.Join(solutionRoot, "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");

        // Verify file exists
        Assert.True(File.Exists(templatePath), 
            $"Data Model template not found. Generate it by running:\n" +
            $"dotnet test --filter \"FullyQualifiedName~BuildDataModelAsset_GeneratesTemplate\" /p:DefineConstants=MANUAL_TESTS\n" +
            $"Expected path: {templatePath}");

        // Verify version matches expected
        await using var batch = await ExcelSession.BeginBatchAsync(templatePath);
        var version = await batch.Execute<string>((ctx, ct) =>
        {
            try
            {
                dynamic props = ctx.Book.BuiltinDocumentProperties;
                var comments = props.Item("Comments").Value?.ToString() ?? "";
                return comments;
            }
            catch
            {
                return "";
            }
        });

        _output.WriteLine($"Template version: {version}");
        _output.WriteLine($"Expected version: {DataModelAssetBuilder.ASSET_VERSION}");

        Assert.Contains(DataModelAssetBuilder.ASSET_VERSION, version);
        
        if (version != DataModelAssetBuilder.ASSET_VERSION)
        {
            _output.WriteLine($"Template version mismatch. Expected v{DataModelAssetBuilder.ASSET_VERSION}.");
            _output.WriteLine("Regenerate template by running BuildDataModelAsset_GeneratesTemplate test.");
        }
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
