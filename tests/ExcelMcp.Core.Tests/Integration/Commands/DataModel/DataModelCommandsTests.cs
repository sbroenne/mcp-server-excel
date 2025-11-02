using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Data Model operations integration tests.
/// Uses DataModelTestsFixture which creates ONE Data Model per test class (60-120s setup).
/// Fixture initialization IS the test for Data Model creation - validates all creation commands.
/// Each test gets its own batch for isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
public partial class DataModelCommandsTests : IClassFixture<DataModelTestsFixture>
{
    protected readonly IDataModelCommands _dataModelCommands;
    protected readonly string _dataModelFile;
    protected readonly DataModelCreationResult _creationResult;

    public DataModelCommandsTests(DataModelTestsFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _dataModelFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the creation test visible in test results and validates:
    /// - FileCommands.CreateEmptyAsync()
    /// - TableCommands.AddToDataModelAsync() for all tables
    /// - DataModelCommands.CreateRelationshipAsync() for all relationships
    /// - DataModelCommands.CreateMeasureAsync() for all measures
    /// - Batch.SaveAsync() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void DataModelCreation_ViaFixture_CreatesCompleteModel()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success, 
            $"Data Model creation failed during fixture initialization: {_creationResult.ErrorMessage}");
        
        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(3, _creationResult.TablesCreated);
        Assert.Equal(3, _creationResult.TablesLoadedToModel);
        Assert.Equal(2, _creationResult.RelationshipsCreated);
        Assert.Equal(3, _creationResult.MeasuresCreated);
        Assert.True(_creationResult.CreationTimeSeconds > 0);
        
        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"✅ Data Model created successfully in {_creationResult.CreationTimeSeconds:F1}s");
    }

    /// <summary>
    /// Tests that Data Model persists correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted all data model components.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task DataModelCreation_Persists_AfterReopenFile()
    {
        // Close and reopen to verify persistence (new batch = new session)
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        
        // Verify tables persisted
        var tables = await _dataModelCommands.ListTablesAsync(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        Assert.Equal(3, tables.Tables.Count);
        Assert.Contains(tables.Tables, t => t.Name == "SalesTable");
        Assert.Contains(tables.Tables, t => t.Name == "CustomersTable");
        Assert.Contains(tables.Tables, t => t.Name == "ProductsTable");
        
        // Verify relationships persisted
        var rels = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(rels.Success, $"ListRelationships failed: {rels.ErrorMessage}");
        Assert.Equal(2, rels.Relationships.Count);
        
        // Verify measures persisted
        var measures = await _dataModelCommands.ListMeasuresAsync(batch);
        Assert.True(measures.Success, $"ListMeasures failed: {measures.ErrorMessage}");
        Assert.Equal(3, measures.Measures.Count);
        Assert.Contains("Total Sales", measures.Measures.Select(m => m.Name));
        Assert.Contains("Average Sale", measures.Measures.Select(m => m.Name));
        Assert.Contains("Total Customers", measures.Measures.Select(m => m.Name));
        
        // This proves creation + save worked correctly
    }
}
