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
    /// Validates that the fixture successfully created the Data Model.
    /// This test makes the fixture initialization (creation test) visible in test results.
    /// Tests the following creation commands:
    /// - FileCommands.CreateEmptyAsync()
    /// - TableCommands.AddToDataModelAsync() for 3 tables
    /// - DataModelCommands.CreateRelationshipAsync() for 2 relationships
    /// - DataModelCommands.CreateMeasureAsync() for 3 measures
    /// - Batch.SaveAsync() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Slow")] // Slow because fixture runs once before this test
    public void Create_CompleteDataModel_SuccessfullyCreatesAllComponents()
    {
        // Assert - Validate the fixture creation succeeded
        Assert.True(_creationResult.Success, 
            $"Data Model creation failed during fixture initialization: {_creationResult.ErrorMessage}");
        
        // Validate all components were created
        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(3, _creationResult.TablesCreated);
        Assert.Equal(3, _creationResult.TablesLoadedToModel);
        Assert.Equal(2, _creationResult.RelationshipsCreated);
        Assert.Equal(3, _creationResult.MeasuresCreated);
        Assert.True(_creationResult.CreationTimeSeconds > 0, "Creation time should be positive");
        
        // This test appears in test results showing creation was tested
        Console.WriteLine($"✅ Data Model creation test passed in {_creationResult.CreationTimeSeconds:F1}s");
        Console.WriteLine($"   - Created {_creationResult.TablesCreated} tables");
        Console.WriteLine($"   - Loaded {_creationResult.TablesLoadedToModel} tables to Data Model");
        Console.WriteLine($"   - Created {_creationResult.RelationshipsCreated} relationships");
        Console.WriteLine($"   - Created {_creationResult.MeasuresCreated} measures");
    }

    /// <summary>
    /// Validates that Data Model persists correctly after file close/reopen.
    /// This tests Batch.SaveAsync() properly persisted all data model components.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task Create_DataModelComponents_PersistsAfterReopen()
    {
        // Act - Close and reopen to verify persistence (new batch = new session)
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        
        // Assert - Verify tables persisted
        var tables = await _dataModelCommands.ListTablesAsync(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        Assert.Equal(3, tables.Tables.Count);
        Assert.Contains(tables.Tables, t => t.Name == "SalesTable");
        Assert.Contains(tables.Tables, t => t.Name == "CustomersTable");
        Assert.Contains(tables.Tables, t => t.Name == "ProductsTable");
        
        // Assert - Verify relationships persisted
        var rels = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(rels.Success, $"ListRelationships failed: {rels.ErrorMessage}");
        Assert.Equal(2, rels.Relationships.Count);
        
        // Assert - Verify measures persisted
        var measures = await _dataModelCommands.ListMeasuresAsync(batch);
        Assert.True(measures.Success, $"ListMeasures failed: {measures.ErrorMessage}");
        Assert.Equal(3, measures.Measures.Count);
        Assert.Contains("Total Sales", measures.Measures.Select(m => m.Name));
        Assert.Contains("Average Sale", measures.Measures.Select(m => m.Name));
        Assert.Contains("Total Customers", measures.Measures.Select(m => m.Name));
    }
}
