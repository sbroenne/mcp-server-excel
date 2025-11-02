using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for Power Query Core operations.
/// Uses PowerQueryTestsFixture which creates ONE Power Query file per test class (~10-15s setup).
/// Fixture initialization IS the test for Power Query creation - validates ImportAsync command.
/// Each test gets its own batch for isolation.
/// Each test uses a unique Excel file for complete test isolation.
///
/// For comprehensive workflow tests (mode switching), see PowerQueryLoadConfigWorkflowTests.cs.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public partial class PowerQueryCommandsTests : IClassFixture<PowerQueryTestsFixture>
{
    protected readonly IPowerQueryCommands _powerQueryCommands;
    protected readonly IFileCommands _fileCommands;
    protected readonly ISheetCommands _sheetCommands;
    protected readonly string _powerQueryFile;
    protected readonly PowerQueryCreationResult _creationResult;
    protected readonly string _tempDir;

    public PowerQueryCommandsTests(PowerQueryTestsFixture fixture)
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();
        _sheetCommands = new SheetCommands();
        _powerQueryFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _tempDir = Path.GetDirectoryName(fixture.TestFilePath)!;
    }

    /// <summary>
    /// Explicit test that validates the fixture creation results.
    /// This makes the creation test visible in test results and validates:
    /// - FileCommands.CreateEmptyAsync()
    /// - PowerQueryCommands.ImportAsync() for all queries
    /// - Batch.SaveAsync() persistence
    /// </summary>
    [Fact]
    [Trait("Speed", "Fast")]
    public void PowerQueryCreation_ViaFixture_CreatesQueriesSuccessfully()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success, 
            $"Power Query creation failed during fixture initialization: {_creationResult.ErrorMessage}");
        
        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(3, _creationResult.MCodeFilesCreated);
        Assert.Equal(3, _creationResult.QueriesImported);
        Assert.True(_creationResult.CreationTimeSeconds > 0);
        
        // This test appears in test results as proof that creation was tested
        Console.WriteLine($"✅ Power Queries created successfully in {_creationResult.CreationTimeSeconds:F1}s");
    }

    /// <summary>
    /// Tests that Power Queries persist correctly after file close/reopen.
    /// Validates that SaveAsync() properly persisted all Power Queries.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task PowerQueryCreation_Persists_AfterReopenFile()
    {
        // Close and reopen to verify persistence (new batch = new session)
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        
        // Verify queries persisted
        var result = await _powerQueryCommands.ListAsync(batch);
        Assert.True(result.Success, $"ListAsync failed: {result.ErrorMessage}");
        Assert.Equal(3, result.Queries.Count);
        Assert.Contains(result.Queries, q => q.Name == "BasicQuery");
        Assert.Contains(result.Queries, q => q.Name == "DataQuery");
        Assert.Contains(result.Queries, q => q.Name == "RefreshableQuery");
        
        // This proves creation + save worked correctly
    }

    /// <summary>
    /// Creates a unique test Power Query M code file.
    /// Used by tests that need to create new queries.
    /// </summary>
    protected string CreateUniqueTestQueryFile(string testName)
    {
        var uniqueFile = Path.Join(_tempDir, $"{testName}_{Guid.NewGuid():N}.pq");
        string mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        System.IO.File.WriteAllText(uniqueFile, mCode);
        return uniqueFile;
    }
}
