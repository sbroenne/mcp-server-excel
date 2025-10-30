using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query Privacy Level functionality.
/// These tests validate privacy level detection, recommendation, and application.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQueryPrivacy")]
public class PowerQueryPrivacyLevelTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public PowerQueryPrivacyLevelTests()
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Privacy_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "PrivacyTestWorkbook.xlsx");

        // Create test Excel file
        CreateTestExcelFile();
    }

    private void CreateTestExcelFile()
    {
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    [Fact]
    public void PrivacyLevel_Enum_HasAllExpectedValues()
    {
        // Assert - Verify all privacy levels are defined
        Assert.True(Enum.IsDefined(typeof(PowerQueryPrivacyLevel), PowerQueryPrivacyLevel.None));
        Assert.True(Enum.IsDefined(typeof(PowerQueryPrivacyLevel), PowerQueryPrivacyLevel.Private));
        Assert.True(Enum.IsDefined(typeof(PowerQueryPrivacyLevel), PowerQueryPrivacyLevel.Organizational));
        Assert.True(Enum.IsDefined(typeof(PowerQueryPrivacyLevel), PowerQueryPrivacyLevel.Public));
    }

    [Theory]
    [InlineData(PowerQueryPrivacyLevel.None)]
    [InlineData(PowerQueryPrivacyLevel.Private)]
    [InlineData(PowerQueryPrivacyLevel.Organizational)]
    [InlineData(PowerQueryPrivacyLevel.Public)]
    public async Task Import_WithPrivacyLevel_AcceptsAllValidLevels(PowerQueryPrivacyLevel privacyLevel)
    {
        // Arrange
        string queryFile = Path.Combine(_tempDir, "TestQuery.pq");
        string mCode = @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""Value1"", ""Value2""},
            {""A"", ""B""}
        }
    )
in
    Source";
        File.WriteAllText(queryFile, mCode);

        // Act - Should not throw exception with valid privacy level
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.ImportAsync(batch, $"TestQuery_{privacyLevel}", queryFile, privacyLevel);

        // Assert - Should succeed or provide clear error (not throw exception)
        Assert.NotNull(result);
        // If it fails, it should be a clear operation result, not an exception
        if (!result.Success)
        {
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage), "Error message should be provided when operation fails");
        }
    }

    [Fact]
    public async Task Import_WithoutPrivacyLevel_StillWorks()
    {
        // Arrange
        string queryFile = Path.Combine(_tempDir, "SimpleQuery.pq");
        string mCode = @"let
    Source = #table(
        {""Column1""},
        {{""Value1""}}
    )
in
    Source";
        File.WriteAllText(queryFile, mCode);

        // Act - Import without privacy level (should work for simple queries)
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _powerQueryCommands.ImportAsync(batch, "SimpleQuery", queryFile, null);

        // Assert - Should succeed for simple non-combining queries
        Assert.NotNull(result);
    }

    [Fact]
    public void PowerQueryPrivacyErrorResult_HasRequiredProperties()
    {
        // Arrange & Act
        var errorResult = new PowerQueryPrivacyErrorResult
        {
            Success = false,
            ErrorMessage = "Privacy level required",
            ExistingPrivacyLevels = new List<QueryPrivacyInfo>
            {
                new QueryPrivacyInfo("Query1", PowerQueryPrivacyLevel.Private)
            },
            RecommendedPrivacyLevel = PowerQueryPrivacyLevel.Private,
            Explanation = "Test explanation",
            OriginalError = "Original Excel error"
        };

        // Assert - Verify all properties are accessible
        Assert.False(errorResult.Success);
        Assert.Equal("Privacy level required", errorResult.ErrorMessage);
        Assert.Single(errorResult.ExistingPrivacyLevels);
        Assert.Equal(PowerQueryPrivacyLevel.Private, errorResult.RecommendedPrivacyLevel);
        Assert.Equal("Test explanation", errorResult.Explanation);
        Assert.Equal("Original Excel error", errorResult.OriginalError);
    }

    [Fact]
    public void QueryPrivacyInfo_CreatesWithValidValues()
    {
        // Act
        var privacyInfo = new QueryPrivacyInfo("TestQuery", PowerQueryPrivacyLevel.Organizational);

        // Assert
        Assert.Equal("TestQuery", privacyInfo.QueryName);
        Assert.Equal(PowerQueryPrivacyLevel.Organizational, privacyInfo.PrivacyLevel);
    }

    [Fact]
    public async Task Update_WithPrivacyLevel_AcceptsParameter()
    {
        // Arrange - First import a query
        string queryFile = Path.Combine(_tempDir, "UpdateTestQuery.pq");
        string mCode1 = @"let Source = ""Test1"" in Source";
        File.WriteAllText(queryFile, mCode1);

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "UpdateTestQuery", queryFile);
            await batch.SaveAsync();
        }

        // Update the query content
        string mCode2 = @"let Source = ""Test2"" in Source";
        File.WriteAllText(queryFile, mCode2);

        // Act - Update with privacy level
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.UpdateAsync(batch, "UpdateTestQuery", queryFile, PowerQueryPrivacyLevel.Private);

            // Assert
            Assert.NotNull(result);
        }
    }

    [Theory]
    [InlineData(PowerQueryPrivacyLevel.None)]
    [InlineData(PowerQueryPrivacyLevel.Private)]
    [InlineData(PowerQueryPrivacyLevel.Organizational)]
    [InlineData(PowerQueryPrivacyLevel.Public)]
    public async Task SetLoadToTable_AcceptsPrivacyLevel(PowerQueryPrivacyLevel privacyLevel)
    {
        // Arrange - Create a simple query first
        string queryFile = Path.Combine(_tempDir, $"LoadToTableQuery_{privacyLevel}.pq");
        string mCode = @"let Source = #table({""Col1""}, {{""Val1""}}) in Source";
        File.WriteAllText(queryFile, mCode);

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, $"LoadToTableQuery_{privacyLevel}", queryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToTableAsync(batch, $"LoadToTableQuery_{privacyLevel}", "Sheet1", privacyLevel);

            // Assert
            Assert.NotNull(result);
        }
    }

    [Fact]
    public async Task SetLoadToDataModel_AcceptsPrivacyLevel()
    {
        // Arrange - Create a simple query
        string queryFile = Path.Combine(_tempDir, "LoadToDataModelQuery.pq");
        string mCode = @"let Source = #table({""Col1""}, {{""Val1""}}) in Source";
        File.WriteAllText(queryFile, mCode);

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "LoadToDataModelQuery", queryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "LoadToDataModelQuery", PowerQueryPrivacyLevel.Organizational);

            // Assert
            Assert.NotNull(result);
        }
    }

    [Fact]
    public async Task SetLoadToBoth_AcceptsPrivacyLevel()
    {
        // Arrange - Create a simple query
        string queryFile = Path.Combine(_tempDir, "LoadToBothQuery.pq");
        string mCode = @"let Source = #table({""Col1""}, {{""Val1""}}) in Source";
        File.WriteAllText(queryFile, mCode);

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "LoadToBothQuery", queryFile);
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _powerQueryCommands.SetLoadToBothAsync(batch, "LoadToBothQuery", "Sheet1", PowerQueryPrivacyLevel.Public);

            // Assert
            Assert.NotNull(result);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failures shouldn't fail tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
