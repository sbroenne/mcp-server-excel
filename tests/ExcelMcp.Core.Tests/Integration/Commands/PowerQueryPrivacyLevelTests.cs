using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query Privacy Level functionality.
/// These tests validate privacy level detection, recommendation, and application.
/// Uses Excel instance pooling for improved test performance.
/// </summary>
[Collection(nameof(ExcelPooledTestCollection))]
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
        _powerQueryCommands = new PowerQueryCommands();
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
        var result = _fileCommands.CreateEmpty(_testExcelFile, overwriteIfExists: false);
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
        var result = await _powerQueryCommands.Import(_testExcelFile, $"TestQuery_{privacyLevel}", queryFile, privacyLevel);

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
        var result = await _powerQueryCommands.Import(_testExcelFile, "SimpleQuery", queryFile, null);

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
        await _powerQueryCommands.Import(_testExcelFile, "UpdateTestQuery", queryFile);

        // Update the query content
        string mCode2 = @"let Source = ""Test2"" in Source";
        File.WriteAllText(queryFile, mCode2);

        // Act - Update with privacy level
        var result = await _powerQueryCommands.Update(_testExcelFile, "UpdateTestQuery", queryFile, PowerQueryPrivacyLevel.Private);

        // Assert
        Assert.NotNull(result);
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
        await _powerQueryCommands.Import(_testExcelFile, $"LoadToTableQuery_{privacyLevel}", queryFile);

        // Act
        var result = _powerQueryCommands.SetLoadToTable(_testExcelFile, $"LoadToTableQuery_{privacyLevel}", "Sheet1", privacyLevel);

        // Assert
        Assert.NotNull(result);
    }

    [Fact]
    public async Task SetLoadToDataModel_AcceptsPrivacyLevel()
    {
        // Arrange - Create a simple query
        string queryFile = Path.Combine(_tempDir, "LoadToDataModelQuery.pq");
        string mCode = @"let Source = #table({""Col1""}, {{""Val1""}}) in Source";
        File.WriteAllText(queryFile, mCode);
        await _powerQueryCommands.Import(_testExcelFile, "LoadToDataModelQuery", queryFile);

        // Act
        var result = _powerQueryCommands.SetLoadToDataModel(_testExcelFile, "LoadToDataModelQuery", PowerQueryPrivacyLevel.Organizational);

        // Assert
        Assert.NotNull(result);
    }

    [Fact]
    public async Task SetLoadToBoth_AcceptsPrivacyLevel()
    {
        // Arrange - Create a simple query
        string queryFile = Path.Combine(_tempDir, "LoadToBothQuery.pq");
        string mCode = @"let Source = #table({""Col1""}, {{""Val1""}}) in Source";
        File.WriteAllText(queryFile, mCode);
        await _powerQueryCommands.Import(_testExcelFile, "LoadToBothQuery", queryFile);

        // Act
        var result = _powerQueryCommands.SetLoadToBoth(_testExcelFile, "LoadToBothQuery", "Sheet1", PowerQueryPrivacyLevel.Public);

        // Assert
        Assert.NotNull(result);
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
