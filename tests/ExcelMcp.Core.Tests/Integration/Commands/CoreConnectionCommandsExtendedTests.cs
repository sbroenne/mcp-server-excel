using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using System.IO;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Extended integration tests for Connection Core operations.
/// These tests cover additional scenarios, edge cases, and validation logic.
/// Tests use Core commands directly (not through CLI wrapper).
/// Uses Excel instance pooling for improved test performance.
/// </summary>
[Collection(nameof(ExcelPooledTestCollection))]
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Connections")]
[Trait("Speed", "Medium")]
public class CoreConnectionCommandsExtendedTests : IDisposable
{
    private readonly ConnectionCommands _commands;
    private readonly FileCommands _fileCommands;
    private readonly string _testDataDir;
    private readonly List<string> _filesToCleanup;
    private bool _disposed;

    public CoreConnectionCommandsExtendedTests()
    {
        _commands = new ConnectionCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _testDataDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_ConnectionExtTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDataDir);

        _filesToCleanup = new List<string>();
    }

    private string CreateTestWorkbook()
    {
        string filePath = Path.Combine(_testDataDir, $"test_{Guid.NewGuid():N}.xlsx");
        var result = _fileCommands.CreateEmpty(filePath, overwriteIfExists: false);

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test workbook: {result.ErrorMessage}");
        }

        _filesToCleanup.Add(filePath);
        return filePath;
    }

    public void Dispose()
    {
        if (_disposed) return;

        // Clean up test files
        foreach (var file in _filesToCleanup)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        // Clean up test directory
        try
        {
            if (Directory.Exists(_testDataDir))
            {
                Directory.Delete(_testDataDir, recursive: true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    #region List Edge Cases

    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.List(filePath);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.List(filePath);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Connections);
    }

    #endregion

    #region View Edge Cases

    [Fact]
    public void View_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.View(filePath, "NonexistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void View_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.View(filePath, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void View_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.View(filePath, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Export Edge Cases

    [Fact]
    public void Export_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "export.json");

        // Act
        var result = _commands.Export(filePath, "NonexistentConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Export_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "export.json");

        // Act
        var result = _commands.Export(filePath, null!, jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Export_NullJsonPath_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Export(filePath, "AnyConnection", null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Update Edge Cases

    [Fact]
    public void Update_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "update.json");

        // Create minimal JSON
        File.WriteAllText(jsonPath, "{\"Name\":\"NonexistentConnection\",\"Type\":\"OLEDB\"}");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Update(filePath, "NonexistentConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Update_MalformedJson_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "malformed.json");

        // Create invalid JSON
        File.WriteAllText(jsonPath, "{invalid json content");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Update(filePath, "AnyConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Update_NonexistentJsonFile_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "nonexistent.json");

        // Act
        var result = _commands.Update(filePath, "AnyConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Update_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "update.json");
        File.WriteAllText(jsonPath, "{}");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Update(filePath, null!, jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Refresh Edge Cases

    [Fact]
    public void Refresh_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Refresh(filePath, "NonexistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refresh_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Refresh(filePath, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Refresh_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Refresh(filePath, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Delete Edge Cases

    [Fact]
    public void Delete_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Delete(filePath, "NonexistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Delete_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Delete(filePath, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Delete_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Delete(filePath, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region LoadTo Edge Cases

    [Fact]
    public void LoadTo_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.LoadTo(filePath, "NonexistentConnection", "Sheet1");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void LoadTo_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.LoadTo(filePath, null!, "Sheet1");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void LoadTo_NullSheetName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.LoadTo(filePath, "AnyConnection", null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void LoadTo_EmptySheetName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.LoadTo(filePath, "AnyConnection", "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region GetProperties Edge Cases

    [Fact]
    public void GetProperties_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.GetProperties(filePath, "NonexistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetProperties_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.GetProperties(filePath, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void GetProperties_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.GetProperties(filePath, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region SetProperties Edge Cases

    [Fact]
    public void SetProperties_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.SetProperties(filePath, "NonexistentConnection", backgroundQuery: true);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetProperties_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.SetProperties(filePath, null!, backgroundQuery: true);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void SetProperties_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.SetProperties(filePath, "", backgroundQuery: true);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Test Edge Cases

    [Fact]
    public void Test_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Test(filePath, "NonexistentConnection");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Test(filePath, null!);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Test_EmptyConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.Test(filePath, "");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Import Edge Cases

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Medium")]
    [Trait("Feature", "Connections")]
    [Trait("RequiresExcel", "true")]
    public void Import_MissingConnectionString_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "import.json");
        File.WriteAllText(jsonPath, "{\"Name\":\"Test\",\"Type\":\"OLEDB\"}");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Import(filePath, "TestConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("ConnectionString is required", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Import_NullConnectionName_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "import.json");
        File.WriteAllText(jsonPath, "{}");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Import(filePath, null!, jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Import_NonexistentJsonFile_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "nonexistent.json");

        // Act
        var result = _commands.Import(filePath, "TestConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
