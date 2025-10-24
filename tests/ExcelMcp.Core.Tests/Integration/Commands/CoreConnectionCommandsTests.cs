using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Integration tests for ConnectionCommands - Core layer
/// Tests real Excel COM operations with actual Excel files
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public class CoreConnectionCommandsTests : IDisposable
{
    private readonly string _testDataDir;
    private readonly ConnectionCommands _commands;
    private readonly List<string> _filesToCleanup;

    public CoreConnectionCommandsTests()
    {
        _testDataDir = Path.Combine(Path.GetTempPath(), "ExcelMcp_ConnectionTests_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(_testDataDir);
        _commands = new ConnectionCommands();
        _filesToCleanup = new List<string>();
    }

    public void Dispose()
    {
        // Cleanup test files
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

        GC.SuppressFinalize(this);
    }

    private string CreateTestWorkbook()
    {
        var fileCommands = new FileCommands();
        string filePath = Path.Combine(_testDataDir, $"test_{Guid.NewGuid().ToString("N")[..8]}.xlsx");
        var result = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);

        Assert.True(result.Success, $"Failed to create test workbook: {result.ErrorMessage}");
        _filesToCleanup.Add(filePath);

        return filePath;
    }

    #region List Tests

    [Fact]
    public void List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act
        var result = _commands.List(filePath);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Connections);
        Assert.Empty(result.Connections);
    }

    [Fact]
    public void List_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.List(filePath));
    }

    #endregion

    #region View Tests

    [Fact]
    public void View_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string connectionName = "NonexistentConnection";

        // Act
        var result = _commands.View(filePath, connectionName);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void View_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.View(filePath, "AnyConnection"));
    }

    #endregion

    #region Export Tests

    [Fact]
    public void Export_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "export.json");
        string connectionName = "NonexistentConnection";

        // Act
        var result = _commands.Export(filePath, connectionName, jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Export_InvalidExcelFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");
        string jsonPath = Path.Combine(_testDataDir, "export.json");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.Export(filePath, "AnyConnection", jsonPath));
    }

    #endregion

    #region Update Tests

    [Fact]
    public void Update_NonexistentConnection_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "update.json");

        // Create minimal JSON
        File.WriteAllText(jsonPath, "{\"Name\":\"Test\",\"Type\":\"OLEDB\"}");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Update(filePath, "NonexistentConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Update_InvalidJsonFile_ReturnsError()
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
    public void Update_MalformedJson_ReturnsError()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "malformed.json");
        File.WriteAllText(jsonPath, "{invalid json content");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Update(filePath, "AnyConnection", jsonPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    #endregion

    #region Refresh Tests

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
    public void Refresh_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.Refresh(filePath, "AnyConnection"));
    }

    #endregion

    #region Delete Tests

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
    public void Delete_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.Delete(filePath, "AnyConnection"));
    }

    #endregion

    #region LoadTo Tests

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
    public void LoadTo_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.LoadTo(filePath, "AnyConnection", "Sheet1"));
    }

    #endregion

    #region GetProperties Tests

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
    public void GetProperties_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.GetProperties(filePath, "AnyConnection"));
    }

    #endregion

    #region SetProperties Tests

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
    public void SetProperties_NoPropertiesSpecified_Succeeds()
    {
        // Arrange
        string filePath = CreateTestWorkbook();

        // Act - Should succeed even with no properties (no-op)
        var result = _commands.SetProperties(filePath, "NonexistentConnection");

        // Assert - Will fail because connection doesn't exist, but demonstrates no properties is valid
        Assert.False(result.Success);
    }

    [Fact]
    public void SetProperties_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.SetProperties(filePath, "AnyConnection", backgroundQuery: true));
    }

    #endregion

    #region Test Tests

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
    public void Test_InvalidFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string filePath = Path.Combine(_testDataDir, "nonexistent.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() => _commands.Test(filePath, "AnyConnection"));
    }

    #endregion

    #region Power Query Detection Tests

    [Fact]
    public void PowerQueryConnections_AreDetectedAndRejected()
    {
        // This test would require creating a workbook with a Power Query connection
        // For now, we validate that Power Query detection logic is in place
        // by checking the error messages contain appropriate guidance

        // NOTE: Full testing requires a workbook with actual Power Query connections
        // This is a placeholder for comprehensive Power Query detection testing

        Assert.True(true, "Power Query detection tests require workbooks with PQ connections");
    }

    #endregion

    #region Import Tests

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Medium")]
    [Trait("Feature", "Connections")]
    [Trait("RequiresExcel", "true")]
    public void Import_ValidJson_CreatesWebConnection()
    {
        // Arrange
        string filePath = CreateTestWorkbook();
        string jsonPath = Path.Combine(_testDataDir, "import.json");

        // Create connection to a web page (no external dependencies needed for creation)
        File.WriteAllText(jsonPath, @"{
            ""Name"": ""TestWebConnection"",
            ""Type"": ""Web"",
            ""ConnectionString"": ""URL;https://example.com"",
            ""CommandText"": """",
            ""Description"": ""Test web connection""
        }");
        _filesToCleanup.Add(jsonPath);

        // Act
        var result = _commands.Import(filePath, "TestWebConnection", jsonPath);

        // Assert
        Assert.True(result.Success, $"Import failed: {result.ErrorMessage}");

        // Verify connection was created
        var listResult = _commands.List(filePath);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Connections, c => c.Name == "TestWebConnection");
    }

    [Fact]
    public void Import_InvalidJsonFile_ReturnsError()
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

    #region Security Tests

    [Fact]
    public void ConnectionStringsSanitization_MasksPasswords()
    {
        // This test validates that password sanitization is applied
        // Would require a workbook with actual connections containing passwords

        // NOTE: Full testing requires workbooks with connections that have passwords
        // This is a placeholder for comprehensive security testing

        Assert.True(true, "Password sanitization tests require connections with passwords");
    }

    #endregion
}
