using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Parameter Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public class CoreParameterCommandsTests : IDisposable
{
    private readonly IParameterCommands _parameterCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public CoreParameterCommandsTests()
    {
        _parameterCommands = new ParameterCommands();
        _fileCommands = new FileCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_ParamTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");

        // Create test Excel file
        var result = _fileCommands.CreateEmpty(_testExcelFile);
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public void List_WithValidFile_ReturnsSuccess()
    {
        // Act
        var result = _parameterCommands.List(_testExcelFile);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
    }

    [Fact]
    public void Create_WithValidParameter_ReturnsSuccess()
    {
        // Act
        var result = _parameterCommands.Create(_testExcelFile, "TestParam", "Sheet1!A1");

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Create_ThenList_ShowsCreatedParameter()
    {
        // Arrange
        string paramName = "IntegrationTestParam";

        // Act
        var createResult = _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!B2");
        var listResult = _parameterCommands.List(_testExcelFile);

        // Assert
        Assert.True(createResult.Success);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Parameters, p => p.Name == paramName);
    }

    [Fact]
    public void Set_WithValidParameter_ReturnsSuccess()
    {
        // Arrange - Use unique parameter name to avoid conflicts
        string paramName = "SetTestParam_" + Guid.NewGuid().ToString("N")[..8];
        var createResult = _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!C1");

        // Ensure parameter was created successfully
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        // Act
        var result = _parameterCommands.Set(_testExcelFile, paramName, "TestValue");

        // Assert
        Assert.True(result.Success, $"Failed to set parameter: {result.ErrorMessage}");
    }

    [Fact]
    public void Set_ThenGet_ReturnsSetValue()
    {
        // Arrange - Use unique parameter name to avoid conflicts
        string paramName = "GetSetParam_" + Guid.NewGuid().ToString("N")[..8];
        string testValue = "Integration Test Value";
        var createResult = _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!D1");

        // Ensure parameter was created successfully
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        // Act
        var setResult = _parameterCommands.Set(_testExcelFile, paramName, testValue);
        var getResult = _parameterCommands.Get(_testExcelFile, paramName);

        // Assert
        Assert.True(setResult.Success, $"Failed to set parameter: {setResult.ErrorMessage}");
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal(testValue, getResult.Value?.ToString());
    }

    [Fact]
    public void Delete_WithValidParameter_ReturnsSuccess()
    {
        // Arrange
        string paramName = "DeleteTestParam";
        _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!E1");

        // Act
        var result = _parameterCommands.Delete(_testExcelFile, paramName);

        // Assert
        Assert.True(result.Success);
    }

    [Fact]
    public void Delete_ThenList_DoesNotShowDeletedParameter()
    {
        // Arrange
        string paramName = "DeletedParam";
        _parameterCommands.Create(_testExcelFile, paramName, "Sheet1!F1");

        // Act
        var deleteResult = _parameterCommands.Delete(_testExcelFile, paramName);
        var listResult = _parameterCommands.List(_testExcelFile);

        // Assert
        Assert.True(deleteResult.Success);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Parameters, p => p.Name == paramName);
    }

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = _parameterCommands.List("nonexistent.xlsx");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void Get_WithNonExistentParameter_ReturnsError()
    {
        // Act
        var result = _parameterCommands.Get(_testExcelFile, "NonExistentParam");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }

        GC.SuppressFinalize(this);
    }
}
