using Sbroenne.ExcelMcp.ComInterop.Session;
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
public class ParameterCommandsTests : IDisposable
{
    private readonly IParameterCommands _parameterCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public ParameterCommandsTests()
    {
        _parameterCommands = new ParameterCommands();
        _fileCommands = new FileCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_ParamTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsx");

        // Create test Excel file
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public async Task List_WithValidFile_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _parameterCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Parameters);
    }

    [Fact]
    public async Task Create_WithValidParameter_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _parameterCommands.CreateAsync(batch, "TestParam", "Sheet1!A1");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify the parameter was actually created by listing parameters
        var listResult = await _parameterCommands.ListAsync(batch);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.Contains(listResult.Parameters, p => p.Name == "TestParam");
    }

    [Fact]
    public async Task Create_ThenList_ShowsCreatedParameter()
    {
        // Arrange
        string paramName = "IntegrationTestParam";

        // Act - Create parameter
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var createResult = await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!B2");
            Assert.True(createResult.Success);
            await batch.SaveAsync();
        }

        // Act - List parameters
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var listResult = await _parameterCommands.ListAsync(batch);

            // Assert
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Parameters, p => p.Name == paramName);
        }
    }

    [Fact]
    public async Task Set_WithValidParameter_ReturnsSuccess()
    {
        // Arrange - Use unique parameter name to avoid conflicts
        string paramName = "SetTestParam_" + Guid.NewGuid().ToString("N")[..8];

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var createResult = await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!C1");
            Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _parameterCommands.SetAsync(batch, paramName, "TestValue");
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success, $"Failed to set parameter: {result.ErrorMessage}");

            // Verify the parameter value was actually set by reading it back
            var getResult = await _parameterCommands.GetAsync(batch, paramName);
            Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
            Assert.Equal("TestValue", getResult.Value?.ToString());
        }
    }

    [Fact]
    public async Task Set_ThenGet_ReturnsSetValue()
    {
        // Arrange - Use unique parameter name to avoid conflicts
        string paramName = "GetSetParam_" + Guid.NewGuid().ToString("N")[..8];
        string testValue = "Integration Test Value";

        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var createResult = await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!D1");
            Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - Set value
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var setResult = await _parameterCommands.SetAsync(batch, paramName, testValue);
            Assert.True(setResult.Success, $"Failed to set parameter: {setResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - Get value
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var getResult = await _parameterCommands.GetAsync(batch, paramName);

            // Assert
            Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
            Assert.Equal(testValue, getResult.Value?.ToString());
        }
    }

    [Fact]
    public async Task Delete_WithValidParameter_ReturnsSuccess()
    {
        // Arrange
        string paramName = "DeleteTestParam";
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!E1");
            await batch.SaveAsync();
        }

        // Act
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var result = await _parameterCommands.DeleteAsync(batch, paramName);
            await batch.SaveAsync();

            // Assert
            Assert.True(result.Success);

            // Verify the parameter was actually deleted by checking it's not in the list
            var listResult = await _parameterCommands.ListAsync(batch);
            Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
            Assert.DoesNotContain(listResult.Parameters, p => p.Name == paramName);
        }
    }

    [Fact]
    public async Task Delete_ThenList_DoesNotShowDeletedParameter()
    {
        // Arrange
        string paramName = "DeletedParam";
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            await _parameterCommands.CreateAsync(batch, paramName, "Sheet1!F1");
            await batch.SaveAsync();
        }

        // Act - Delete parameter
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var deleteResult = await _parameterCommands.DeleteAsync(batch, paramName);
            Assert.True(deleteResult.Success);
            await batch.SaveAsync();
        }

        // Act - List parameters
        await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
        {
            var listResult = await _parameterCommands.ListAsync(batch);

            // Assert
            Assert.True(listResult.Success);
            Assert.DoesNotContain(listResult.Parameters, p => p.Name == paramName);
        }
    }

    [Fact]
    public async Task List_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
        });
    }

    [Fact]
    public async Task Get_WithNonExistentParameter_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _parameterCommands.GetAsync(batch, "NonExistentParam");

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
