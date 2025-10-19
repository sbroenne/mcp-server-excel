using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.CLI.Tests.Commands;

/// <summary>
/// Integration tests for worksheet operations using Excel COM automation.
/// These tests require Excel installation and validate sheet manipulation commands.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Worksheets")]
public class SheetCommandsTests
{
    private readonly SheetCommands _sheetCommands;

    public SheetCommandsTests()
    {
        _sheetCommands = new SheetCommands();
    }

    [Theory]
    [InlineData("sheet-list")]
    [InlineData("sheet-create", "test.xlsx")]
    [InlineData("sheet-rename", "test.xlsx", "Sheet1")]
    [InlineData("sheet-delete", "test.xlsx")]
    [InlineData("sheet-clear", "test.xlsx")]
    public void Commands_WithInsufficientArgs_ReturnsError(params string[] args)
    {
        // Act & Assert based on command
        int result = args[0] switch
        {
            "sheet-list" => _sheetCommands.List(args),
            "sheet-create" => _sheetCommands.Create(args),
            "sheet-rename" => _sheetCommands.Rename(args),
            "sheet-delete" => _sheetCommands.Delete(args),
            "sheet-clear" => _sheetCommands.Clear(args),
            _ => throw new ArgumentException($"Unknown command: {args[0]}")
        };

        Assert.Equal(1, result);
    }

    [Fact]
    public void List_WithNonExistentFile_ReturnsError()
    {
        // Arrange
        string[] args = { "sheet-list", "nonexistent.xlsx" };

        // Act
        int result = _sheetCommands.List(args);

        // Assert
        Assert.Equal(1, result);
    }

    [Theory]
    [InlineData("sheet-create", "nonexistent.xlsx", "NewSheet")]
    [InlineData("sheet-rename", "nonexistent.xlsx", "Old", "New")]
    [InlineData("sheet-delete", "nonexistent.xlsx", "Sheet1")]
    [InlineData("sheet-clear", "nonexistent.xlsx", "Sheet1")]
    public void Commands_WithNonExistentFile_ReturnsError(params string[] args)
    {
        // Act
        int result = args[0] switch
        {
            "sheet-create" => _sheetCommands.Create(args),
            "sheet-rename" => _sheetCommands.Rename(args),
            "sheet-delete" => _sheetCommands.Delete(args),
            "sheet-clear" => _sheetCommands.Clear(args),
            _ => throw new ArgumentException($"Unknown command: {args[0]}")
        };

        // Assert
        Assert.Equal(1, result);
    }
}
