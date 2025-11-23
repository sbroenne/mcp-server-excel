using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for CreateEmpty → Open → List workflow (the exact LLM scenario that was failing)
/// </summary>
public partial class FileCommandsTests
{
    [Fact]
    public void CreateEmpty_ThenOpenAndList_ReturnsSheet1()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ThenOpenAndList_ReturnsSheet1)}_{Guid.NewGuid():N}.xlsx");

        // Act 1: CreateEmpty (what LLM called first)
        _fileCommands.CreateEmpty(testFile);

        // Assert CreateEmpty succeeded
        Assert.True(System.IO.File.Exists(testFile), "File should exist after CreateEmpty");

        // Act 2: Open the file in a new session (what LLM called second)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act 3: List worksheets (what LLM called third)
        var sheetCommands = new SheetCommands();
        var listResult = sheetCommands.List(batch);

        // Assert List succeeded (this was failing before the fix)
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Worksheets);
        Assert.NotEmpty(listResult.Worksheets);

        // Assert Sheet1 exists (created by CreateEmpty)
        Assert.Single(listResult.Worksheets);
        Assert.Equal("Sheet1", listResult.Worksheets[0].Name);
        Assert.Equal(1, listResult.Worksheets[0].Index);
    }

    [Fact]
    public void CreateEmpty_ThenOpenAndListMultipleTimes_ConsistentResults()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ThenOpenAndListMultipleTimes_ConsistentResults)}_{Guid.NewGuid():N}.xlsx");

        // Act 1: CreateEmpty
        _fileCommands.CreateEmpty(testFile);

        // Act 2: Open and List multiple times to verify consistency
        var sheetCommands = new SheetCommands();

        for (int i = 0; i < 3; i++)
        {
            using var batch = ExcelSession.BeginBatch(testFile);
            var listResult = sheetCommands.List(batch);

            Assert.True(listResult.Success, $"List iteration {i} failed: {listResult.ErrorMessage}");
            Assert.Single(listResult.Worksheets);
            Assert.Equal("Sheet1", listResult.Worksheets[0].Name);
        }
    }

    [Fact]
    public void CreateEmpty_ThenOpenAndCreateMoreSheets_AllSheetsVisible()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ThenOpenAndCreateMoreSheets_AllSheetsVisible)}_{Guid.NewGuid():N}.xlsx");

        // Act 1: CreateEmpty
        _fileCommands.CreateEmpty(testFile);

        // Act 2: Open session and create multiple sheets (like the LLM tried to do)
        using var batch = ExcelSession.BeginBatch(testFile);
        var sheetCommands = new SheetCommands();

        var sheetsToCreate = new[] { "Dashboard", "Azure_IaaS_Analysis", "AVS_Sizing" };

        foreach (var sheetName in sheetsToCreate)
        {
            var createSheetResult = sheetCommands.Create(batch, sheetName);
            Assert.True(createSheetResult.Success, $"Failed to create sheet '{sheetName}': {createSheetResult.ErrorMessage}");
        }

        // Act 3: List all sheets
        var listResult = sheetCommands.List(batch);

        // Assert
        Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Worksheets);
        Assert.Equal(4, listResult.Worksheets.Count); // Sheet1 + 3 new sheets

        // Verify all expected sheets exist
        Assert.Contains(listResult.Worksheets, s => s.Name == "Sheet1");
        Assert.Contains(listResult.Worksheets, s => s.Name == "Dashboard");
        Assert.Contains(listResult.Worksheets, s => s.Name == "Azure_IaaS_Analysis");
        Assert.Contains(listResult.Worksheets, s => s.Name == "AVS_Sizing");
    }
}

