using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for worksheet shape operations.
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public void AddShape_InsertsShapeIntoWorksheetAndCanBeCounted()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Shapes_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        var addResult = _sheetCommands.AddShape(batch, sheetName, "A1");
        Assert.True(addResult.Success, $"Expected shape add to succeed but got error: {addResult.ErrorMessage}");

        var countResult = _sheetCommands.GetShapeCount(batch, sheetName);
        Assert.True(countResult.Success);
        Assert.True(countResult.ShapeCount > 0);
    }
}
