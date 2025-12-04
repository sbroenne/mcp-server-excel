using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for Sheet lifecycle operations.
/// These tests require Excel installation and validate Core worksheet lifecycle management.
/// Tests use Core commands directly (not through CLI wrapper).
/// Single-workbook tests share one Excel file with unique sheets for isolation.
/// Cross-workbook tests (CopyToWorkbook, MoveToWorkbook) use their own file pairs.
/// Data operations (read, write, clear) moved to RangeCommandsTests.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
public partial class SheetCommandsTests : IClassFixture<SheetTestsFixture>
{
    private readonly SheetCommands _sheetCommands;
    private readonly SheetTestsFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the <see cref="SheetCommandsTests"/> class.
    /// </summary>
    public SheetCommandsTests(SheetTestsFixture fixture)
    {
        _sheetCommands = new SheetCommands();
        _fixture = fixture;
    }
}
