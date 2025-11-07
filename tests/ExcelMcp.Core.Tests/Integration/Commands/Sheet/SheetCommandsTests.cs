using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for Sheet lifecycle operations.
/// These tests require Excel installation and validate Core worksheet lifecycle management.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// Data operations (read, write, clear) moved to RangeCommandsTests.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
public partial class SheetCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly SheetCommands _sheetCommands;
    private readonly string _tempDir;

    /// <summary>
    /// Initializes a new instance of the <see cref="SheetCommandsTests"/> class.
    /// </summary>
    public SheetCommandsTests(TempDirectoryFixture fixture)
    {
        _sheetCommands = new SheetCommands();
        _tempDir = fixture.TempDir;
    }
}
