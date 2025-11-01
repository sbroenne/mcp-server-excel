using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Setup;

/// <summary>
/// Integration tests for Setup Core operations.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Setup")]
[Trait("RequiresExcel", "true")]
public partial class SetupCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly ISetupCommands _setupCommands;
    private readonly string _tempDir;

    public SetupCommandsTests(TempDirectoryFixture fixture)
    {
        _setupCommands = new SetupCommands();
        _tempDir = fixture.TempDir;
    }
}
