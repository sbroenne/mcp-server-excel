using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Integration tests for Parameter Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public partial class NamedRangeCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly INamedRangeCommands _parameterCommands;
    private readonly string _tempDir;

    public NamedRangeCommandsTests(TempDirectoryFixture fixture)
    {
        _parameterCommands = new NamedRangeCommands();
        _tempDir = fixture.TempDir;
    }
}
