using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Parameter;

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
public partial class ParameterCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IParameterCommands _parameterCommands;
    private readonly string _tempDir;

    public ParameterCommandsTests(TempDirectoryFixture fixture)
    {
        _parameterCommands = new ParameterCommands();
        _tempDir = fixture.TempDir;
    }
}
