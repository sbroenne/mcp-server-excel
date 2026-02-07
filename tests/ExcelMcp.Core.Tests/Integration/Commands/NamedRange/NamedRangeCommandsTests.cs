using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Integration tests for Parameter Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Uses shared fixture for test isolation - each test uses unique named range names.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public partial class NamedRangeCommandsTests : IClassFixture<NamedRangeTestsFixture>
{
    private readonly NamedRangeCommands _parameterCommands;
    private readonly NamedRangeTestsFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the <see cref="NamedRangeCommandsTests"/> class.
    /// </summary>
    public NamedRangeCommandsTests(NamedRangeTestsFixture fixture)
    {
        _parameterCommands = new NamedRangeCommands();
        _fixture = fixture;
    }
}




