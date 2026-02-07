using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Integration tests for RangeCommands - main partial class with shared fixture.
/// Uses RangeTestsFixture to create ONE file shared across all tests in this class.
/// Each test creates its own sheet within that file for isolation.
/// Other test methods are in partial files: Values.cs, Formulas.cs, Editing.cs, Search.cs, Discovery.cs, Hyperlinks.cs
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public partial class RangeCommandsTests : IClassFixture<RangeTestsFixture>
{
    private readonly ITestOutputHelper _output;
    private readonly RangeCommands _commands;
    private readonly RangeTestsFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the test class with shared fixture
    /// </summary>
    public RangeCommandsTests(ITestOutputHelper output, RangeTestsFixture fixture)
    {
        _output = output;
        _commands = new RangeCommands();
        _fixture = fixture;
    }
}




