using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Integration tests for RangeCommands - main partial class with shared fixture
/// Other test methods are in partial files: Values.cs, Formulas.cs, Editing.cs, Search.cs, Discovery.cs, Hyperlinks.cs
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public partial class RangeCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly ITestOutputHelper _output;
    private readonly RangeCommands _commands;
    private readonly string _tempDir;
    /// <inheritdoc/>

    public RangeCommandsTests(ITestOutputHelper output, TempDirectoryFixture fixture)
    {
        _output = output;
        _commands = new RangeCommands();
        _tempDir = fixture.TempDir;
    }
}
