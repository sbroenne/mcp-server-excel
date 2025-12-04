using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

/// <summary>
/// Integration tests for Script (VBA) Core operations.
/// These tests require Excel installation and VBA trust enabled.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBA")]
public partial class VbaCommandsTests : IClassFixture<VbaTestsFixture>
{
    private readonly VbaCommands _scriptCommands;
    private readonly VbaTestsFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the <see cref="VbaCommandsTests"/> class.
    /// </summary>
    public VbaCommandsTests(VbaTestsFixture fixture)
    {
        _scriptCommands = new VbaCommands();
        _fixture = fixture;
    }
}
