using Sbroenne.ExcelMcp.Core.Commands.QueryTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

/// <summary>
/// Comprehensive integration tests for QueryTableCommands.
/// Tests all QueryTable operations with batch API pattern.
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "QueryTable")]
[Trait("RequiresExcel", "true")]
public partial class QueryTableCommandsTests(TempDirectoryFixture fixture) : IClassFixture<TempDirectoryFixture>
{
    private readonly QueryTableCommands _commands = new();
    private readonly string _tempDir = fixture.TempDir;
}
