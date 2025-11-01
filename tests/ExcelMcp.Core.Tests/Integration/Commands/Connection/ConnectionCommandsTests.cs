using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Comprehensive integration tests for ConnectionCommands.
/// Tests all connection operations with batch API pattern.
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests(TempDirectoryFixture fixture) : IClassFixture<TempDirectoryFixture>
{
    private readonly ConnectionCommands _commands = new();
    private readonly string _tempDir = fixture.TempDir;
}
