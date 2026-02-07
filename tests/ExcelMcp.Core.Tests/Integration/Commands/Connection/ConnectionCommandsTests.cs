using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Connection;

/// <summary>
/// Comprehensive integration tests for ConnectionCommands.
/// Tests all connection operations with batch API pattern.
/// Uses ConnectionTestsFixture for efficient test file creation.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connections")]
[Trait("RequiresExcel", "true")]
public partial class ConnectionCommandsTests(ConnectionTestsFixture fixture) : IClassFixture<ConnectionTestsFixture>
{
    private readonly ConnectionCommands _commands = new();
    private readonly ConnectionTestsFixture _fixture = fixture;
}




