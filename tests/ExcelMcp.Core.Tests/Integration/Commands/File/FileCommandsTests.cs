using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Integration tests for File Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
///
/// WHAT LLMs NEED TO KNOW:
/// 1. TestFile returns metadata (Exists, IsValid, Message) without Success flag
/// 2. File creation uses SessionManager.CreateSessionForNewFile (create action)
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test Excel COM file operations and Result objects
/// - ✅ Test business rules (valid extensions, file metadata)
/// - ❌ DO NOT test CLI argument parsing (CLI's responsibility)
/// - ❌ DO NOT test JSON serialization (MCP Server's responsibility)
/// - ❌ DO NOT test infrastructure (paths, directories, OS validation)
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("RequiresExcel", "true")]
public partial class FileCommandsTests : IClassFixture<FileTestsFixture>
{
    // Performance: use concrete type to satisfy CA1859 (test code, not API surface)
    private readonly FileCommands _fileCommands;
    private readonly FileTestsFixture _fixture;

    public FileCommandsTests(FileTestsFixture fixture)
    {
        _fileCommands = new FileCommands();
        _fixture = fixture;
    }
}




