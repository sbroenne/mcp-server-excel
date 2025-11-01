using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Script;

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
public partial class ScriptCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IScriptCommands _scriptCommands;
    private readonly ISetupCommands _setupCommands;
    private readonly string _tempDir;

    public ScriptCommandsTests(TempDirectoryFixture fixture)
    {
        _scriptCommands = new ScriptCommands();
        _setupCommands = new SetupCommands();
        _tempDir = fixture.TempDir;
    }

    /// <summary>
    /// Helper to create test VBA file
    /// </summary>
    private string CreateTestVbaFile(string fileName = "TestModule.vba")
    {
        string vbaCode = @"Option Explicit

Public Function TestFunction() As String
    TestFunction = ""Hello from VBA""
End Function

Public Sub TestSubroutine()
    MsgBox ""Test VBA""
End Sub";

        var vbaFile = Path.Combine(_tempDir, fileName);
        File.WriteAllText(vbaFile, vbaCode);
        return vbaFile;
    }
}
