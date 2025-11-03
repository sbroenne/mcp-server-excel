using System.IO;
using Sbroenne.ExcelMcp.Core.Commands;
using System.IO;
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
public partial class VbaCommandsTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IVbaCommands _scriptCommands;
    private readonly string _tempDir;

    public VbaCommandsTests(TempDirectoryFixture fixture)
    {
        _scriptCommands = new VbaCommands();
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

        var vbaFile = Path.Join(_tempDir, fileName);
        System.IO.File.WriteAllText(vbaFile, vbaCode);
        return vbaFile;
    }
}
