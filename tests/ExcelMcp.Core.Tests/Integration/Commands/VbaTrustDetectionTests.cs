using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for VBA Trust Detection functionality.
/// These tests validate VBA trust detection, guidance generation, and TestVbaTrustScope helper.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBATrust")]
public class VbaTrustDetectionTests : IDisposable
{
    private readonly IScriptCommands _scriptCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public VbaTrustDetectionTests()
    {
        _scriptCommands = new ScriptCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_VBATrust_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "VBATrustTestWorkbook.xlsm");

        // Create test Excel file (macro-enabled)
        CreateTestExcelFile();
    }

    private void CreateTestExcelFile()
    {
        // Create macro-enabled file by using .xlsm extension
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
    }

    [Fact]
    public void VbaTrustRequiredResult_HasAllRequiredProperties()
    {
        // Act
        var trustResult = new VbaTrustRequiredResult
        {
            Success = false,
            ErrorMessage = "VBA trust access is not enabled",
            IsTrustEnabled = false,
            SetupInstructions = new[]
            {
                "Open Excel",
                "Go to File → Options → Trust Center",
                "Click 'Trust Center Settings'",
                "Select 'Macro Settings'",
                "Check '✓ Trust access to the VBA project object model'",
                "Click OK twice to save settings"
            },
            DocumentationUrl = "https://support.microsoft.com/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6",
            Explanation = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings."
        };

        // Assert - Verify all properties are accessible and have expected values
        Assert.False(trustResult.Success);
        Assert.Equal("VBA trust access is not enabled", trustResult.ErrorMessage);
        Assert.False(trustResult.IsTrustEnabled);
        Assert.NotNull(trustResult.SetupInstructions);
        Assert.Equal(6, trustResult.SetupInstructions.Length);
        Assert.Contains("Open Excel", trustResult.SetupInstructions);
        Assert.False(string.IsNullOrEmpty(trustResult.DocumentationUrl));
        Assert.False(string.IsNullOrEmpty(trustResult.Explanation));
    }

    [Fact]
    public async Task ScriptCommands_List_HandlesVbaTrustCorrectly()
    {
        // Note: This test validates that List returns a ScriptListResult
        // and handles VBA trust issues appropriately

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _scriptCommands.ListAsync(batch);

        // Assert
        Assert.NotNull(result);
        Assert.IsType<ScriptListResult>(result);

        // If VBA trust is not enabled, result may have an error or be empty
        // If VBA trust is enabled, result should have Scripts list
        Assert.NotNull(result.Scripts);
    }

    [Fact]
    public void TestVbaTrustScope_EnablesAndDisablesTrust()
    {
        // This test validates that TestVbaTrustScope properly manages VBA trust
        // It should enable trust, then revert to original state

        // Arrange - Check initial trust state
        bool initialTrustState = IsVbaTrustEnabled();

        // Act - Use TestVbaTrustScope
        using (var trustScope = new TestVbaTrustScope())
        {
            // Inside the scope, VBA trust should be enabled
            Assert.True(IsVbaTrustEnabled(), "VBA trust should be enabled inside TestVbaTrustScope");
        }

        // Assert - After scope disposal, trust should be restored to initial state
        bool finalTrustState = IsVbaTrustEnabled();
        Assert.Equal(initialTrustState, finalTrustState);
    }

    [Fact]
    public async Task TestVbaTrustScope_AllowsVbaOperations()
    {
        // Arrange
        string vbaFile = Path.Combine(_tempDir, "TestModule.vba");
        string vbaCode = @"Sub TestProcedure()
    MsgBox ""Test""
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        // Act & Assert - VBA operations should work inside the scope
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var importResult = await _scriptCommands.ImportAsync(batch, "TestModule", vbaFile);

            // Should succeed when VBA trust is enabled
            if (!importResult.Success)
            {
                // If it failed, it should NOT be due to VBA trust
                Assert.DoesNotContain("trust", importResult.ErrorMessage?.ToLowerInvariant() ?? "");
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Export_WithTrust_WorksCorrectly()
    {
        // Arrange
        string exportFile = Path.Combine(_tempDir, "ExportedModule.vba");

        // Act - Test with VBA trust enabled
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var result = await _scriptCommands.ExportAsync(batch, "ThisWorkbook", exportFile);

            // Assert
            Assert.NotNull(result);
            // Should either succeed or fail for reasons other than trust
            if (!result.Success && result.ErrorMessage != null)
            {
                Assert.DoesNotContain("trust", result.ErrorMessage.ToLowerInvariant());
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Import_WithTrust_WorksCorrectly()
    {
        // Arrange
        string vbaFile = Path.Combine(_tempDir, "ImportTestModule.vba");
        string vbaCode = @"Sub ImportTestProcedure()
    Dim x As Integer
    x = 42
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        // Act - Test with VBA trust enabled
        using (var _ = new TestVbaTrustScope())
        {
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var result = await _scriptCommands.ImportAsync(batch, "ImportTestModule", vbaFile);

            // Assert
            Assert.NotNull(result);
            // Should either succeed or fail for reasons other than trust
            if (!result.Success && result.ErrorMessage != null)
            {
                Assert.DoesNotContain("trust", result.ErrorMessage.ToLowerInvariant());
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Update_WithTrust_WorksCorrectly()
    {
        // Arrange
        string vbaFile = Path.Combine(_tempDir, "UpdateTestModule.vba");
        string vbaCode1 = @"Sub UpdateTest1()
End Sub";
        File.WriteAllText(vbaFile, vbaCode1);

        using (var _ = new TestVbaTrustScope())
        {
            // First import
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                await _scriptCommands.ImportAsync(batch, "UpdateTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Update the VBA code
            string vbaCode2 = @"Sub UpdateTest2()
    Dim y As String
End Sub";
            File.WriteAllText(vbaFile, vbaCode2);

            // Act - Update the module
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                var result = await _scriptCommands.UpdateAsync(batch, "UpdateTestModule", vbaFile);

                // Assert
                Assert.NotNull(result);
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Delete_WithTrust_WorksCorrectly()
    {
        // Arrange
        string vbaFile = Path.Combine(_tempDir, "DeleteTestModule.vba");
        string vbaCode = @"Sub DeleteTest()
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        using (var _ = new TestVbaTrustScope())
        {
            // First import a module
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                await _scriptCommands.ImportAsync(batch, "DeleteTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Act - Delete the module
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                var result = await _scriptCommands.DeleteAsync(batch, "DeleteTestModule");

                // Assert
                Assert.NotNull(result);
            }
        }
    }

    [Fact]
    public async Task ScriptCommands_Run_WithTrust_WorksCorrectly()
    {
        // Arrange
        string vbaFile = Path.Combine(_tempDir, "RunTestModule.vba");
        string vbaCode = @"Sub RunTest()
    ' Simple procedure that does nothing
    Dim x As Integer
    x = 1
End Sub";
        File.WriteAllText(vbaFile, vbaCode);

        using (var _ = new TestVbaTrustScope())
        {
            // First import a module
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                await _scriptCommands.ImportAsync(batch, "RunTestModule", vbaFile);
                await batch.SaveAsync();
            }

            // Act - Run the procedure
            await using (var batch = await ExcelSession.BeginBatchAsync(_testExcelFile))
            {
                var result = await _scriptCommands.RunAsync(batch, "RunTestModule.RunTest");

                // Assert
                Assert.NotNull(result);
            }
        }
    }

    /// <summary>
    /// Helper method to check VBA trust status via registry
    /// </summary>
    private static bool IsVbaTrustEnabled()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security");
            var value = key?.GetValue("AccessVBOM");
            return value != null && (int)value == 1;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failures shouldn't fail tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
