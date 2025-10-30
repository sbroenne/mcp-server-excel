using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.PowerQuery;

/// <summary>
/// Regression tests for Power Query QueryTable persistence.
///
/// HISTORY (2025-10-29): These tests were created to debug why ImportAsync with loadToWorksheet=true
/// created QueryTables that existed in memory but were lost after file reopen.
///
/// ROOT CAUSE DISCOVERED: SetLoadToTableAsync was using RefreshAll() which is ASYNCHRONOUS
/// and doesn't properly persist individual QueryTables to disk. Individual queryTable.Refresh(false)
/// is SYNCHRONOUS and required for proper persistence.
///
/// FIX IMPLEMENTED: Changed RefreshImmediately=true in SetLoadToTableAsync (line 1847 in PowerQueryCommands.cs)
/// and removed RefreshAll() block. This makes CreateQueryTable call queryTable.Refresh(false) synchronously,
/// following Microsoft's documented VBA pattern: Create → Refresh(False) → Save
///
/// REGRESSION VALUE: These tests now verify that QueryTable persistence continues to work correctly
/// and catch any future regressions in the QueryTable creation/refresh logic.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "true")]
public class PowerQueryLoadConfigDebugTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _tempDir;
    private readonly IPowerQueryCommands _commands;
    private readonly IFileCommands _fileCommands;

    public PowerQueryLoadConfigDebugTests(ITestOutputHelper output)
    {
        _output = output;
        _tempDir = Path.Combine(Path.GetTempPath(), $"PQ_LoadConfig_Debug_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        var dataModelCommands = new DataModelCommands();
        _commands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failure is non-critical
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Regression test: Verifies that QueryTable persists properly after Formula changes.
    ///
    /// HISTORICAL CONTEXT: This test revealed the RefreshAll() bug on 2025-10-29.
    /// QueryTables created with RefreshAll() existed in memory but were lost after file reopen.
    ///
    /// CURRENT STATUS: Bug fixed - QueryTables now persist correctly using individual Refresh(false).
    /// This test now serves as regression protection against future QueryTable persistence issues.
    /// </summary>
    [Fact(Timeout = 60000)]
    public async Task Debug_ManualFormulaChange_ChecksIfQueryTableSurvives()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, "formula-change-test.xlsx");
        var queryFile = Path.Combine(_tempDir, "test-query.pq");
        var mCode1 = @"let Source = #table({""Col1""}, {{""Data1""}}) in Source";
        var mCode2 = @"let Source = #table({""Col1""}, {{""DataModified""}}) in Source";

        await File.WriteAllTextAsync(queryFile, mCode1);
        await _fileCommands.CreateEmptyAsync(testFile);

        _output.WriteLine($"Created test file: {testFile}");

        // Step 1: Import with auto-load
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            _output.WriteLine("\nSTEP 1: Import with auto-load (loadToWorksheet=true)...");
            var importResult = await _commands.ImportAsync(batch, "TestQuery", queryFile, privacyLevel: null, loadToWorksheet: true);

            _output.WriteLine($"Import Success: {importResult.Success}");
            _output.WriteLine($"Import ErrorMessage: {importResult.ErrorMessage}");
            _output.WriteLine($"Import WorkflowHint: {importResult.WorkflowHint}");
            _output.WriteLine($"Import SuggestedNextActions: {string.Join(", ", importResult.SuggestedNextActions ?? new List<string>())}");

            Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

            // CRITICAL DEBUG: Check if QueryTable exists BEFORE SaveAsync
            _output.WriteLine("\nDEBUG: Checking QueryTable immediately after ImportAsync (before explicit save)...");
            var configBeforeSave = await _commands.GetLoadConfigAsync(batch, "TestQuery");
            _output.WriteLine($"LoadMode BEFORE save: {configBeforeSave.LoadMode}, TargetSheet: {configBeforeSave.TargetSheet}");

            await batch.SaveAsync();

            // DEBUG: Check if QueryTable exists AFTER SaveAsync
            _output.WriteLine("\nDEBUG: Checking QueryTable after explicit SaveAsync...");
            var configAfterSave = await _commands.GetLoadConfigAsync(batch, "TestQuery");
            _output.WriteLine($"LoadMode AFTER save: {configAfterSave.LoadMode}, TargetSheet: {configAfterSave.TargetSheet}");
        }

        // Step 2: Verify it loaded to table
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            _output.WriteLine("\nSTEP 2: Verify loaded to table...");
            var config = await _commands.GetLoadConfigAsync(batch, "TestQuery");
            _output.WriteLine($"LoadMode: {config.LoadMode}, TargetSheet: {config.TargetSheet}");
            Assert.Equal(PowerQueryLoadMode.LoadToTable, config.LoadMode);
        }

        // Step 3: Manually change Formula property WITHOUT calling UpdateAsync
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            _output.WriteLine("\nSTEP 3: Manually changing Formula property...");

            await File.WriteAllTextAsync(queryFile, mCode2);
            string newMCode = await File.ReadAllTextAsync(queryFile);

            await batch.ExecuteAsync<int>((ctx, ct) =>
            {
                dynamic? query = null;
                try
                {
                    dynamic queries = ctx.Book.Queries;
                    query = queries.Item("TestQuery");

                    _output.WriteLine($"Query found, current formula length: {((string)query.Formula).Length}");

                    // ⚠️ THIS IS THE CRITICAL LINE - Does this destroy the QueryTable?
                    query.Formula = newMCode;

                    _output.WriteLine("Formula property changed.");
                    return ValueTask.FromResult(0);
                }
                finally
                {
                    if (query != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(query);
                    }
                }
            });

            await batch.SaveAsync();
            _output.WriteLine("Changes saved.");
        }

        // Step 4: Check load config after Formula change
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            _output.WriteLine("\nSTEP 4: Checking load config after Formula change...");
            var configAfter = await _commands.GetLoadConfigAsync(batch, "TestQuery");

            _output.WriteLine($"LoadMode after Formula change: {configAfter.LoadMode}");
            _output.WriteLine($"TargetSheet after Formula change: {configAfter.TargetSheet}");

            if (configAfter.LoadMode == PowerQueryLoadMode.ConnectionOnly)
            {
                _output.WriteLine("⚠️ CRITICAL: Excel REMOVED the QueryTable when Formula was changed!");
                _output.WriteLine("⚠️ This explains why UpdateAsync restoration isn't working.");
            }
            else if (configAfter.LoadMode == PowerQueryLoadMode.LoadToTable)
            {
                _output.WriteLine("✅ Excel PRESERVED the QueryTable after Formula change.");
                _output.WriteLine("✅ This means the bug is in our UpdateAsync restoration logic.");
            }

            // Step 5: Try to restore manually
            _output.WriteLine("\nSTEP 5: Attempting manual restoration with SetLoadToTableAsync...");
            var restoreResult = await _commands.SetLoadToTableAsync(batch, "TestQuery", "TestQuery", privacyLevel: null);

            _output.WriteLine($"Restore Success: {restoreResult.Success}");
            _output.WriteLine($"Restore ErrorMessage: {restoreResult.ErrorMessage}");

            if (restoreResult.Success)
            {
                await batch.SaveAsync();
                _output.WriteLine("Restore saved successfully.");
            }
        }

        // Step 6: Verify final state
        await using (var batch = await ExcelSession.BeginBatchAsync(testFile))
        {
            _output.WriteLine("\nSTEP 6: Verifying final state after manual restoration...");
            var finalConfig = await _commands.GetLoadConfigAsync(batch, "TestQuery");

            _output.WriteLine($"Final LoadMode: {finalConfig.LoadMode}");
            _output.WriteLine($"Final TargetSheet: {finalConfig.TargetSheet}");

            if (finalConfig.LoadMode == PowerQueryLoadMode.LoadToTable)
            {
                _output.WriteLine("✅ Manual restoration with SetLoadToTableAsync WORKED.");
                _output.WriteLine("✅ This proves SetLoadToTableAsync can restore config after Formula change.");
                _output.WriteLine("⚠️ BUG is likely in UpdateAsync's save timing or batch handling.");
            }
            else
            {
                _output.WriteLine("❌ Manual restoration FAILED.");
                _output.WriteLine("❌ This suggests a deeper issue with SetLoadToTableAsync or save timing.");
            }
        }
    }
}
