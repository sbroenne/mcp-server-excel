using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for SetGrandTotals operation - show/hide row and column grand totals.
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    public void SetGrandTotals_ShowBoth_EnablesBothGrandTotals()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_ShowBoth_EnablesBothGrandTotals));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create PivotTable
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item[1];
            dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
            dynamic newSheet = ctx.Book.Worksheets.Add();
            newSheet.Name = "PivotSheet";
            dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

            // Add fields
            dynamic? rowField = null;
            dynamic? colField = null;
            dynamic? dataField = null;
            try
            {
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1; // xlRowField

                colField = pivot.PivotFields("Region");
                colField.Orientation = 2; // xlColumnField

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4; // xlDataField

                return new { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref colField);
                ComUtilities.Release(ref dataField);
            }
        });

        // Act
        var result = _pivotCommands.SetGrandTotals(batch, "TestPivot", true, true);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");

        // Verify actual Excel state
        var grandTotalState = batch.Execute((ctx, ct) =>
        {
            dynamic? verifySheet = null;
            dynamic? verifyPivotTables = null;
            dynamic? verifyPivot = null;
            try
            {
                verifySheet = ctx.Book.Worksheets["PivotSheet"];
                verifyPivotTables = verifySheet.PivotTables();
                verifyPivot = verifyPivotTables("TestPivot");
                return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
            }
            finally
            {
                ComUtilities.Release(ref verifyPivot);
                ComUtilities.Release(ref verifyPivotTables);
                ComUtilities.Release(ref verifySheet);
            }
        });
        Assert.True(grandTotalState.RowGrand, "Expected RowGrand to be true");
        Assert.True(grandTotalState.ColumnGrand, "Expected ColumnGrand to be true");
    }

    [Fact]
    public void SetGrandTotals_HideBoth_DisablesBothGrandTotals()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_HideBoth_DisablesBothGrandTotals));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create PivotTable (grand totals on by default)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item[1];
            dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
            dynamic newSheet = ctx.Book.Worksheets.Add();
            newSheet.Name = "PivotSheet";
            dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

            dynamic? rowField = null;
            dynamic? colField = null;
            dynamic? dataField = null;
            try
            {
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1;

                colField = pivot.PivotFields("Region");
                colField.Orientation = 2;

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4;

                return new { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref colField);
                ComUtilities.Release(ref dataField);
            }
        });

        // Act - Hide both grand totals
        var result = _pivotCommands.SetGrandTotals(batch, "TestPivot", false, false);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");

        // Verify actual Excel state
        var grandTotalState = batch.Execute((ctx, ct) =>
        {
            dynamic? verifySheet = null;
            dynamic? verifyPivotTables = null;
            dynamic? verifyPivot = null;
            try
            {
                verifySheet = ctx.Book.Worksheets["PivotSheet"];
                verifyPivotTables = verifySheet.PivotTables();
                verifyPivot = verifyPivotTables("TestPivot");
                return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
            }
            finally
            {
                ComUtilities.Release(ref verifyPivot);
                ComUtilities.Release(ref verifyPivotTables);
                ComUtilities.Release(ref verifySheet);
            }
        });
        Assert.False(grandTotalState.RowGrand, "Expected RowGrand to be false");
        Assert.False(grandTotalState.ColumnGrand, "Expected ColumnGrand to be false");
    }

    [Fact]
    public void SetGrandTotals_ShowRowHideColumn_EnablesRowDisablesColumnGrandTotals()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_ShowRowHideColumn_EnablesRowDisablesColumnGrandTotals));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create PivotTable
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item[1];
            dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
            dynamic newSheet = ctx.Book.Worksheets.Add();
            newSheet.Name = "PivotSheet";
            dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

            dynamic? rowField = null;
            dynamic? colField = null;
            dynamic? dataField = null;
            try
            {
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1;

                colField = pivot.PivotFields("Region");
                colField.Orientation = 2;

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4;

                return new { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref colField);
                ComUtilities.Release(ref dataField);
            }
        });

        // Act - Show row, hide column
        var result = _pivotCommands.SetGrandTotals(batch, "TestPivot", true, false);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");

        // Verify actual Excel state
        var grandTotalState = batch.Execute((ctx, ct) =>
        {
            dynamic? verifySheet = null;
            dynamic? verifyPivotTables = null;
            dynamic? verifyPivot = null;
            try
            {
                verifySheet = ctx.Book.Worksheets["PivotSheet"];
                verifyPivotTables = verifySheet.PivotTables();
                verifyPivot = verifyPivotTables("TestPivot");
                return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
            }
            finally
            {
                ComUtilities.Release(ref verifyPivot);
                ComUtilities.Release(ref verifyPivotTables);
                ComUtilities.Release(ref verifySheet);
            }
        });
        Assert.True(grandTotalState.RowGrand, "Expected RowGrand to be true");
        Assert.False(grandTotalState.ColumnGrand, "Expected ColumnGrand to be false");
    }

    [Fact]
    public void SetGrandTotals_HideRowShowColumn_DisablesRowEnablesColumnGrandTotals()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_HideRowShowColumn_DisablesRowEnablesColumnGrandTotals));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create PivotTable
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item[1];
            dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
            dynamic newSheet = ctx.Book.Worksheets.Add();
            newSheet.Name = "PivotSheet";
            dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

            dynamic? rowField = null;
            dynamic? colField = null;
            dynamic? dataField = null;
            try
            {
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1;

                colField = pivot.PivotFields("Region");
                colField.Orientation = 2;

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4;

                return new { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref colField);
                ComUtilities.Release(ref dataField);
            }
        });

        // Act - Hide row, show column
        var result = _pivotCommands.SetGrandTotals(batch, "TestPivot", false, true);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");

        // Verify actual Excel state
        var grandTotalState = batch.Execute((ctx, ct) =>
        {
            dynamic? verifySheet = null;
            dynamic? verifyPivotTables = null;
            dynamic? verifyPivot = null;
            try
            {
                verifySheet = ctx.Book.Worksheets["PivotSheet"];
                verifyPivotTables = verifySheet.PivotTables();
                verifyPivot = verifyPivotTables("TestPivot");
                return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
            }
            finally
            {
                ComUtilities.Release(ref verifyPivot);
                ComUtilities.Release(ref verifyPivotTables);
                ComUtilities.Release(ref verifySheet);
            }
        });
        Assert.False(grandTotalState.RowGrand, "Expected RowGrand to be false");
        Assert.True(grandTotalState.ColumnGrand, "Expected ColumnGrand to be true");
    }

    [Fact]
    public void SetGrandTotals_MultipleSequentialChanges_AppliesEachConfiguration()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_MultipleSequentialChanges_AppliesEachConfiguration));

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create PivotTable
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item[1];
            dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
            dynamic newSheet = ctx.Book.Worksheets.Add();
            newSheet.Name = "PivotSheet";
            dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

            dynamic? rowField = null;
            dynamic? colField = null;
            dynamic? dataField = null;
            try
            {
                rowField = pivot.PivotFields("Product");
                rowField.Orientation = 1;

                colField = pivot.PivotFields("Region");
                colField.Orientation = 2;

                dataField = pivot.PivotFields("Sales");
                dataField.Orientation = 4;

                return new { Success = true };
            }
            finally
            {
                ComUtilities.Release(ref rowField);
                ComUtilities.Release(ref colField);
                ComUtilities.Release(ref dataField);
            }
        });

        // Act & Assert - Multiple sequential changes
        var result1 = _pivotCommands.SetGrandTotals(batch, "TestPivot", true, true);
        Assert.True(result1.Success, $"Change 1 failed: {result1.ErrorMessage}");

        var result2 = _pivotCommands.SetGrandTotals(batch, "TestPivot", false, false);
        Assert.True(result2.Success, $"Change 2 failed: {result2.ErrorMessage}");

        var result3 = _pivotCommands.SetGrandTotals(batch, "TestPivot", true, false);
        Assert.True(result3.Success, $"Change 3 failed: {result3.ErrorMessage}");

        var result4 = _pivotCommands.SetGrandTotals(batch, "TestPivot", false, true);
        Assert.True(result4.Success, $"Change 4 failed: {result4.ErrorMessage}");

        // Verify final Excel state matches last configuration (false, true)
        var grandTotalState = batch.Execute((ctx, ct) =>
        {
            dynamic? verifySheet = null;
            dynamic? verifyPivotTables = null;
            dynamic? verifyPivot = null;
            try
            {
                verifySheet = ctx.Book.Worksheets["PivotSheet"];
                verifyPivotTables = verifySheet.PivotTables();
                verifyPivot = verifyPivotTables("TestPivot");
                return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
            }
            finally
            {
                ComUtilities.Release(ref verifyPivot);
                ComUtilities.Release(ref verifyPivotTables);
                ComUtilities.Release(ref verifySheet);
            }
        });
        Assert.False(grandTotalState.RowGrand, "Expected RowGrand to be false after final change");
        Assert.True(grandTotalState.ColumnGrand, "Expected ColumnGrand to be true after final change");
    }

    [Fact]
    public void SetGrandTotals_RoundTrip_PersistsConfiguration()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetGrandTotals_RoundTrip_PersistsConfiguration));

        // Session 1: Create and configure
        using (var batch1 = ExcelSession.BeginBatch(testFile))
        {
            batch1.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item[1];
                dynamic pivotCache = ctx.Book.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheet.Range["A1:D10"]);
                dynamic newSheet = ctx.Book.Worksheets.Add();
                newSheet.Name = "PivotSheet";
                dynamic pivot = pivotCache.CreatePivotTable(newSheet.Range["A1"], "TestPivot");

                dynamic? rowField = null;
                dynamic? colField = null;
                dynamic? dataField = null;
                try
                {
                    rowField = pivot.PivotFields("Product");
                    rowField.Orientation = 1;

                    colField = pivot.PivotFields("Region");
                    colField.Orientation = 2;

                    dataField = pivot.PivotFields("Sales");
                    dataField.Orientation = 4;

                    return new { Success = true };
                }
                finally
                {
                    ComUtilities.Release(ref rowField);
                    ComUtilities.Release(ref colField);
                    ComUtilities.Release(ref dataField);
                }
            });

            var setResult = _pivotCommands.SetGrandTotals(batch1, "TestPivot", true, false);
            Assert.True(setResult.Success, $"SetGrandTotals failed: {setResult.ErrorMessage}");

            batch1.Save();
        }

        // Session 2: Verify persistence
        using (var batch2 = ExcelSession.BeginBatch(testFile))
        {
            var readResult = _pivotCommands.Read(batch2, "TestPivot");
            Assert.True(readResult.Success, $"Read failed: {readResult.ErrorMessage}");

            // Verify grand total state was persisted (showRowGrandTotals=true, showColumnGrandTotals=false)
            var grandTotalState = batch2.Execute((ctx, ct) =>
            {
                dynamic? verifySheet = null;
                dynamic? verifyPivotTables = null;
                dynamic? verifyPivot = null;
                try
                {
                    verifySheet = ctx.Book.Worksheets["PivotSheet"];
                    verifyPivotTables = verifySheet.PivotTables();
                    verifyPivot = verifyPivotTables("TestPivot");
                    return new { RowGrand = (bool)verifyPivot.RowGrand, ColumnGrand = (bool)verifyPivot.ColumnGrand };
                }
                finally
                {
                    ComUtilities.Release(ref verifyPivot);
                    ComUtilities.Release(ref verifyPivotTables);
                    ComUtilities.Release(ref verifySheet);
                }
            });
            Assert.True(grandTotalState.RowGrand, "Expected RowGrand=true to be persisted");
            Assert.False(grandTotalState.ColumnGrand, "Expected ColumnGrand=false to be persisted");
        }
    }
}




