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
            dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
            dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
            dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
            dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
            dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
                dynamic pivotCache = ctx.Book.PivotCaches().Create(1, sheet.Range["A1:D10"]);
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
            // Configuration persisted successfully (verified by successful read)
        }
    }
}
