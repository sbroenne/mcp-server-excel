using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Analysis;

public sealed partial class AnalysisCommands
{
    /// <inheritdoc />
    public GoalSeekResult GoalSeek(
        IExcelBatch batch,
        string sheetName,
        string formulaCell,
        double? goal,
        string changingCell)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(formulaCell);
        if (goal is null)
        {
            throw new ArgumentNullException(nameof(goal));
        }
        ArgumentException.ThrowIfNullOrWhiteSpace(changingCell);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? formulaRange = null;
            Excel.Range? changingRange = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
                formulaRange = sheet.Range[formulaCell];
                changingRange = sheet.Range[changingCell];

                var converged = formulaRange.GoalSeek(goal.Value, changingRange);
                return new GoalSeekResult
                {
                    Success = true,
                    Converged = converged,
                    FormulaValue = Convert.ToDouble(formulaRange.Value2),
                    ChangingValue = Convert.ToDouble(changingRange.Value2),
                    Message = converged
                        ? $"Goal Seek reached {goal.Value} in '{formulaCell}'."
                        : $"Goal Seek completed without converging on {goal.Value} in '{formulaCell}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref changingRange);
                ComUtilities.Release(ref formulaRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
