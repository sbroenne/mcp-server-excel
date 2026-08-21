using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query View operations - STANDALONE implementation.
/// Based on Microsoft WorkbookQuery object model documentation.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// View Power Query details: M code, description, and load configuration.
    /// STANDALONE implementation following Microsoft WorkbookQuery API.
    /// </summary>
    /// <remarks>
    /// Microsoft Docs Reference:
    /// - WorkbookQuery.Name property (Read/Write String)
    /// - WorkbookQuery.Description property (Read/Write String)
    /// - WorkbookQuery.Formula property (Read/Write String) - The Power Query M code
    ///
    /// Load configuration detection follows the pattern established in Update:
    /// - QueryTable (from LoadTo/Create) - created via sheet.QueryTables.Add()
    /// - ListObject (from previous Update) - created via sheet.ListObjects.Add()
    /// Both are checked to determine if query is connection-only or loaded to worksheet.
    /// </remarks>
    public PowerQueryViewResult View(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;

            try
            {
                // STEP 1: Find the Power Query
                queries = ctx.Book.Queries;
                query = null;
                for (int i = 1; i <= queries.Count; i++)
                {
                    Excel.WorkbookQuery? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            q = null; // Don't release - keeping reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref q!);
                    }
                }

                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 2: Read WorkbookQuery properties (per Microsoft docs)
                string mCode = query.Formula?.ToString() ?? "";
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                var loadState = DetectLoadState(ctx.Book, queryName, ct);
                result.LoadMode = loadState.LoadMode;
                result.TargetSheet = loadState.TargetSheet;
                result.HasConnection = loadState.HasConnection;
                result.IsLoadedToDataModel = loadState.IsLoadedToDataModel;
                result.IsConnectionOnly = loadState.IsConnectionOnly;

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }
}
