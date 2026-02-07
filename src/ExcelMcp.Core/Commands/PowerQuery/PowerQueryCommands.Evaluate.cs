using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query evaluate operation - executes M code and returns results without creating a permanent query
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryEvaluateResult Evaluate(IExcelBatch batch, string mCode)
    {
        var result = new PowerQueryEvaluateResult
        {
            FilePath = batch.WorkbookPath,
            MCode = mCode
        };

        // Validate M code
        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code is required for evaluate action", nameof(mCode));
        }

        // Generate unique temporary names
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var tempQueryName = $"__pq_eval_{uniqueId}";
        var tempSheetName = $"__pq_eval_{uniqueId}";

        return batch.Execute((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            dynamic? query = null;
            dynamic? worksheets = null;
            dynamic? tempSheet = null;
            dynamic? listObjects = null;
            dynamic? listObject = null;
            dynamic? queryTable = null;
            dynamic? range = null;
            dynamic? usedRange = null;

            try
            {
                // STEP 1: Create temporary query with the M code
                queriesCollection = ctx.Book.Queries;
                query = queriesCollection.Add(tempQueryName, mCode);

                // STEP 2: Create temporary worksheet and load data to it
                worksheets = ctx.Book.Worksheets;
                tempSheet = worksheets.Add();
                tempSheet.Name = tempSheetName;

                // STEP 3: Load query to worksheet using QueryTable
                // Connection string for Power Query - must include Extended Properties
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={tempQueryName};Extended Properties=\"\"";
                range = tempSheet.Range["A1"];

                listObjects = tempSheet.ListObjects;
                listObject = listObjects.Add(
                    0,                  // SourceType: 0 = xlSrcExternal
                    connectionString,   // Source: connection string
                    Type.Missing,       // LinkSource
                    1,                  // XlListObjectHasHeaders: xlYes = 1
                    range               // Destination: starting cell
                );

                // Get the QueryTable to refresh (this executes the M code)
                queryTable = listObject.QueryTable;

                // Configure QueryTable to select from the query
                queryTable.CommandType = 2; // xlCmdSql
                queryTable.CommandText = $"SELECT * FROM [{tempQueryName}]";
                queryTable.BackgroundQuery = false; // Synchronous

                // STEP 4: Refresh to execute the M code (errors will throw via QueryTable.Refresh)
                // This is the key step - if M code has errors, this will throw!
                queryTable.Refresh(false); // false = synchronous

                // STEP 5: Read the results from the worksheet
                // Get the data range from the ListObject
                dynamic? dataBodyRange = null;
                dynamic? headerRowRange = null;
                try
                {
                    // Get column names from header row
                    headerRowRange = listObject.HeaderRowRange;
                    if (headerRowRange != null)
                    {
                        dynamic? headerValues = headerRowRange.Value2;
                        if (headerValues != null)
                        {
                            if (headerValues is object[,] headers2D)
                            {
                                for (int col = 1; col <= headers2D.GetLength(1); col++)
                                {
                                    result.Columns.Add(headers2D[1, col]?.ToString() ?? $"Column{col}");
                                }
                            }
                            else if (headerValues is object[] headers1D)
                            {
                                for (int i = 0; i < headers1D.Length; i++)
                                {
                                    result.Columns.Add(headers1D[i]?.ToString() ?? $"Column{i + 1}");
                                }
                            }
                            else
                            {
                                // Single cell
                                result.Columns.Add(headerValues?.ToString() ?? "Column1");
                            }
                        }
                    }
                    result.ColumnCount = result.Columns.Count;

                    // Get data rows
                    dataBodyRange = listObject.DataBodyRange;
                    if (dataBodyRange != null)
                    {
                        dynamic? dataValues = dataBodyRange.Value2;
                        if (dataValues != null)
                        {
                            if (dataValues is object[,] data2D)
                            {
                                int rowCount = data2D.GetLength(0);
                                int colCount = data2D.GetLength(1);

                                for (int row = 1; row <= rowCount; row++)
                                {
                                    var rowData = new List<object?>();
                                    for (int col = 1; col <= colCount; col++)
                                    {
                                        rowData.Add(ConvertCellValue(data2D[row, col]));
                                    }
                                    result.Rows.Add(rowData);
                                }
                            }
                            else if (dataValues is object[] data1D)
                            {
                                // Single row
                                var rowData = new List<object?>();
                                foreach (var val in data1D)
                                {
                                    rowData.Add(ConvertCellValue(val));
                                }
                                result.Rows.Add(rowData);
                            }
                            else
                            {
                                // Single cell
                                result.Rows.Add([ConvertCellValue(dataValues)]);
                            }
                        }
                    }
                    result.RowCount = result.Rows.Count;
                }
                finally
                {
                    ComUtilities.Release(ref headerRowRange);
                    ComUtilities.Release(ref dataBodyRange);
                }

                result.Success = true;
            }
            finally
            {
                // STEP 6: Cleanup - delete temporary objects
                // Order matters: delete table, sheet, query, connections

                // Delete the ListObject (table)
                try
                {
                    if (listObject != null)
                    {
                        listObject.Delete();
                    }
                }
                catch (COMException) { /* ignore cleanup errors */ }

                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref listObject);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref usedRange);

                // Delete the temporary worksheet
                try
                {
                    if (tempSheet != null)
                    {
                        // Suppress alerts to avoid "Are you sure you want to delete?" prompt
                        dynamic? app = null;
                        try
                        {
                            app = ctx.Book.Application;
                            bool originalAlerts = app.DisplayAlerts;
                            app.DisplayAlerts = false;
                            try
                            {
                                tempSheet.Delete();
                            }
                            finally
                            {
                                app.DisplayAlerts = originalAlerts;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref app);
                        }
                    }
                }
                catch (COMException) { /* ignore cleanup errors */ }

                ComUtilities.Release(ref tempSheet);
                ComUtilities.Release(ref worksheets);

                // Delete the temporary query
                try
                {
                    if (query != null)
                    {
                        query.Delete();
                    }
                }
                catch (COMException) { /* ignore cleanup errors */ }

                ComUtilities.Release(ref query);
                ComUtilities.Release(ref queriesCollection);

                // Clean up any lingering connections
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int i = connections.Count; i >= 1; i--)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Contains(tempQueryName))
                            {
                                conn.Delete();
                            }
                        }
                        catch (COMException) { /* ignore cleanup errors */ }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch (COMException) { /* ignore cleanup errors */ }
                finally
                {
                    ComUtilities.Release(ref connections);
                }
            }

            return result;
        });
    }

    /// <summary>
    /// Converts Excel cell values to JSON-friendly types
    /// </summary>
    private static object? ConvertCellValue(object? value)
    {
        if (value == null || value == DBNull.Value)
            return null;

        return value switch
        {
            DateTime dt => dt.ToString("O"), // ISO 8601 format
            double d => d,
            int i => i,
            long l => l,
            bool b => b,
            string s => s,
            _ => value.ToString()
        };
    }
}


