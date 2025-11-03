using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query load configuration operations
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> SetConnectionOnlyAsync(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-connection-only"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Remove any existing connections and QueryTables for this query
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting connection only: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToTableResult> SetLoadToTableAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new PowerQueryLoadToTableResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-table",
            QueryName = queryName,
            SheetName = sheetName,
            WorkflowStatus = "Failed"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<PowerQueryLoadToTableResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? targetSheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            try
            {
                // STEP 1: Verify query exists
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 2: Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        if (sheet.Name == sheetName)
                        {
                            targetSheet = sheet;
                            sheet = null; // Don't release - we're keeping it
                            break;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // STEP 3: Configure query (remove old connections, create new QueryTable)
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryName,
                    RefreshImmediately = true // CRITICAL: Refresh synchronously to persist QueryTable properly
                };
                PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

                result.ConfigurationApplied = true;

                // Note: RefreshImmediately=true causes CreateQueryTable to call queryTable.Refresh(false)
                // which is SYNCHRONOUS and ensures proper persistence when workbook is saved.
                // This follows Microsoft's documented pattern: Create → Refresh(False) → Save
                // (See VBA example: https://learn.microsoft.com/en-us/office/troubleshoot/excel/...)
                // RefreshAll() is ASYNCHRONOUS and unreliable for individual QueryTable persistence.

                // STEP 4: VERIFY data was actually loaded
                queryTables = targetSheet.QueryTables;
                string normalizedName = queryName.Replace(" ", "_");
                bool foundQueryTable = false;
                int rowsLoaded = 0;

                for (int qt = 1; qt <= queryTables.Count; qt++)
                {
                    dynamic? qt_obj = null;
                    try
                    {
                        qt_obj = queryTables.Item(qt);
                        string qtName = qt_obj.Name?.ToString() ?? "";

                        if (qtName.Equals(normalizedName, StringComparison.OrdinalIgnoreCase) ||
                            qtName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                        {
                            foundQueryTable = true;

                            // Get row count from ResultRange
                            try
                            {
                                dynamic? resultRange = qt_obj.ResultRange;
                                if (resultRange != null)
                                {
                                    rowsLoaded = resultRange.Rows.Count;
                                    ComUtilities.Release(ref resultRange);
                                }
                            }
                            catch
                            {
                                // If we can't get row count, at least we found the QueryTable
                                rowsLoaded = 0;
                            }

                            queryTable = qt_obj;
                            qt_obj = null; // Keep reference
                            break;
                        }
                    }
                    finally
                    {
                        if (qt_obj != null)
                        {
                            ComUtilities.Release(ref qt_obj);
                        }
                    }
                }

                if (foundQueryTable)
                {
                    result.Success = true;
                    result.DataLoadedToTable = true;
                    result.RowsLoaded = rowsLoaded;
                    result.WorkflowStatus = "Complete";
                }
                else
                {
                    result.Success = false;
                    result.DataLoadedToTable = false;
                    result.RowsLoaded = 0;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = $"Configuration applied but QueryTable not found after refresh";
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting load to table: {ex.Message}";
                result.WorkflowStatus = "Failed";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToDataModelResult> SetLoadToDataModelAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadToDataModelResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-data-model",
            QueryName = queryName,
            ConfigurationApplied = false,
            DataLoadedToModel = false,
            RowsLoaded = 0,
            TablesInDataModel = 0,
            WorkflowStatus = "Failed"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.ExecuteAsync<PowerQueryLoadToDataModelResult>(async (ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // STEP 1: Configure query to load to data model
                // Remove existing table connections
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Configure Data Model loading using Connections.Add2
                bool configSuccess = SetQueryLoadToDataModel(ctx.Book, queryName, out string? configError);
                result.ConfigurationApplied = configSuccess;

                if (!configSuccess)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to configure query for Data Model loading: {configError ?? "Unknown error"}";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 2: Verify data was actually loaded to Data Model
                dynamic? model = null;
                dynamic? modelTables = null;
                try
                {
                    model = ctx.Book.Model;
                    modelTables = model.ModelTables;
                    result.TablesInDataModel = modelTables.Count;

                    // Find the query's table in the Data Model
                    bool foundTable = false;
                    int rowCount = 0;

                    for (int i = 1; i <= modelTables.Count; i++)
                    {
                        dynamic? table = null;
                        try
                        {
                            table = modelTables.Item(i);
                            string tableName = table.Name?.ToString() ?? "";

                            // Match by query name (Excel may add prefixes/suffixes)
                            if (tableName.Contains(queryName, StringComparison.OrdinalIgnoreCase))
                            {
                                foundTable = true;

                                // Get row count
                                try
                                {
                                    rowCount = (int)table.RecordCount;
                                }
                                catch
                                {
                                    rowCount = 0; // RecordCount may not be available immediately
                                }

                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref table);
                        }
                    }

                    result.DataLoadedToModel = foundTable;
                    result.RowsLoaded = rowCount;

                    if (foundTable)
                    {
                        result.Success = true;
                        result.WorkflowStatus = "Complete";
                    }
                    else
                    {
                        result.Success = false;
                        result.ErrorMessage = "Query configured and refreshed, but table not found in Data Model";
                        result.WorkflowStatus = "Partial";
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                    ComUtilities.Release(ref model);
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("protected", StringComparison.OrdinalIgnoreCase) || 
                                             comEx.Message.Contains("sensitivity label", StringComparison.OrdinalIgnoreCase))
            {
                // Microsoft Purview sensitivity label error - encrypted file
                // Get M code to extract file path from File.Contents()
                string? filePath = null;
                try
                {
                    var viewResult = await ViewAsync(batch, queryName);
                    if (viewResult.Success && !string.IsNullOrEmpty(viewResult.MCode))
                    {
                        filePath = ExtractFileContentsPath(viewResult.MCode);
                    }
                }
                catch (COMException)
                {
                    // If we can't get M code due to COM error, continue without file path
                }

                string filePathInfo = !string.IsNullOrEmpty(filePath) 
                    ? $"\n\nSource file: {filePath}" 
                    : "";

                result.Success = false;
                result.ErrorMessage = $"Source Excel file has Microsoft Purview sensitivity labels (encryption).{filePathInfo}\n\n" +
                                    "Power Query cannot access encrypted Excel files.\n\n" +
                                    "SOLUTION 1 (Recommended): Change sensitivity label to Public\n" +
                                    "  - Open the source file in Excel\n" +
                                    "  - Click Home tab → Sensitivity button → Select \"Public\" label\n" +
                                    "  - Save and close\n" +
                                    "  - Retry: excel_powerquery(action: 'set-load-to-data-model', queryName: '{queryName}')\n\n" +
                                    "SOLUTION 2: Modify M code to use different data source\n" +
                                    "  - Replace File.Contents() with Excel.CurrentWorkbook() if data is in same workbook\n" +
                                    "  - Export source data to CSV and use Csv.Document()\n" +
                                    "  - Use ODBC or SQL connection if source is a database\n\n" +
                                    "Technical details: https://learn.microsoft.com/en-us/power-query/connectors/excel#known-issues-and-limitations";
                
                result.WorkflowStatus = "Failed";
                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error in atomic load-to-data-model operation: {ex.Message}";
                result.WorkflowStatus = "Failed";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadToBothResult> SetLoadToBothAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new PowerQueryLoadToBothResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-set-load-to-both",
            QueryName = queryName,
            SheetName = sheetName,
            WorkflowStatus = "Failed"
        };

        return await batch.Execute<PowerQueryLoadToBothResult>((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? sheets = null;
            dynamic? targetSheet = null;
            try
            {
                // STEP 1: Verify query exists
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    result.WorkflowStatus = "Failed";
                    return result;
                }

                // STEP 2: Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        if (sheet.Name == sheetName)
                        {
                            targetSheet = sheet;
                            sheet = null; // Don't release - we're keeping it
                            break;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // STEP 4: Configure query for BOTH table and Data Model loading
                ConnectionHelpers.RemoveConnections(ctx.Book, queryName);
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, queryName);

                // Create QueryTable for worksheet loading
                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryName,
                    RefreshImmediately = true // CRITICAL: Refresh synchronously to persist QueryTable properly
                };
                PowerQueryHelpers.CreateQueryTable(targetSheet, queryName, queryTableOptions);

                // Configure query for Data Model loading
                if (!SetQueryLoadToDataModel(ctx.Book, queryName, out string? dmConfigError))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to configure query for Data Model loading: {dmConfigError ?? "Unknown error"}";
                    result.WorkflowStatus = "Partial";
                    return result;
                }

                result.ConfigurationApplied = true;

                // STEP 6: VERIFY data loaded to BOTH destinations
                bool foundInTable = false;
                bool foundInDataModel = false;
                int tableRows = 0;
                int modelRows = 0;
                int tablesInDataModel = 0;

                // Verify table loading
                dynamic? queryTables = null;
                try
                {
                    queryTables = targetSheet.QueryTables;
                    string normalizedName = queryName.Replace(" ", "_");

                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        dynamic? qt_obj = null;
                        try
                        {
                            qt_obj = queryTables.Item(qt);
                            string qtName = qt_obj.Name?.ToString() ?? "";

                            if (qtName.Equals(normalizedName, StringComparison.OrdinalIgnoreCase) ||
                                qtName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                            {
                                foundInTable = true;

                                // Get row count from ResultRange
                                try
                                {
                                    dynamic? resultRange = qt_obj.ResultRange;
                                    if (resultRange != null)
                                    {
                                        tableRows = resultRange.Rows.Count;
                                        ComUtilities.Release(ref resultRange);
                                    }
                                }
                                catch
                                {
                                    tableRows = 0;
                                }
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref qt_obj);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queryTables);
                }

                // Verify Data Model loading
                dynamic? model = null;
                dynamic? modelTables = null;
                try
                {
                    model = ctx.Book.Model;
                    if (model != null)
                    {
                        modelTables = model.ModelTables;
                        tablesInDataModel = modelTables.Count;

                        for (int t = 1; t <= modelTables.Count; t++)
                        {
                            dynamic? table = null;
                            try
                            {
                                table = modelTables.Item(t);
                                string tableName = table.Name?.ToString() ?? "";

                                if (tableName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                                {
                                    foundInDataModel = true;
                                    modelRows = table.RecordCount;
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                    ComUtilities.Release(ref model);
                }

                // Set result based on verification
                result.DataLoadedToTable = foundInTable;
                result.DataLoadedToModel = foundInDataModel;
                result.RowsLoadedToTable = tableRows;
                result.RowsLoadedToModel = modelRows;
                result.TablesInDataModel = tablesInDataModel;

                if (foundInTable && foundInDataModel)
                {
                    result.Success = true;
                    result.WorkflowStatus = "Complete";
                }
                else if (foundInTable && !foundInDataModel)
                {
                    result.Success = false;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = "Data loaded to table but not to Data Model";
                }
                else if (!foundInTable && foundInDataModel)
                {
                    result.Success = false;
                    result.WorkflowStatus = "Partial";
                    result.ErrorMessage = "Data loaded to Data Model but not to table";
                }
                else
                {
                    result.Success = false;
                    result.WorkflowStatus = "Failed";
                    result.ErrorMessage = "Data not loaded to either destination";
                }

                return result;
            }
            catch (COMException comEx) when (comEx.Message.Contains("Information is needed in order to combine data") ||
                                             comEx.Message.Contains("Formula.Firewall", StringComparison.OrdinalIgnoreCase))
            {
                // Privacy level error - must be configured manually in Excel UI
                result.Success = false;
                result.ErrorMessage = "Privacy level error: This query combines data from multiple sources. " +
                                    "Open the file in Excel and configure privacy levels manually: " +
                                    "File → Options → Privacy. See COMMANDS.md for details.";
                result.WorkflowStatus = "Failed";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error in atomic load-to-both operation: {ex.Message}";
                result.WorkflowStatus = "Failed";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return await batch.Execute<PowerQueryLoadConfigResult>((ctx, ct) =>
        {
            try
            {
                dynamic query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Check for QueryTables first (table loading)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                dynamic worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic worksheet = worksheets.Item(ws);
                    dynamic queryTables = worksheet.QueryTables;

                    for (int qt = 1; qt <= queryTables.Count; qt++)
                    {
                        try
                        {
                            dynamic queryTable = queryTables.Item(qt);
                            string qtName = queryTable.Name?.ToString() ?? "";

                            // Check if this QueryTable is for our query
                            if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                                qtName.Contains(queryName.Replace(" ", "_")))
                            {
                                hasTableConnection = true;
                                targetSheet = worksheet.Name;
                                break;
                            }
                        }
                        catch
                        {
                            // Skip invalid QueryTables
                            continue;
                        }
                    }
                    if (hasTableConnection) break;
                }

                // Check for connections (for data model or other types)
                dynamic connections = ctx.Book.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";

                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        result.HasConnection = true;

                        // If we don't have a table connection but have a workbook connection,
                        // it's likely a data model connection
                        if (!hasTableConnection)
                        {
                            hasDataModelConnection = true;
                        }
                    }
                    else if (connName.Equals($"DataModel_{queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        // This is our explicit data model connection marker
                        result.HasConnection = true;
                        hasDataModelConnection = true;
                    }
                }

                // Always check for named range markers that indicate data model loading
                // (even if we have table connections, for LoadToBoth mode)
                if (!hasDataModelConnection)
                {
                    // Check for our data model marker
                    try
                    {
                        dynamic names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        for (int i = 1; i <= names.Count; i++)
                        {
                            try
                            {
                                dynamic existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    hasDataModelConnection = true;
                                    break;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // Cannot check names
                    }

                    // Fallback: Check if the query has data model indicators
                    if (!hasDataModelConnection)
                    {
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query, ctx.Book);
                    }
                }

                // Determine load mode
                if (hasTableConnection && hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToBoth;
                }
                else if (hasTableConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToTable;
                }
                else if (hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToDataModel;
                }
                else
                {
                    result.LoadMode = PowerQueryLoadMode.ConnectionOnly;
                }

                result.TargetSheet = targetSheet;
                result.IsLoadedToDataModel = hasDataModelConnection;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting load config: {ex.Message}";
                return result;
            }
        });
    }
}
