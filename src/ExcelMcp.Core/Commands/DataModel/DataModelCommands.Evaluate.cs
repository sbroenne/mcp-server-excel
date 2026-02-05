using System.Runtime.InteropServices;

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model DAX EVALUATE query execution
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public DaxEvaluateResult Evaluate(IExcelBatch batch, string daxQuery)
    {
        // Validate input
        if (string.IsNullOrWhiteSpace(daxQuery))
        {
            throw new ArgumentException("daxQuery is required for evaluate action", nameof(daxQuery));
        }

        var result = new DaxEvaluateResult
        {
            FilePath = batch.WorkbookPath,
            DaxQuery = daxQuery
        };

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? dataModelConn = null;
            dynamic? modelConn = null;
            dynamic? adoConnection = null;
            dynamic? recordset = null;
            dynamic? fields = null;

            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Get the DataModelConnection (Type 7 connection to embedded Analysis Services)
                dataModelConn = model.DataModelConnection;
                if (dataModelConn == null)
                {
                    throw new InvalidOperationException("No DataModelConnection available - workbook may not have a Data Model");
                }

                modelConn = dataModelConn.ModelConnection;
                if (modelConn == null)
                {
                    throw new InvalidOperationException("No ModelConnection available from DataModelConnection");
                }

                // Get the ADO connection - this is a live MSOLAP connection to the embedded AS engine
                adoConnection = modelConn.ADOConnection;
                if (adoConnection == null)
                {
                    throw new InvalidOperationException("No ADOConnection available - cannot execute DAX query");
                }

                // Execute the DAX EVALUATE query directly via ADO
                // Wrap in try-catch to provide helpful error message when MSOLAP is missing
                try
                {
                    recordset = adoConnection.Execute(daxQuery);
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80040154))
                {
                    // REGDB_E_CLASSNOTREG (0x80040154) = "Class not registered"
                    // This occurs when MSOLAP provider is not installed
                    throw new InvalidOperationException(DataModelErrorMessages.MsolapProviderNotInstalled(), ex);
                }

                // Get field (column) information
                fields = recordset.Fields;
                int fieldCount = fields.Count;
                result.ColumnCount = fieldCount;

                // Extract column names (fully qualified: Table[Column])
                for (int i = 0; i < fieldCount; i++)
                {
                    dynamic? field = null;
                    try
                    {
                        field = fields.Item(i);
                        string fieldName = field.Name?.ToString() ?? $"Column{i}";
                        result.Columns.Add(fieldName);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }

                // Read all rows from the recordset
                while (!recordset.EOF)
                {
                    var row = new List<object?>();

                    for (int i = 0; i < fieldCount; i++)
                    {
                        dynamic? field = null;
                        try
                        {
                            field = fields.Item(i);
                            object? value = field.Value;

                            // Convert DBNull to null
                            if (value == DBNull.Value || value == null)
                            {
                                row.Add(null);
                            }
                            else
                            {
                                // Convert to JSON-friendly types
                                row.Add(ConvertToJsonFriendly(value));
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref field);
                        }
                    }

                    result.Rows.Add(row);
                    recordset.MoveNext();
                }

                result.RowCount = result.Rows.Count;
                result.Success = true;
            }
            finally
            {
                // Close recordset if open
                if (recordset != null)
                {
                    try
                    {
                        // State 1 = adStateOpen
                        if ((int)recordset.State == 1)
                        {
                            recordset.Close();
                        }
                    }
                    catch
                    {
                        // Ignore errors closing recordset
                    }
                }

                ComUtilities.Release(ref fields);
                ComUtilities.Release(ref recordset);
                ComUtilities.Release(ref adoConnection);
                ComUtilities.Release(ref modelConn);
                ComUtilities.Release(ref dataModelConn);
                ComUtilities.Release(ref model);
            }

            return result;
        }, timeoutCts.Token);
    }

    /// <summary>
    /// Converts a value from ADO recordset to a JSON-friendly type
    /// </summary>
    private static object ConvertToJsonFriendly(object value)
    {
        return value switch
        {
            // Dates
            DateTime dt => dt.ToString("O"), // ISO 8601 format

            // Numbers - preserve precision
            decimal d => d,
            double dbl => dbl,
            float f => f,
            long l => l,
            int i => i,
            short s => s,
            byte b => b,

            // Booleans
            bool bl => bl,

            // Strings and others - convert to string
            _ => value.ToString() ?? ""
        };
    }
}


