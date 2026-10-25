using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

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
            Excel.Model? model = null;
            Excel.WorkbookConnection? dataModelConn = null;
            Excel.ModelConnection? modelConn = null;
            ADODB._Connection? adoConnection = null;
            ADODB._Recordset? recordset = null;
            ADODB.Fields? fields = null;
            Exception? operationFailure = null;

            try
            {
                ct.ThrowIfCancellationRequested();

                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new OperationFailureException(
                        OperationFailureCategory.Prerequisite, DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Get the DataModelConnection (Type 7 connection to embedded Analysis Services)
                dataModelConn = model!.DataModelConnection;
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
                adoConnection = ComUtilities.ReadAdoConnection(modelConn);
                if (adoConnection == null)
                {
                    throw new InvalidOperationException("No ADOConnection available - cannot execute DAX query");
                }

                var adoDiagnostics = CollectAdoDiagnostics(adoConnection);

                // Execute the DAX EVALUATE query directly via ADO
                // Wrap in try-catch to provide helpful error message when MSOLAP is missing
                try
                {
                    recordset = adoConnection.Execute(daxQuery, out _, (int)ADODB.ExecuteOptionEnum.adOptionUnspecified);
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80040154))
                {
                    // REGDB_E_CLASSNOTREG (0x80040154) = "Class not registered"
                    // This occurs when the specific MSOLAP provider in the ADO connection is not registered.
                    throw new OperationFailureException(
                        OperationFailureCategory.DependencyUnavailable,
                        DataModelErrorMessages.MsolapClassNotRegistered(adoDiagnostics), ex);
                }
                catch (COMException ex)
                {
                    string failureMessage = GetDaxEvaluationFailureMessage(adoConnection, ex);
                    throw new InvalidOperationException(
                        failureMessage, ex);
                }

                ct.ThrowIfCancellationRequested();

                // Get field (column) information
                fields = recordset.Fields;
                int fieldCount = fields.Count;
                result.ColumnCount = fieldCount;

                // Extract column names (fully qualified: Table[Column])
                for (int i = 0; i < fieldCount; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    ADODB.Field? field = null;
                    try
                    {
                        field = fields[i];
                        string fieldName = field.Name?.ToString() ?? $"Column{i}";
                        result.Columns.Add(fieldName);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }

                // Read all rows from the recordset
                resultExtractionStarting?.Invoke();
                while (!recordset.EOF)
                {
                    ct.ThrowIfCancellationRequested();
                    var row = new List<object?>();

                    for (int i = 0; i < fieldCount; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        ADODB.Field? field = null;
                        try
                        {
                            field = fields[i];
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
            catch (Exception exception)
            {
                operationFailure = exception;
                throw;
            }
            finally
            {
                try
                {
                    CloseRecordset(recordset, operationFailure);
                }
                finally
                {
                    ComUtilities.Release(ref fields);
                    ComUtilities.Release(ref recordset);
                    ComUtilities.Release(ref adoConnection);
                    ComUtilities.Release(ref modelConn);
                    ComUtilities.Release(ref dataModelConn);
                    ComUtilities.Release(ref model);
                }
            }

            return result;
        }, timeoutCts.Token);
    }

    [SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Provider diagnostics are secondary to the original DAX COM failure; any diagnostic-read exception is included in the returned message while the original COM exception remains the inner exception.")]
    private static string GetDaxEvaluationFailureMessage(ADODB._Connection adoConnection, COMException failure)
    {
        ADODB.Errors? errors = null;
        var details = new List<string>();
        int? errorCount = null;
        try
        {
            errors = adoConnection.Errors;
            errorCount = errors.Count;
            for (int index = 0; index < errorCount; index++)
            {
                ADODB.Error? error = null;
                try
                {
                    error = errors[index];
                    string description = error.Description?.ToString() ?? string.Empty;
                    string source = error.Source?.ToString() ?? string.Empty;
                    string number = error.Number.ToString("X8",
                        System.Globalization.CultureInfo.InvariantCulture);
                    if (!string.IsNullOrWhiteSpace(description) ||
                        !string.IsNullOrWhiteSpace(source) ||
                        number != "00000000")
                    {
                        var identity = string.IsNullOrWhiteSpace(source)
                            ? $"0x{number}"
                            : $"0x{number} from {source}";
                        details.Add(string.IsNullOrWhiteSpace(description)
                            ? identity
                            : $"{identity}: {description}");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref error);
                }
            }
        }
        catch (Exception diagnosticsFailure)
        {
            details.Add(
                $"The provider error collection could not be read " +
                $"({diagnosticsFailure.GetType().Name}): {diagnosticsFailure.Message}");
        }
        finally
        {
            ComUtilities.Release(ref errors);
        }

        if (details.Count == 0)
        {
            details.Add(errorCount == 0
                ? "The ADO provider returned no error records."
                : $"The ADO provider returned {errorCount?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "an unreadable number of"} error records without details.");
        }

        string hresult = unchecked((uint)failure.HResult).ToString("X8",
            System.Globalization.CultureInfo.InvariantCulture);
        var message = $"DAX evaluation failed (HRESULT 0x{hresult}): {failure.Message}";
        return $"{message} Engine details: {string.Join(" | ", details)}";
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
