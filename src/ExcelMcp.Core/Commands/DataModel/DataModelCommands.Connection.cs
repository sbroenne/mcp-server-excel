using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model connection metadata operations.
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public DataModelConnectionResult ReadConnection(IExcelBatch batch)
    {
        var result = new DataModelConnectionResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Model? model = null;
            Excel.WorkbookConnection? workbookConnection = null;
            Excel.ModelConnection? modelConnection = null;
            Excel.ModelTables? modelTables = null;
            object? commandText = null;
            try
            {
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;
                workbookConnection = model.DataModelConnection;
                modelConnection = workbookConnection.ModelConnection;
                modelTables = model.ModelTables;

                result.ModelName = model.Name ?? string.Empty;
                result.ConnectionName = workbookConnection.Name ?? string.Empty;
                result.Description = workbookConnection.Description ?? string.Empty;
                result.ConnectionTypeValue = Convert.ToInt32(workbookConnection.Type, CultureInfo.InvariantCulture);
                result.ConnectionType = ConnectionHelpers.GetConnectionTypeName(result.ConnectionTypeValue);
                result.InModel = workbookConnection.InModel;
                result.CommandTypeValue = Convert.ToInt32(modelConnection.CommandType, CultureInfo.InvariantCulture);
                result.CommandType = GetModelCommandTypeName(result.CommandTypeValue);

                commandText = modelConnection.CommandText;
                result.CommandText = GetCommandText(commandText);

                for (int i = 1; i <= modelTables.Count; i++)
                {
                    Excel.ModelTable? table = null;
                    try
                    {
                        table = modelTables.Item(i);
                        result.TableNames.Add(table.Name ?? string.Empty);
                    }
                    finally
                    {
                        ComUtilities.Release(ref table);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref commandText);
                ComUtilities.Release(ref modelTables);
                ComUtilities.Release(ref modelConnection);
                ComUtilities.Release(ref workbookConnection);
                ComUtilities.Release(ref model);
            }
        });
    }

    private static string GetModelCommandTypeName(int commandType)
    {
        return commandType switch
        {
            1 => "CUBE",
            2 => "SQL",
            3 => "TABLE",
            4 => "DEFAULT",
            5 => "LIST",
            6 => "TABLE_COLLECTION",
            7 => "EXCEL",
            8 => "DAX",
            _ => $"Unknown ({commandType})"
        };
    }

    private static string? GetCommandText(object? commandText)
    {
        return commandText switch
        {
            null => null,
            object[] values => string.Join(Environment.NewLine, values.Select(value => value?.ToString() ?? string.Empty)),
            _ => commandText.ToString()
        };
    }
}
