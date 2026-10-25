using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class DataModelCommands
{
    /// <inheritdoc />
    public RenameResult RenameTable(IExcelBatch batch, string oldName, string newName)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new RenameResult
            {
                ObjectType = "data-model-table",
                OldName = oldName,
                NewName = newName,
                NormalizedOldName = RenameNameRules.Normalize(oldName),
                NormalizedNewName = RenameNameRules.Normalize(newName)
            };
            if (RenameNameRules.IsEmpty(result.NormalizedOldName))
            {
                result.ErrorMessage = "Old table name cannot be empty or whitespace.";
                return result;
            }
            if (RenameNameRules.IsEmpty(result.NormalizedNewName))
            {
                result.ErrorMessage = "New table name cannot be empty or whitespace.";
                return result;
            }
            if (RenameNameRules.IsNoOp(result.NormalizedOldName, result.NormalizedNewName))
            {
                result.Success = true;
                return result;
            }
            ct.ThrowIfCancellationRequested();
            if (!HasDataModelTables(ctx.Book))
            {
                result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                return result;
            }

            Excel.Model? model = null;
            Excel.ModelTable? table = null;
            try
            {
                model = ctx.Book.Model;
                table = FindModelTable(model, result.NormalizedOldName);
                if (table is null)
                {
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(result.NormalizedOldName);
                    return result;
                }
                var existingNames = new List<string>();
                ForEachTable(model, (candidate, _) =>
                {
                    ct.ThrowIfCancellationRequested();
                    existingNames.Add(candidate.Name ?? string.Empty);
                });
                if (RenameNameRules.HasConflict(existingNames, result.NormalizedNewName, result.NormalizedOldName))
                {
                    result.ErrorMessage = $"A table named '{result.NormalizedNewName}' already exists (case-insensitive match).";
                    return result;
                }

                // ModelTable.Name is read-only; renaming its source cannot change the cached model identity.
                result.ErrorMessage = $"Cannot rename Data Model table '{result.NormalizedOldName}': " +
                    "Excel's Data Model table names are immutable after creation. " +
                    "No workbook changes were made. Renaming its source Power Query or connection cannot rename the Data Model table.";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }
        });
    }
}
