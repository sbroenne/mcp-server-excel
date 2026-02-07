using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public RenameResult Rename(IExcelBatch batch, string oldName, string newName)
    {
        return batch.Execute((ctx, _) =>
        {
            var result = new RenameResult
            {
                ObjectType = "power-query",
                OldName = oldName,
                NewName = newName,
                NormalizedOldName = RenameNameRules.Normalize(oldName),
                NormalizedNewName = RenameNameRules.Normalize(newName)
            };

            // Validate new name is not empty
            if (RenameNameRules.IsEmpty(result.NormalizedNewName))
            {
                result.Success = false;
                result.ErrorMessage = "New query name cannot be empty or whitespace.";
                return result;
            }

            // No-op when normalized names are exactly equal
            if (RenameNameRules.IsNoOp(result.NormalizedOldName, result.NormalizedNewName))
            {
                result.Success = true;
                return result;
            }

            dynamic? queries = null;
            dynamic? targetQuery = null;
            try
            {
                queries = ctx.Book.Queries;

                // Find target query (case-sensitive exact match first)
                targetQuery = ComUtilities.FindQuery(ctx.Book, result.NormalizedOldName);
                if (targetQuery == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Power Query '{result.NormalizedOldName}' not found.";
                    return result;
                }

                // Collect existing query names for conflict detection
                var existingNames = new List<string>();
                int count = queries.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? q = null;
                    try
                    {
                        q = queries.Item(i);
                        existingNames.Add((string)q.Name);
                    }
                    finally
                    {
                        if (q != null) ComUtilities.Release(ref q!);
                    }
                }

                // Check for conflicts (case-insensitive, excluding target)
                if (RenameNameRules.HasConflict(existingNames, result.NormalizedNewName, result.NormalizedOldName))
                {
                    result.Success = false;
                    result.ErrorMessage = $"A query named '{result.NormalizedNewName}' already exists (case-insensitive match).";
                    return result;
                }

                // Attempt COM rename (includes case-only renames)
                targetQuery.Name = result.NormalizedNewName;

                result.Success = true;
                return result;
            }
            finally
            {
                if (targetQuery != null) ComUtilities.Release(ref targetQuery!);
                if (queries != null) ComUtilities.Release(ref queries!);
            }
        });
    }
}


