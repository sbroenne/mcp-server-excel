using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryPrivacySettingsResult GetPrivacySettings(IExcelBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        var result = batch.Execute((context, _) =>
        {
            Excel.Queries? queries = null;
            try
            {
                queries = context.Book.Queries;
                return CreatePrivacySettingsResult(queries.FastCombine);
            }
            finally
            {
                ComUtilities.Release(ref queries);
            }
        });
        result.FilePath = batch.WorkbookPath;
        return result;
    }

    /// <inheritdoc />
    public PowerQueryPrivacySettingsResult SetPrivacySettings(
        IExcelBatch batch,
        bool ignorePrivacyLevels)
    {
        ArgumentNullException.ThrowIfNull(batch);

        var result = batch.Execute((context, _) =>
        {
            Excel.Queries? queries = null;
            try
            {
                queries = context.Book.Queries;
                queries.FastCombine = ignorePrivacyLevels;
                bool actualValue = queries.FastCombine;
                if (actualValue != ignorePrivacyLevels)
                {
                    throw new InvalidOperationException(
                        $"Excel did not confirm the requested Power Query privacy setting. " +
                        $"Requested ignorePrivacyLevels={ignorePrivacyLevels}, actual value is {actualValue}.");
                }

                return CreatePrivacySettingsResult(actualValue);
            }
            finally
            {
                ComUtilities.Release(ref queries);
            }
        });
        result.FilePath = batch.WorkbookPath;
        return result;
    }

    private static PowerQueryPrivacySettingsResult CreatePrivacySettingsResult(
        bool ignorePrivacyLevels) =>
        new()
        {
            Success = true,
            IgnorePrivacyLevels = ignorePrivacyLevels,
            Message = ignorePrivacyLevels
                ? "This workbook is configured to ignore privacy levels when combining sources. Data from different sources may be combined and sent between them. Excel's global privacy option may override this setting."
                : "This workbook is configured not to ignore privacy levels when combining sources. Excel's global privacy option may override this setting."
        };
}
