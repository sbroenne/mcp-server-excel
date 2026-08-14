using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc />
    public ExternalLinkListResult ListExternalLinks(IExcelBatch batch)
    {
        return batch.Execute((context, _) =>
        {
            var sources = GetExternalLinkSources(context.Book);
            return new ExternalLinkListResult
            {
                Success = true,
                Links = sources
                    .Select(source => new ExternalLinkInfo { Source = source })
                    .ToList()
            };
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateExternalLink(IExcelBatch batch, string linkSource)
    {
        ValidateLinkSource(linkSource);
        return batch.Execute((context, _) =>
        {
            var source = FindExternalLinkSource(context.Book, linkSource);
            context.Book.UpdateLink(source, Excel.XlLink.xlExcelLinks);
            return new OperationResult
            {
                Success = true,
                Action = "update-external-link",
                Message = $"External link '{source}' was updated"
            };
        });
    }

    /// <inheritdoc />
    public OperationResult BreakExternalLink(IExcelBatch batch, string linkSource)
    {
        ValidateLinkSource(linkSource);
        return batch.Execute((context, _) =>
        {
            var source = FindExternalLinkSource(context.Book, linkSource);
            context.Book.BreakLink(source, Excel.XlLinkType.xlLinkTypeExcelLinks);
            return new OperationResult
            {
                Success = true,
                Action = "break-external-link",
                Message = $"External link '{source}' was permanently broken"
            };
        });
    }

    private static List<string> GetExternalLinkSources(Excel.Workbook workbook)
    {
        var rawSources = workbook.LinkSources(Excel.XlLink.xlExcelLinks);
        if (rawSources is not Array sources)
        {
            return [];
        }

        var result = new List<string>(sources.Length);
        foreach (var source in sources)
        {
            var text = Convert.ToString(source, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(text))
            {
                result.Add(text);
            }
        }

        return result;
    }

    private static string FindExternalLinkSource(Excel.Workbook workbook, string requestedSource)
    {
        return GetExternalLinkSources(workbook)
            .FirstOrDefault(source => string.Equals(source, requestedSource, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"External Excel link '{requestedSource}' was not found.");
    }

    private static void ValidateLinkSource(string linkSource)
    {
        if (string.IsNullOrWhiteSpace(linkSource))
        {
            throw new ArgumentException("External link source cannot be empty.", nameof(linkSource));
        }
    }
}
