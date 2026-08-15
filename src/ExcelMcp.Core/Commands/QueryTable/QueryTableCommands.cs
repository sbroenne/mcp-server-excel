using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Typed Excel PIA implementation of worksheet QueryTable operations.
/// </summary>
public sealed class QueryTableCommands : IQueryTableCommands
{
    /// <inheritdoc />
    public QueryTableListResult List(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var result = new QueryTableListResult { FilePath = batch.WorkbookPath };
            Excel.Sheets? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int sheetIndex = 1; sheetIndex <= sheets.Count; sheetIndex++)
                {
                    Excel.Worksheet? sheet = null;
                    Excel.QueryTables? queryTables = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets.Item[sheetIndex];
                        queryTables = sheet.QueryTables;
                        for (int queryIndex = 1; queryIndex <= queryTables.Count; queryIndex++)
                        {
                            Excel.QueryTable? queryTable = null;
                            Excel.Range? destination = null;
                            try
                            {
                                queryTable = queryTables.Item(queryIndex);
                                destination = queryTable.Destination;
                                result.QueryTables.Add(new QueryTableInfo
                                {
                                    Name = queryTable.Name,
                                    SheetName = sheet.Name,
                                    Destination = destination.Address[false, false],
                                    SourceType = GetSourceType(queryTable.QueryType),
                                    IsRefreshing = queryTable.Refreshing
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref destination);
                                ComUtilities.Release(ref queryTable);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref sheet);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public QueryTableViewResult View(IExcelBatch batch, string sheetName, string queryTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            Excel.Range? destination = null;
            try
            {
                destination = queryTable.Destination;
                var sourceType = GetSourceType(queryTable.QueryType);
                return new QueryTableViewResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Name = queryTable.Name,
                    SheetName = sheetName,
                    Destination = destination.Address[false, false],
                    SourceType = sourceType,
                    Connection = ConnectionStringSanitizer.Sanitize(
                        Convert.ToString(queryTable.Connection)) ?? string.Empty,
                    BackgroundQuery = queryTable.BackgroundQuery,
                    RefreshOnFileOpen = queryTable.RefreshOnFileOpen,
                    RefreshPeriod = queryTable.RefreshPeriod,
                    AdjustColumnWidth = queryTable.AdjustColumnWidth,
                    PreserveFormatting = queryTable.PreserveFormatting,
                    Delimiter = sourceType == "text" ? GetTextDelimiter(queryTable) : null,
                    Encoding = sourceType == "text" ? Convert.ToInt32(queryTable.TextFilePlatform) : null,
                    WebSelectionType = sourceType == "web" ? ToWebSelectionType(queryTable.WebSelectionType) : null,
                    WebTables = sourceType == "web" ? queryTable.WebTables : null,
                    WebFormatting = sourceType == "web" ? ToWebFormatting(queryTable.WebFormatting) : null
                };
            }
            finally
            {
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult CreateText(
        IExcelBatch batch,
        string queryTableName,
        string sourcePath,
        string sheetName,
        string destinationAddress,
        string delimiter = ",",
        string textQualifier = "double-quote",
        int encoding = 65001,
        bool hasHeaders = true)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException("Text import source file not found.", sourcePath);
        }

        if (delimiter.Length != 1)
        {
            throw new ArgumentException("delimiter must contain exactly one character.", nameof(delimiter));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? destination = null;
            Excel.QueryTables? queryTables = null;
            Excel.QueryTable? queryTable = null;
            try
            {
                sheet = FindSheet(ctx.Book, sheetName);
                EnsureQueryTableNameAvailable(sheet, queryTableName);
                destination = sheet.Range[destinationAddress];
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add($"TEXT;{Path.GetFullPath(sourcePath)}", destination);
                queryTable.Name = queryTableName;
                queryTable.FieldNames = hasHeaders;
                queryTable.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                queryTable.TextFilePlatform = encoding;
                queryTable.TextFileTextQualifier = ParseTextQualifier(textQualifier);
                SetTextDelimiter(queryTable, delimiter[0]);
                queryTable.BackgroundQuery = false;
                queryTable.Refresh(false);
                return Success(batch.WorkbookPath, "create-text");
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult CreateWeb(
        IExcelBatch batch,
        string queryTableName,
        string url,
        string sheetName,
        string destinationAddress,
        string selectionType = "all-tables",
        string? webTables = null,
        string formatting = "none")
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out _))
        {
            throw new ArgumentException("url must be an absolute HTTP, HTTPS, or file URI.", nameof(url));
        }

        var parsedSelection = ParseWebSelectionType(selectionType);
        if (parsedSelection == Excel.XlWebSelectionType.xlSpecifiedTables && string.IsNullOrWhiteSpace(webTables))
        {
            throw new ArgumentException("webTables is required when selectionType is 'specified-tables'.", nameof(webTables));
        }

        var parsedFormatting = ParseWebFormatting(formatting);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? destination = null;
            Excel.QueryTables? queryTables = null;
            Excel.QueryTable? queryTable = null;
            try
            {
                sheet = FindSheet(ctx.Book, sheetName);
                EnsureQueryTableNameAvailable(sheet, queryTableName);
                destination = sheet.Range[destinationAddress];
                queryTables = sheet.QueryTables;
                queryTable = queryTables.Add($"URL;{url}", destination);
                queryTable.Name = queryTableName;
                queryTable.WebSelectionType = parsedSelection;
                queryTable.WebFormatting = parsedFormatting;
                if (parsedSelection == Excel.XlWebSelectionType.xlSpecifiedTables)
                {
                    queryTable.WebTables = webTables;
                }

                queryTable.BackgroundQuery = false;
                queryTable.Refresh(false);
                return Success(batch.WorkbookPath, "create-web");
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetProperties(
        IExcelBatch batch,
        string sheetName,
        string queryTableName,
        bool? backgroundQuery = null,
        bool? refreshOnFileOpen = null,
        int? refreshPeriod = null,
        bool? adjustColumnWidth = null,
        bool? preserveFormatting = null)
    {
        if (refreshPeriod < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(refreshPeriod), "refreshPeriod cannot be negative.");
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            try
            {
                if (backgroundQuery.HasValue) queryTable.BackgroundQuery = backgroundQuery.Value;
                if (refreshOnFileOpen.HasValue) queryTable.RefreshOnFileOpen = refreshOnFileOpen.Value;
                if (refreshPeriod.HasValue) queryTable.RefreshPeriod = refreshPeriod.Value;
                if (adjustColumnWidth.HasValue) queryTable.AdjustColumnWidth = adjustColumnWidth.Value;
                if (preserveFormatting.HasValue) queryTable.PreserveFormatting = preserveFormatting.Value;
                return Success(batch.WorkbookPath, "set-properties");
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Refresh(IExcelBatch batch, string sheetName, string queryTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            try
            {
                queryTable.Refresh(false);
                return Success(batch.WorkbookPath, "refresh");
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    /// <inheritdoc />
    public RefreshStatusResult GetRefreshStatus(IExcelBatch batch, string sheetName, string queryTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            try
            {
                return new RefreshStatusResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SupportsRefreshStatus = true,
                    IsRefreshing = queryTable.Refreshing
                };
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    /// <inheritdoc />
    public RefreshCancellationResult CancelRefresh(IExcelBatch batch, string sheetName, string queryTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            try
            {
                bool wasRefreshing = queryTable.Refreshing;
                if (wasRefreshing)
                {
                    queryTable.CancelRefresh();
                }

                return new RefreshCancellationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SupportsCancellation = true,
                    WasRefreshing = wasRefreshing,
                    Cancelled = wasRefreshing
                };
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string sheetName, string queryTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.QueryTable? queryTable = FindQueryTable(ctx.Book, sheetName, queryTableName);
            try
            {
                queryTable.Delete();
                return Success(batch.WorkbookPath, "delete");
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }
        });
    }

    private static Excel.Worksheet FindSheet(Excel.Workbook workbook, string sheetName)
    {
        return ComUtilities.FindSheet(workbook, sheetName)
            ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
    }

    private static Excel.QueryTable FindQueryTable(
        Excel.Workbook workbook,
        string sheetName,
        string queryTableName)
    {
        Excel.Worksheet? sheet = null;
        Excel.QueryTables? queryTables = null;
        try
        {
            sheet = FindSheet(workbook, sheetName);
            queryTables = sheet.QueryTables;
            for (int index = 1; index <= queryTables.Count; index++)
            {
                Excel.QueryTable? candidate = null;
                try
                {
                    candidate = queryTables.Item(index);
                    if (candidate.Name.Equals(queryTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        var match = candidate;
                        candidate = null;
                        return match;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidate);
                }
            }

            throw new InvalidOperationException($"QueryTable '{queryTableName}' was not found on sheet '{sheetName}'.");
        }
        finally
        {
            ComUtilities.Release(ref queryTables);
            ComUtilities.Release(ref sheet);
        }
    }

    private static void EnsureQueryTableNameAvailable(Excel.Worksheet sheet, string queryTableName)
    {
        Excel.QueryTables? queryTables = null;
        try
        {
            queryTables = sheet.QueryTables;
            for (int index = 1; index <= queryTables.Count; index++)
            {
                Excel.QueryTable? candidate = null;
                try
                {
                    candidate = queryTables.Item(index);
                    if (candidate.Name.Equals(queryTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            $"QueryTable '{queryTableName}' already exists on sheet '{sheet.Name}'.");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidate);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref queryTables);
        }
    }

    private static string GetSourceType(Excel.XlQueryType queryType)
    {
        return queryType switch
        {
            Excel.XlQueryType.xlTextImport => "text",
            Excel.XlQueryType.xlWebQuery => "web",
            Excel.XlQueryType.xlOLEDBQuery or Excel.XlQueryType.xlODBCQuery => "database",
            _ => "other"
        };
    }

    private static void SetTextDelimiter(Excel.QueryTable queryTable, char delimiter)
    {
        queryTable.TextFileCommaDelimiter = delimiter == ',';
        queryTable.TextFileSemicolonDelimiter = delimiter == ';';
        queryTable.TextFileTabDelimiter = delimiter == '\t';
        queryTable.TextFileSpaceDelimiter = delimiter == ' ';
        queryTable.TextFileOtherDelimiter = delimiter is ',' or ';' or '\t' or ' '
            ? string.Empty
            : delimiter.ToString();
    }

    private static string GetTextDelimiter(Excel.QueryTable queryTable)
    {
        if (queryTable.TextFileCommaDelimiter) return ",";
        if (queryTable.TextFileSemicolonDelimiter) return ";";
        if (queryTable.TextFileTabDelimiter) return "\t";
        if (queryTable.TextFileSpaceDelimiter) return " ";
        return Convert.ToString(queryTable.TextFileOtherDelimiter) ?? string.Empty;
    }

    private static Excel.XlTextQualifier ParseTextQualifier(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "double-quote" => Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
            "single-quote" => Excel.XlTextQualifier.xlTextQualifierSingleQuote,
            "none" => Excel.XlTextQualifier.xlTextQualifierNone,
            _ => throw new ArgumentException(
                "textQualifier must be 'double-quote', 'single-quote', or 'none'.",
                nameof(value))
        };
    }

    private static Excel.XlWebSelectionType ParseWebSelectionType(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "entire-page" => Excel.XlWebSelectionType.xlEntirePage,
            "all-tables" => Excel.XlWebSelectionType.xlAllTables,
            "specified-tables" => Excel.XlWebSelectionType.xlSpecifiedTables,
            _ => throw new ArgumentException(
                "selectionType must be 'entire-page', 'all-tables', or 'specified-tables'.",
                nameof(value))
        };
    }

    private static string ToWebSelectionType(Excel.XlWebSelectionType value)
    {
        return value switch
        {
            Excel.XlWebSelectionType.xlEntirePage => "entire-page",
            Excel.XlWebSelectionType.xlAllTables => "all-tables",
            Excel.XlWebSelectionType.xlSpecifiedTables => "specified-tables",
            _ => value.ToString()
        };
    }

    private static Excel.XlWebFormatting ParseWebFormatting(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "none" => Excel.XlWebFormatting.xlWebFormattingNone,
            "rich-text" => Excel.XlWebFormatting.xlWebFormattingRTF,
            "all" => Excel.XlWebFormatting.xlWebFormattingAll,
            _ => throw new ArgumentException(
                "formatting must be 'none', 'rich-text', or 'all'.",
                nameof(value))
        };
    }

    private static string ToWebFormatting(Excel.XlWebFormatting value)
    {
        return value switch
        {
            Excel.XlWebFormatting.xlWebFormattingNone => "none",
            Excel.XlWebFormatting.xlWebFormattingRTF => "rich-text",
            Excel.XlWebFormatting.xlWebFormattingAll => "all",
            _ => value.ToString()
        };
    }

    private static OperationResult Success(string workbookPath, string action)
    {
        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath,
            Action = action
        };
    }
}
