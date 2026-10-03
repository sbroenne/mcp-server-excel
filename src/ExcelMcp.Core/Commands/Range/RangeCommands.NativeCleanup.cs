using System.Globalization;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    /// <inheritdoc />
    public RemoveDuplicatesResult RemoveDuplicates(IExcelBatch batch, string sheetName,
        string rangeAddress, List<int> keyColumns, bool hasHeaders = true)
    {
        ArgumentNullException.ThrowIfNull(keyColumns);
        if (keyColumns.Count is < 1 or > 16_383 || keyColumns.Distinct().Count() != keyColumns.Count ||
            keyColumns.Any(column => column < 1))
            throw new ArgumentException(
                "keyColumns must contain 1 through 16383 distinct positive column indices. " +
                "Native exact counting needs one additional temporary identity column.", nameof(keyColumns));
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? source = null;
            Excel.Range? remaining = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                source = ResolveFillRange(ctx, sheetName, rangeAddress);
                var size = GetContentDimensions(source);
                if (keyColumns.Any(column => column > size.Columns))
                    throw new ArgumentException("A key column is outside the selected rectangle.", nameof(keyColumns));
                int headerRows = hasHeaders ? 1 : 0;
                if (size.Rows <= headerRows)
                    throw new ArgumentException("The selection must contain at least one data row.", nameof(rangeAddress));

                int retainedRows = WithNativeCleanupWorksheet(ctx, ct, (scratch, token) =>
                    CountNativeUniqueRows(source, scratch, size.Rows, keyColumns, hasHeaders, token));
                ct.ThrowIfCancellationRequested();
                source.RemoveDuplicates(keyColumns.Select(column => (object)column).ToArray(),
                    hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
                remaining = source.Resize[retainedRows + headerRows, size.Columns];
                return new RemoveDuplicatesResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "remove-duplicates",
                    Success = true,
                    SourceRange = source.Address,
                    RemovedRows = size.Rows - headerRows - retainedRows,
                    RemainingRows = retainedRows,
                    RemainingRange = remaining.Address
                };
            }
            finally
            {
                ComUtilities.Release(ref remaining);
                ComUtilities.Release(ref source);
            }
        });
    }

    /// <inheritdoc />
    public TextToColumnsResult TextToColumns(IExcelBatch batch, string sheetName, string sourceRange,
        string destinationCell, TextToColumnsOptions options,
        OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty)
    {
        ValidateTextParsingOptions(options);
        ValidateOverwritePolicy(overwritePolicy);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? source = null;
            Excel.Range? anchor = null;
            Excel.Range? destination = null;
            Excel.Worksheet? sourceSheet = null;
            Excel.Worksheet? destinationSheet = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                source = ResolveFillRange(ctx, sheetName, sourceRange);
                anchor = ResolveFillRange(ctx, sheetName, destinationCell);
                var size = GetContentDimensions(source);
                if (size.Columns != 1 || GetContentDimensions(anchor) != (1, 1))
                    throw new ArgumentException("TextToColumns requires a single source column and a single destination cell.");
                sourceSheet = source.Worksheet;
                destinationSheet = anchor.Worksheet;
                if (!string.Equals(sourceSheet.Name, destinationSheet.Name, StringComparison.Ordinal))
                    throw new ArgumentException("TextToColumns source and destination must be on the same worksheet.");
                // TextToColumns parses formula text, not calculated formula results.
                var sourceValues = source.Formula;
                string decimalSeparator = options.DecimalSeparator ?? ctx.App.DecimalSeparator;
                string thousandsSeparator = options.ThousandsSeparator ?? ctx.App.ThousandsSeparator;
                if (decimalSeparator == thousandsSeparator)
                    throw new ArgumentException("Decimal and thousands separators must differ.");
                int width = WithNativeCleanupWorksheet(ctx, ct, (scratch, token) =>
                    MeasureNativeTextWidth(scratch, sourceValues, size.Rows, options,
                        decimalSeparator, thousandsSeparator, token));
                if ((long)anchor.Row + size.Rows - 1 > 1_048_576 || (long)anchor.Column + width - 1 > 16_384)
                    throw new ArgumentException("The complete parsed output exceeds worksheet boundaries.");
                destination = anchor.Resize[size.Rows, width];
                if (RangeMergeDiscovery.GetMergeCellsState(destination.MergeCells) != false)
                    throw new ArgumentException("The complete parsed output must not intersect merged cells.");
                int sourceRow = source.Row;
                int sourceColumn = source.Column;
                int destinationRow = destination.Row;
                int destinationColumn = destination.Column;
                EnsureDestinationWritable(ctx, destination, overwritePolicy, ct, (row, column) =>
                    destinationColumn + column != sourceColumn ||
                    destinationRow + row < sourceRow || destinationRow + row >= sourceRow + size.Rows);
                ct.ThrowIfCancellationRequested();
                ParseNativeText(source, anchor, options, decimalSeparator, thousandsSeparator);
                return new TextToColumnsResult
                {
                    FilePath = batch.WorkbookPath,
                    Action = "text-to-columns",
                    Success = true,
                    SourceRange = source.Address,
                    DestinationRange = destination.Address,
                    OutputColumns = width
                };
            }
            finally
            {
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref destinationSheet);
                ComUtilities.Release(ref sourceSheet);
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref source);
            }
        });
    }

    private static T WithNativeCleanupWorksheet<T>(ExcelContext context, CancellationToken token,
        Func<Excel.Worksheet, CancellationToken, T> operation)
    {
        Excel.Window? previousWindow = null;
        Excel.Workbooks? workbooks = null;
        Excel.Workbook? scratchBook = null;
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Exception? failure = null;
        T result = default!;
        try
        {
            token.ThrowIfCancellationRequested();
            previousWindow = context.App.ActiveWindow;
            workbooks = context.App.Workbooks;
            scratchBook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            scratchBook.Date1904 = context.Book.Date1904;
            sheets = scratchBook.Worksheets;
            sheet = (Excel.Worksheet)sheets[1];
            result = operation(sheet, token);
        }
        catch (Exception exception)
        {
            failure = exception;
        }
        finally
        {
            ComUtilities.Release(ref sheet);
            ComUtilities.Release(ref sheets);
            try
            {
                if (scratchBook is not null)
                    ExcelShutdownService.CloseTemporaryWorkbook(scratchBook);
            }
            catch (Exception exception)
            {
                failure = CombineNativeCleanupFailure(failure, exception);
            }
            finally
            {
                ComUtilities.Release(ref scratchBook);
                ComUtilities.Release(ref workbooks);
                try
                {
                    previousWindow?.Activate();
                }
                catch (Exception exception)
                {
                    failure = CombineNativeCleanupFailure(failure, exception);
                }
                finally
                {
                    ComUtilities.Release(ref previousWindow);
                }
            }
        }
        if (failure is not null)
            ExceptionDispatchInfo.Capture(failure).Throw();
        return result;
    }

    private static Exception CombineNativeCleanupFailure(Exception? primary, Exception cleanup) =>
        primary is null
            ? new InvalidOperationException("Native cleanup preflight could not restore its temporary workbook/window state.", cleanup)
            : new AggregateException("Native cleanup preflight and its cleanup both failed.", primary, cleanup);

    private static int CountNativeUniqueRows(Excel.Range source, Excel.Worksheet scratch, int rows,
        List<int> keys, bool hasHeaders, CancellationToken token)
    {
        Excel.Range? sourceColumns = null;
        Excel.Range? identity = null;
        Excel.Range? records = null;
        try
        {
            sourceColumns = source.Columns;
            for (int key = 0; key < keys.Count; key++)
            {
                token.ThrowIfCancellationRequested();
                Excel.Range? sourceColumn = null;
                Excel.Range? target = null;
                try
                {
                    sourceColumn = sourceColumns[keys[key]];
                    target = scratch.Range[$"{GetColumnLetter(key + 1)}1:{GetColumnLetter(key + 1)}{rows}"];
                    CopyNativeCleanupValuesAndFormats(sourceColumn, target, rows, token);
                }
                finally
                {
                    ComUtilities.Release(ref target);
                    ComUtilities.Release(ref sourceColumn);
                }
            }
            identity = scratch.Range[$"{GetColumnLetter(keys.Count + 1)}1:{GetColumnLetter(keys.Count + 1)}{rows}"];
            var ids = new object[rows, 1];
            for (int row = 0; row < rows; row++)
            {
                token.ThrowIfCancellationRequested();
                ids[row, 0] = row + 1;
            }
            identity.Value2 = ids;
            records = scratch.Range[$"A1:{GetColumnLetter(keys.Count + 1)}{rows}"];
            records.RemoveDuplicates(Enumerable.Range(1, keys.Count).Select(column => (object)column).ToArray(),
                hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
            object? retained = identity.Value2;
            int count = 0;
            for (int row = hasHeaders ? 1 : 0; row < rows; row++)
            {
                token.ThrowIfCancellationRequested();
                if (GetInspectionCell(retained, rows, 1, row, 0) is not null)
                    count++;
            }
            return count;
        }
        finally
        {
            ComUtilities.Release(ref records);
            ComUtilities.Release(ref identity);
            ComUtilities.Release(ref sourceColumns);
        }
    }

    private static void CopyNativeCleanupValuesAndFormats(Excel.Range source, Excel.Range target,
        int rows, CancellationToken token)
    {
        target.NumberFormat = "@";
        target.Value2 = PreserveNativeTextValues(source.Value2, rows, token);
        object? numberFormat = source.NumberFormat;
        if (numberFormat is string uniform)
        {
            target.NumberFormat = uniform;
            return;
        }
        Excel.Range? sourceCells = null;
        Excel.Range? targetCells = null;
        try
        {
            sourceCells = source.Cells;
            targetCells = target.Cells;
            for (int row = 1; row <= rows; row++)
            {
                token.ThrowIfCancellationRequested();
                Excel.Range? sourceCell = null;
                Excel.Range? targetCell = null;
                try
                {
                    sourceCell = sourceCells[row, 1];
                    targetCell = targetCells[row, 1];
                    targetCell.NumberFormat = sourceCell.NumberFormat;
                }
                finally
                {
                    ComUtilities.Release(ref targetCell);
                    ComUtilities.Release(ref sourceCell);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref targetCells);
            ComUtilities.Release(ref sourceCells);
        }
    }

    private static object?[,] PreserveNativeTextValues(object? values, int rows, CancellationToken token)
    {
        var preserved = new object?[rows, 1];
        for (int row = 0; row < rows; row++)
        {
            token.ThrowIfCancellationRequested();
            object? value = GetInspectionCell(values, rows, 1, row, 0);
            // A native entry prefix prevents Excel from consuming a leading quote or evaluating text as a formula.
            preserved[row, 0] = value is string text ? "'" + text : value;
        }
        return preserved;
    }

    private static int MeasureNativeTextWidth(Excel.Worksheet scratch, object? values, int rows,
        TextToColumnsOptions options, string decimalSeparator, string thousandsSeparator, CancellationToken token)
    {
        Excel.Range? input = null;
        Excel.Range? markers = null;
        Excel.Range? firstRow = null;
        try
        {
            input = scratch.Range[$"A1:A{rows}"];
            input.NumberFormat = "@";
            input.Value2 = PreserveNativeTextValues(values, rows, token);
            string marker;
            do
            {
                marker = Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture);
                token.ThrowIfCancellationRequested();
            }
            while (Enumerable.Range(0, rows).Any(row =>
                Convert.ToString(GetInspectionCell(values, rows, 1, row, 0), CultureInfo.InvariantCulture)?
                    .Contains(marker, StringComparison.Ordinal) == true));
            markers = scratch.Range["B1:XFD1"];
            markers.Value2 = marker;
            ParseNativeText(input, input, options, decimalSeparator, thousandsSeparator);
            firstRow = scratch.Range["A1:XFD1"];
            object? output = firstRow.Value2;
            int width = 1;
            for (int column = 1; column < 16_384; column++)
            {
                token.ThrowIfCancellationRequested();
                if (!string.Equals(GetInspectionCell(output, 1, 16_384, 0, column) as string,
                    marker, StringComparison.Ordinal))
                    width = column + 1;
            }
            return width;
        }
        finally
        {
            ComUtilities.Release(ref firstRow);
            ComUtilities.Release(ref markers);
            ComUtilities.Release(ref input);
        }
    }

    private static void ParseNativeText(Excel.Range source, Excel.Range destination,
        TextToColumnsOptions options, string decimalSeparator, string thousandsSeparator)
    {
        object fieldInfo = options.Fields is { Count: > 0 }
            ? options.Fields.Select(field => (object)new object[] { field.Position, (int)field.DataType }).ToArray()
            : Type.Missing;
        source.TextToColumns(Destination: destination,
            DataType: options.Mode == TextParsingMode.Delimited
                ? Excel.XlTextParsingType.xlDelimited : Excel.XlTextParsingType.xlFixedWidth,
            TextQualifier: options.Qualifier switch
            {
                TextFieldQualifier.DoubleQuote => Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                TextFieldQualifier.SingleQuote => Excel.XlTextQualifier.xlTextQualifierSingleQuote,
                TextFieldQualifier.None => Excel.XlTextQualifier.xlTextQualifierNone,
                _ => throw new ArgumentOutOfRangeException(nameof(options))
            },
            ConsecutiveDelimiter: options.ConsecutiveDelimiters, Tab: options.Tab, Semicolon: options.Semicolon,
            Comma: options.Comma, Space: options.Space, Other: options.OtherDelimiter is not null,
            OtherChar: options.OtherDelimiter is null ? Type.Missing : options.OtherDelimiter,
            FieldInfo: fieldInfo, DecimalSeparator: decimalSeparator, ThousandsSeparator: thousandsSeparator,
            TrailingMinusNumbers: options.TrailingMinusNumbers);
    }

    private static void ValidateTextParsingOptions(TextToColumnsOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (!Enum.IsDefined(options.Mode) || !Enum.IsDefined(options.Qualifier))
            throw new ArgumentException("Unknown text parsing mode or qualifier.", nameof(options));
        foreach (var separator in new[] { options.OtherDelimiter, options.DecimalSeparator, options.ThousandsSeparator })
            if (separator is not null && separator.Length != 1)
                throw new ArgumentException("Each explicit delimiter or number separator must be one character.", nameof(options));
        if (options.Mode == TextParsingMode.Delimited &&
            !(options.Tab || options.Semicolon || options.Comma || options.Space || options.OtherDelimiter is not null))
            throw new ArgumentException("Delimited parsing requires at least one delimiter.", nameof(options));
        if (options.Fields is not null)
        {
            if (options.Fields.Any(field => field is null || !Enum.IsDefined(field.DataType) ||
                field.Position < (options.Mode == TextParsingMode.Delimited ? 1 : 0)) ||
                options.Fields.Select(field => field.Position).Distinct().Count() != options.Fields.Count)
                throw new ArgumentException("Fields require valid distinct positions and native data types.", nameof(options));
        }
        if (options.Mode == TextParsingMode.FixedWidth)
        {
            if (options.Tab || options.Semicolon || options.Comma || options.Space ||
                options.OtherDelimiter is not null || options.ConsecutiveDelimiters)
                throw new ArgumentException("FixedWidth cannot be combined with delimiter settings.", nameof(options));
            if (options.Fields is not { Count: > 0 } || options.Fields[0].Position != 0 ||
                !options.Fields.Select(field => field.Position).SequenceEqual(options.Fields.Select(field => field.Position).Order()) ||
                options.Fields.All(field => field.DataType == TextFieldType.Skip))
                throw new ArgumentException("FixedWidth requires ascending fields beginning at zero and at least one output field.", nameof(options));
        }
    }
}
