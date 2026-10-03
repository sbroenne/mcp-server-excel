using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class ConditionalFormattingCommands
{
    /// <inheritdoc/>
    public ConditionalFormatListResult UpdateRule(IExcelBatch batch, string sheetName,
        int rulePriority, string expectedFingerprint, ConditionalRuleUpdateOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ValidateUpdateValues(options);
        return EditSelectedRule(batch, sheetName, rulePriority, expectedFingerprint,
            (rule, current, sheet, book, count) => ApplyRuleUpdate(rule, current, sheet, book, options));
    }

    /// <inheritdoc/>
    public ConditionalFormatListResult DeleteRule(IExcelBatch batch, string sheetName,
        int rulePriority, string expectedFingerprint) =>
        EditSelectedRule(batch, sheetName, rulePriority, expectedFingerprint,
            (rule, current, sheet, book, count) => DeleteNativeRule(rule));

    /// <inheritdoc/>
    public ConditionalFormatListResult SetRulePriority(IExcelBatch batch, string sheetName,
        int rulePriority, string expectedFingerprint, int newPriority)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(newPriority, 1);
        return EditSelectedRule(batch, sheetName, rulePriority, expectedFingerprint,
            (rule, current, sheet, book, count) =>
            {
                ArgumentOutOfRangeException.ThrowIfGreaterThan(newPriority, count);
                if (current.Priority != newPriority)
                    SetNativePriority(rule, newPriority);
                if (GetNativePriority(rule) != newPriority)
                    throw new InvalidOperationException("Excel did not apply the requested rule priority. List worksheet rules again before retrying.");
            });
    }

    private static ConditionalFormatListResult EditSelectedRule(IExcelBatch batch, string sheetName,
        int priority, string fingerprint, Action<object, ConditionalFormatRuleInfo, Excel.Worksheet, Excel.Workbook, int> edit)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(priority, 1);
        ArgumentException.ThrowIfNullOrWhiteSpace(fingerprint);
        return batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            Excel.FormatConditions? conditions = null;
            object? selected = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet as Excel.Worksheet ??
                        throw new InvalidOperationException("The active sheet is not a worksheet.")
                    : (Excel.Worksheet)sheets.Item[sheetName];
                cells = sheet.Cells;
                conditions = cells.FormatConditions;
                var rules = ReadFormatConditions(conditions, ct);
                var current = rules.SingleOrDefault(item => item.Priority == priority);
                if (current is null || !string.Equals(current.Fingerprint, fingerprint, StringComparison.Ordinal))
                    throw new InvalidOperationException("The selected rule changed or was removed. List worksheet rules again and use its current priority and fingerprint.");
                for (int i = 1; i <= conditions.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    object? candidate = null;
                    try
                    {
                        candidate = conditions.Item(i);
                        if (GetNativePriority(candidate) == priority)
                        {
                            selected = candidate;
                            candidate = null;
                            break;
                        }
                    }
                    finally { ComUtilities.Release(ref candidate); }
                }
                if (selected is null)
                    throw new InvalidOperationException("The selected rule changed. List worksheet rules again.");
                ct.ThrowIfCancellationRequested();
                int lastPriority = Math.Max(conditions.Count, rules.Max(item => item.Priority ?? 0));
                edit(selected, current, sheet, ctx.Book, lastPriority);
                return new ConditionalFormatListResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    Rules = ReadFormatConditions(conditions, ct)
                };
            }
            finally
            {
                ComUtilities.Release(ref selected);
                ComUtilities.Release(ref conditions);
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    private static int GetNativePriority(object rule) => rule switch
    {
        Excel.FormatCondition item => item.Priority,
        Excel.ColorScale item => item.Priority,
        Excel.Databar item => item.Priority,
        Excel.IconSetCondition item => item.Priority,
        Excel.Top10 item => item.Priority,
        Excel.AboveAverage item => item.Priority,
        Excel.UniqueValues item => item.Priority,
        _ => throw new NotSupportedException("This native conditional-rule type is not supported for editing.")
    };

    private static void SetNativePriority(object rule, int priority)
    {
        switch (rule)
        {
            case Excel.FormatCondition item: item.Priority = priority; break;
            case Excel.ColorScale item: item.Priority = priority; break;
            case Excel.Databar item: item.Priority = priority; break;
            case Excel.IconSetCondition item: item.Priority = priority; break;
            case Excel.Top10 item: item.Priority = priority; break;
            case Excel.AboveAverage item: item.Priority = priority; break;
            case Excel.UniqueValues item: item.Priority = priority; break;
            default: throw new NotSupportedException("This native conditional-rule type is not supported for editing.");
        }
    }

    private static void DeleteNativeRule(object rule)
    {
        switch (rule)
        {
            case Excel.FormatCondition item: item.Delete(); break;
            case Excel.ColorScale item: item.Delete(); break;
            case Excel.Databar item: item.Delete(); break;
            case Excel.IconSetCondition item: item.Delete(); break;
            case Excel.Top10 item: item.Delete(); break;
            case Excel.AboveAverage item: item.Delete(); break;
            case Excel.UniqueValues item: item.Delete(); break;
            default: throw new NotSupportedException("This native conditional-rule type is not supported for editing.");
        }
    }

    private static void SetNativeStopIfTrue(object rule, bool value)
    {
        switch (rule)
        {
            case Excel.FormatCondition item: item.StopIfTrue = value; break;
            case Excel.Top10 item: item.StopIfTrue = value; break;
            case Excel.AboveAverage item: item.StopIfTrue = value; break;
            case Excel.UniqueValues item: item.StopIfTrue = value; break;
            default: throw new NotSupportedException("StopIfTrue is not available for this native conditional-rule type.");
        }
    }

    private static void SetNativeAppliesTo(object rule, Excel.Range range)
    {
        switch (rule)
        {
            case Excel.FormatCondition item: item.ModifyAppliesToRange(range); break;
            case Excel.ColorScale item: item.ModifyAppliesToRange(range); break;
            case Excel.Databar item: item.ModifyAppliesToRange(range); break;
            case Excel.IconSetCondition item: item.ModifyAppliesToRange(range); break;
            case Excel.Top10 item: item.ModifyAppliesToRange(range); break;
            case Excel.AboveAverage item: item.ModifyAppliesToRange(range); break;
            case Excel.UniqueValues item: item.ModifyAppliesToRange(range); break;
            default: throw new NotSupportedException("This native conditional-rule type is not supported for editing.");
        }
    }

    private static void ValidateUpdateValues(ConditionalRuleUpdateOptions options)
    {
        var supplied = JsonSerializer.SerializeToElement(options);
        bool any = false;
        foreach (var property in supplied.EnumerateObject())
        {
            if (property.Value.ValueKind == JsonValueKind.Null)
                continue;
            any = true;
            if (property.Value.ValueKind != JsonValueKind.String)
                continue;
            var text = property.Value.GetString()!;
            ArgumentException.ThrowIfNullOrWhiteSpace(text, property.Name);
            if (property.Name.EndsWith("Color", StringComparison.Ordinal))
                _ = FormattingHelpers.ParseColor(text);
            else if (property.Name.EndsWith("Type", StringComparison.Ordinal) && property.Name != nameof(options.OperatorType))
                _ = ParseConditionValueType(text);
        }
        if (!any)
            throw new ArgumentException("Supply at least one rule setting.", nameof(options));
        if (options.OperatorType is not null)
            _ = ParseConditionalFormattingOperator(options.OperatorType);
        if (options.InteriorPattern is not null)
            _ = ParseInteriorPattern(options.InteriorPattern);
        if (options.BorderStyle is not null)
            _ = FormattingHelpers.ParseBorderStyle(options.BorderStyle);
        if (options.DataBarDirection is not null)
            _ = ParseDataBarDirection(options.DataBarDirection);
        if (options.IconSetId is not null)
            _ = ParseIconSetId(options.IconSetId);
        if (options.TopBottom is not null)
            _ = ParseTopBottom(options.TopBottom);
        if (options.AboveBelow is not null)
            _ = ParseAboveBelow(options.AboveBelow);
        if (options.DatePeriod is not null)
            _ = ParseTimePeriod(options.DatePeriod);
        if (options.Rank is <= 0 or > 1000)
            throw new ArgumentOutOfRangeException(nameof(options), "Rank must be between 1 and 1000.");
        if (options.StandardDeviations is < 1 or > 3)
            throw new ArgumentOutOfRangeException(nameof(options), "StandardDeviations must be between 1 and 3.");
    }

    private static void ValidateOptionsForType(ConditionalRuleUpdateOptions options, string type)
    {
        foreach (var property in JsonSerializer.SerializeToElement(options).EnumerateObject())
        {
            if (property.Value.ValueKind == JsonValueKind.Null)
                continue;
            bool allowed = property.Name switch
            {
                nameof(options.AppliesTo) => true,
                nameof(options.StopIfTrue) => type is not ("colorScale" or "dataBar" or "iconSet"),
                nameof(options.OperatorType) => type == "cellValue",
                nameof(options.Formula1) => type is "cellValue" or "expression",
                nameof(options.Formula2) => type == "cellValue",
                nameof(options.Rank) or nameof(options.Top10Percent) or nameof(options.TopBottom) => type == "top10",
                nameof(options.AboveBelow) or nameof(options.StandardDeviations) => type == "aboveAverage",
                nameof(options.DatePeriod) => type == "timePeriod",
                nameof(options.DuplicateValues) => type == "uniqueValues",
                _ when property.Name.StartsWith("ColorScale", StringComparison.Ordinal) => type == "colorScale",
                _ when property.Name.StartsWith("DataBar", StringComparison.Ordinal) => type == "dataBar",
                _ when property.Name.StartsWith("Icon", StringComparison.Ordinal) => type == "iconSet",
                _ => type is not ("colorScale" or "dataBar" or "iconSet")
            };
            if (!allowed)
                throw new ArgumentException($"{property.Name} is not applicable to a {type} rule.", nameof(options));
        }
    }

    private static void ApplyRuleUpdate(object rule, ConditionalFormatRuleInfo current,
        Excel.Worksheet sheet, Excel.Workbook book, ConditionalRuleUpdateOptions options)
    {
        ValidateOptionsForType(options, current.Type);
        int? basicOperator = null;
        if (options.OperatorType is not null || options.Formula1 is not null || options.Formula2 is not null)
            basicOperator = ValidateAndParseBasicRuleArguments(NormalizeRuleType(current.Type),
                options.OperatorType ?? current.Operator, options.Formula1 ?? current.Formula1,
                options.Formula2 ?? current.Formula2);
        if (current.Type == "top10" && (options.Top10Percent ?? current.Top10?.Percent) == true &&
            (options.Rank ?? current.Top10?.Rank) > 100)
            throw new ArgumentException("Percentage rank cannot exceed 100.", nameof(options));
        if (options.StandardDeviations.HasValue &&
            ParseAboveBelow(options.AboveBelow ?? current.AboveBelow ??
                throw new InvalidOperationException("Excel did not expose the average comparison.")) is not (4 or 5))
            throw new ArgumentException("StandardDeviations requires an above/below-standard-deviation comparison.", nameof(options));
        if (options.Formula2 is not null && basicOperator is not (1 or 2))
            throw new ArgumentException("Formula2 is applicable only to between/notBetween comparisons.", nameof(options));
        ValidateVisualUpdate(current, options);
        Excel.Range? appliesTo = null;
        try
        {
            if (options.AppliesTo is not null)
            {
                appliesTo = RangeHelpers.ResolveRange(book, sheet.Name, options.AppliesTo, out var error) as Excel.Range ??
                    throw new ArgumentException(error ?? "The applies-to range could not be resolved.", nameof(options));
                SetNativeAppliesTo(rule, appliesTo);
            }
            switch (rule)
            {
                case Excel.FormatCondition basic:
                    if (basicOperator.HasValue)
                        basic.Modify((Excel.XlFormatConditionType)basic.Type, basicOperator.Value,
                            options.Formula1 ?? current.Formula1, options.Formula2 ?? current.Formula2);
                    if (options.DatePeriod is not null)
                        basic.DateOperator = (Excel.XlTimePeriods)ParseTimePeriod(options.DatePeriod);
                    if (options.StopIfTrue.HasValue)
                        basic.StopIfTrue = options.StopIfTrue.Value;
                    break;
                case Excel.Top10 top:
                    if (options.Top10Percent.HasValue) top.Percent = options.Top10Percent.Value;
                    if (options.Rank.HasValue) top.Rank = options.Rank.Value;
                    if (options.TopBottom is not null) top.TopBottom = (Excel.XlTopBottom)ParseTopBottom(options.TopBottom);
                    if (options.StopIfTrue.HasValue) top.StopIfTrue = options.StopIfTrue.Value;
                    break;
                case Excel.AboveAverage average:
                    if (options.AboveBelow is not null) average.AboveBelow = (Excel.XlAboveBelow)ParseAboveBelow(options.AboveBelow);
                    if (options.StandardDeviations.HasValue) average.NumStdDev = options.StandardDeviations.Value;
                    if (options.StopIfTrue.HasValue) average.StopIfTrue = options.StopIfTrue.Value;
                    break;
                case Excel.UniqueValues unique:
                    if (options.DuplicateValues.HasValue)
                        unique.DupeUnique = options.DuplicateValues.Value ? Excel.XlDupeUnique.xlDuplicate : Excel.XlDupeUnique.xlUnique;
                    if (options.StopIfTrue.HasValue) unique.StopIfTrue = options.StopIfTrue.Value;
                    break;
                case Excel.ColorScale scale: UpdateColorScale(scale, options); break;
                case Excel.Databar bar: UpdateDataBar(bar, options); break;
                case Excel.IconSetCondition icons: UpdateIconSet(icons, book, options); break;
                default: throw new NotSupportedException("This native conditional-rule type is not supported for editing.");
            }
            ApplyRuleFormatting(rule, options.InteriorColor, options.InteriorPattern, options.FontColor,
                options.FontBold, options.FontItalic, options.BorderStyle, options.BorderColor);
        }
        finally
        {
            ComUtilities.Release(ref appliesTo);
        }
    }
}
