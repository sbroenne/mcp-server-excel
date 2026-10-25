using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.CLI.Commands;

internal static class BatchCommandSummary
{
    internal const int MaxLength = 160;

    internal static string Create(string command, JsonElement? args)
    {
        var details = new List<string>();
        string? dimensions = null;
        if (args is { ValueKind: JsonValueKind.Object } arguments)
        {
            AddRange(details, arguments, "sheetName", "rangeAddress", null);
            AddRange(details, arguments, "sourceSheet", "sourceRange", "source");
            AddRange(details, arguments, "targetSheet", "targetRange", "target");
            foreach (var name in new[] { "tableName", "queryName", "pivotTableName", "chartName", "oldName", "newName" })
            {
                if (ReadString(arguments, name) is { } value)
                {
                    details.Add($"{name}={value}");
                }
            }

            var arrayName = command.Equals("range.set-values", StringComparison.OrdinalIgnoreCase) ? "values"
                : command.Equals("range.set-formulas", StringComparison.OrdinalIgnoreCase) ? "formulas" : null;
            if (arrayName is not null)
            {
                dimensions = ReadDimensions(arguments, arrayName);
            }
        }

        var summary = details.Count == 0 ? command : $"{command}: {string.Join(", ", details)}";
        if (dimensions is not null)
        {
            summary += $" ({dimensions})";
        }
        return ToSingleLine(summary);
    }

    private static void AddRange(List<string> details, JsonElement args, string sheetKey, string rangeKey, string? label)
    {
        var sheet = ReadString(args, sheetKey);
        var range = ReadString(args, rangeKey);
        if (sheet is not null && range is not null)
        {
            var quotedSheet = sheet.Any(c => !char.IsLetterOrDigit(c) && c is not '_' and not '.')
                ? $"'{sheet.Replace("'", "''", StringComparison.Ordinal)}'" : sheet;
            var target = $"{quotedSheet}!{range}";
            details.Add(label is null ? target : $"{label}={target}");
        }
        else if (range is not null)
        {
            details.Add($"{rangeKey}={range}");
        }
        else if (sheet is not null)
        {
            details.Add($"{sheetKey}={sheet}");
        }
    }

    private static string? ReadString(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String)
        {
            return null;
        }
        try
        {
            return value.GetString() is { Length: > 0 } text ? text : null;
        }
        catch (InvalidOperationException)
        {
            // Parsed JSON can contain unmatched surrogate escapes. Omit only the
            // unreadable display detail, preserving command validation and errors.
            return null;
        }
    }

    private static string? ReadDimensions(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var array) || array.ValueKind != JsonValueKind.Array || array.GetArrayLength() == 0)
        {
            return null;
        }
        int? columns = null;
        foreach (var row in array.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() == 0 ||
                (columns.HasValue && columns.Value != row.GetArrayLength()))
            {
                return null;
            }
            columns = row.GetArrayLength();
        }
        var rows = array.GetArrayLength();
        return FormattableString.Invariant($"{rows} {(rows == 1 ? "row" : "rows")} x {columns} {(columns == 1 ? "column" : "columns")}");
    }

    private static string ToSingleLine(string text)
    {
        var result = new StringBuilder(MaxLength + 1);
        foreach (var rune in text.EnumerateRunes())
        {
            if (Rune.IsControl(rune) || Rune.IsWhiteSpace(rune) || Rune.GetUnicodeCategory(rune) == UnicodeCategory.Format)
            {
                result.Append(' ');
            }
            else
            {
                result.Append(rune.ToString());
            }
            if (result.Length > MaxLength)
            {
                var length = MaxLength - 3;
                if (char.IsHighSurrogate(result[length - 1]))
                {
                    length--;
                }
                result.Length = length;
                return result.Append("...").ToString();
            }
        }
        return result.ToString();
    }
}
