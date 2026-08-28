using System.Globalization;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

internal static class TypedCellValueParser
{
    internal const string DefaultDateNumberFormat = "yyyy-mm-dd";
    internal const string DefaultDateTimeNumberFormat = "yyyy-mm-dd hh:mm:ss";

    private static readonly string[] s_dateTimeFormats =
    [
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF"
    ];

    private static readonly string[] s_dateTimeOffsetFormats =
    [
        "yyyy-MM-dd'T'HH:mm:sszzz",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz"
    ];

    private static readonly string[] s_utcDateTimeOffsetFormats =
    [
        "yyyy-MM-dd'T'HH:mm:ss'Z'",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'"
    ];

    internal static PreparedCellValue Parse(object? input, int rowIndex, int columnIndex)
    {
        if (input is TypedCellValue typedValue)
        {
            return ParseTypedValue(typedValue, rowIndex, columnIndex);
        }

        if (input is JsonElement { ValueKind: JsonValueKind.Object } jsonObject)
        {
            return ParseTypedValue(ParseJsonObject(jsonObject, rowIndex, columnIndex), rowIndex, columnIndex);
        }

        return PreparedCellValue.FromPrimitive(input);
    }

    internal static double ToExcelSerial(
        DateTime value,
        bool use1904DateSystem,
        int rowIndex,
        int columnIndex)
    {
        var minimum = use1904DateSystem
            ? new DateTime(1904, 1, 1)
            : new DateTime(1900, 1, 1);

        if (value < minimum)
        {
            var dateSystem = use1904DateSystem ? "1904" : "1900";
            throw CreateError(
                rowIndex,
                columnIndex,
                $"value must be on or after {minimum:yyyy-MM-dd} for the workbook's {dateSystem} date system.");
        }

        var wholeDays = (value.Date.Ticks - minimum.Ticks) / TimeSpan.TicksPerDay;
        var timeOfDay = value.TimeOfDay.Ticks / (double)TimeSpan.TicksPerDay;
        var serial = wholeDays + timeOfDay;

        // Excel's 1900 system includes the fictitious 1900-02-29 at serial 60.
        if (!use1904DateSystem)
        {
            serial += value < new DateTime(1900, 3, 1) ? 1d : 2d;
        }

        return serial;
    }

    private static TypedCellValue ParseJsonObject(
        JsonElement jsonObject,
        int rowIndex,
        int columnIndex)
    {
        if (!jsonObject.TryGetProperty("type", out var typeElement)
            || typeElement.ValueKind != JsonValueKind.String)
        {
            throw CreateError(rowIndex, columnIndex, "type is required as a string.");
        }

        if (!jsonObject.TryGetProperty("value", out var valueElement)
            || valueElement.ValueKind != JsonValueKind.String)
        {
            throw CreateError(rowIndex, columnIndex, "value is required as a string.");
        }

        string? numberFormat = null;
        if (jsonObject.TryGetProperty("numberFormat", out var numberFormatElement))
        {
            if (numberFormatElement.ValueKind != JsonValueKind.String)
            {
                throw CreateError(rowIndex, columnIndex, "numberFormat must be a string when provided.");
            }

            numberFormat = numberFormatElement.GetString();
        }

        return new TypedCellValue
        {
            Type = ParseType(typeElement.GetString(), rowIndex, columnIndex),
            Value = valueElement.GetString(),
            NumberFormat = numberFormat
        };
    }

    private static TypedCellValueType ParseType(
        string? type,
        int rowIndex,
        int columnIndex)
    {
        return type switch
        {
            "date" => TypedCellValueType.Date,
            "datetime" => TypedCellValueType.DateTime,
            "datetime-offset" => TypedCellValueType.DateTimeOffset,
            _ => throw CreateError(
                rowIndex,
                columnIndex,
                $"type '{type}' is invalid. Valid types: date, datetime, datetime-offset.")
        };
    }

    private static PreparedCellValue ParseTypedValue(
        TypedCellValue typedValue,
        int rowIndex,
        int columnIndex)
    {
        if (!typedValue.Type.HasValue)
        {
            throw CreateError(rowIndex, columnIndex, "type is required.");
        }

        if (string.IsNullOrWhiteSpace(typedValue.Value))
        {
            throw CreateError(rowIndex, columnIndex, "value is required and must not be empty.");
        }

        if (typedValue.NumberFormat != null && string.IsNullOrWhiteSpace(typedValue.NumberFormat))
        {
            throw CreateError(rowIndex, columnIndex, "numberFormat must not be empty.");
        }

        var parsedDateTime = typedValue.Type.Value switch
        {
            TypedCellValueType.Date => ParseDate(typedValue.Value, rowIndex, columnIndex),
            TypedCellValueType.DateTime => ParseDateTime(typedValue.Value, rowIndex, columnIndex),
            TypedCellValueType.DateTimeOffset => ParseDateTimeOffset(typedValue.Value, rowIndex, columnIndex),
            _ => throw CreateError(
                rowIndex,
                columnIndex,
                $"type '{typedValue.Type}' is invalid. Valid types: date, datetime, datetime-offset.")
        };

        var defaultNumberFormat = typedValue.Type == TypedCellValueType.Date
            ? DefaultDateNumberFormat
            : DefaultDateTimeNumberFormat;

        return PreparedCellValue.FromTypedDate(
            parsedDateTime,
            typedValue.NumberFormat ?? defaultNumberFormat,
            rowIndex,
            columnIndex);
    }

    private static DateTime ParseDate(string value, int rowIndex, int columnIndex)
    {
        if (!DateOnly.TryParseExact(
                value,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date))
        {
            throw CreateError(
                rowIndex,
                columnIndex,
                $"date value '{value}' is invalid. Expected ISO format yyyy-MM-dd.");
        }

        return date.ToDateTime(TimeOnly.MinValue, DateTimeKind.Unspecified);
    }

    private static DateTime ParseDateTime(string value, int rowIndex, int columnIndex)
    {
        if (!DateTime.TryParseExact(
                value,
                s_dateTimeFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var dateTime))
        {
            throw CreateError(
                rowIndex,
                columnIndex,
                $"datetime value '{value}' is invalid. Expected yyyy-MM-ddTHH:mm:ss with optional fractional seconds and without a timezone.");
        }

        return DateTime.SpecifyKind(dateTime, DateTimeKind.Unspecified);
    }

    private static DateTime ParseDateTimeOffset(string value, int rowIndex, int columnIndex)
    {
        DateTimeOffset parsed;
        var parsedSuccessfully = value.EndsWith('Z')
            ? DateTimeOffset.TryParseExact(
                value,
                s_utcDateTimeOffsetFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out parsed)
            : DateTimeOffset.TryParseExact(
                value,
                s_dateTimeOffsetFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out parsed);

        if (!parsedSuccessfully)
        {
            throw CreateError(
                rowIndex,
                columnIndex,
                $"datetime-offset value '{value}' is invalid. Expected yyyy-MM-ddTHH:mm:ss with Z or an offset such as +02:00; fractional seconds are optional.");
        }

        return DateTime.SpecifyKind(parsed.UtcDateTime, DateTimeKind.Unspecified);
    }

    private static ArgumentException CreateError(
        int rowIndex,
        int columnIndex,
        string message)
    {
        return new ArgumentException(
            $"Invalid typed value at row {rowIndex}, column {columnIndex}: {message}");
    }
}

internal readonly record struct PreparedCellValue(
    object? PrimitiveValue,
    DateTime DateTimeValue,
    string? NumberFormat,
    bool IsTypedDate,
    int RowIndex,
    int ColumnIndex)
{
    internal static PreparedCellValue FromPrimitive(object? value) =>
        new(value, default, null, false, 0, 0);

    internal static PreparedCellValue FromTypedDate(
        DateTime value,
        string numberFormat,
        int rowIndex,
        int columnIndex) =>
        new(null, value, numberFormat, true, rowIndex, columnIndex);
}
