using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    private const int BuiltInPropertyValueFailure = unchecked((int)0x80004005);

    /// <inheritdoc />
    public WorkbookInfoResult GetInfo(IExcelBatch batch) =>
        batch.Execute((context, _) => ReadWorkbookInfo(context));

    private static WorkbookInfoResult ReadWorkbookInfo(ExcelContext context)
    {
        var fileFormat = context.Book.FileFormat;
        var autoSaveOn = ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn);
        return new WorkbookInfoResult
        {
            Success = true,
            FilePath = context.Book.FullName,
            Name = context.Book.Name,
            FullName = context.Book.FullName,
            DirectoryPath = context.Book.Path,
            Format = GetFormatName(fileFormat),
            FormatCode = Convert.ToInt32(fileFormat, CultureInfo.InvariantCulture),
            Saved = context.Book.Saved,
            AutoSaveOn = autoSaveOn,
            Warning = ExcelCapabilities.AutoSaveWarning(autoSaveOn),
            ReadOnly = context.Book.ReadOnly,
            HasPassword = context.Book.HasPassword,
            WriteReserved = context.Book.WriteReserved
        };
    }

    /// <inheritdoc />
    public WorkbookInfoResult SetAutoSave(IExcelBatch batch, bool enabled)
    {
        return batch.Execute((context, _) =>
        {
            if (context.Book.ReadOnly)
                throw new InvalidOperationException("Cannot change AutoSave on a read-only workbook.");
            ApplyAutoSaveSetting(
                () => ExcelCapabilities.ReadAutoSaveStatus(() => context.Book.AutoSaveOn),
                value => context.Book.AutoSaveOn = value,
                enabled);
            return ReadWorkbookInfo(context);
        });
    }

    internal static void ApplyAutoSaveSetting(
        Func<bool?> readAutoSave,
        Action<bool> writeAutoSave,
        bool enabled)
    {
        if (readAutoSave() is null)
            throw new NotSupportedException("AutoSave is unavailable in this version of Excel.");
        writeAutoSave(enabled);
        if (readAutoSave() != enabled)
            throw new InvalidOperationException("Excel did not apply the requested AutoSave setting. Inspect the workbook before retrying.");
    }

    /// <inheritdoc />
    public DocumentPropertyListResult ListDocumentProperties(
        IExcelBatch batch,
        bool includeBuiltIn = true,
        bool includeCustom = true)
    {
        return batch.Execute((context, _) =>
        {
            var result = new DocumentPropertyListResult { Success = true };
            if (includeBuiltIn)
            {
                AddProperties(context.Book, DocumentPropertyScope.BuiltIn, result.Properties);
            }

            if (includeCustom)
            {
                AddProperties(context.Book, DocumentPropertyScope.Custom, result.Properties);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DocumentPropertyResult GetDocumentProperty(
        IExcelBatch batch,
        string propertyName,
        DocumentPropertyScope scope = DocumentPropertyScope.Custom)
    {
        ValidatePropertyName(propertyName);
        return batch.Execute((context, _) =>
        {
            ExcelDocumentProperties.Property? property = null;
            try
            {
                property = GetRequiredProperty(context.Book, propertyName, scope);
                return new DocumentPropertyResult
                {
                    Success = true,
                    Property = ReadProperty(property, scope)
                };
            }
            finally
            {
                property?.Dispose();
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetDocumentProperty(
        IExcelBatch batch,
        string propertyName,
        string value,
        DocumentPropertyScope scope = DocumentPropertyScope.Custom)
    {
        ValidatePropertyName(propertyName);
        return batch.Execute((context, _) =>
        {
            ExcelDocumentProperties? properties = null;
            ExcelDocumentProperties.Property? property = null;
            try
            {
                properties = GetPropertyCollection(context.Book, scope);
                property = TryGetProperty(properties, propertyName);
                if (property == null)
                {
                    if (scope == DocumentPropertyScope.BuiltIn)
                    {
                        throw new InvalidOperationException($"Built-in document property '{propertyName}' was not found.");
                    }

                    property = properties.Add(propertyName, Office.MsoDocProperties.msoPropertyTypeString, value);
                }
                else
                {
                    property.SetValue(value);
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-document-property",
                    Message = $"{GetScopeName(scope)} document property '{propertyName}' was set"
                };
            }
            finally
            {
                property?.Dispose();
                properties?.Dispose();
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteDocumentProperty(IExcelBatch batch, string propertyName)
    {
        ValidatePropertyName(propertyName);
        return batch.Execute((context, _) =>
        {
            ExcelDocumentProperties.Property? property = null;
            try
            {
                property = GetRequiredProperty(context.Book, propertyName, DocumentPropertyScope.Custom);
                property.Delete();
                return new OperationResult
                {
                    Success = true,
                    Action = "delete-document-property",
                    Message = $"Custom document property '{propertyName}' was deleted"
                };
            }
            finally
            {
                property?.Dispose();
            }
        });
    }

    private static void AddProperties(
        Excel.Workbook workbook,
        DocumentPropertyScope scope,
        List<DocumentPropertyInfo> target)
    {
        ExcelDocumentProperties? properties = null;
        try
        {
            properties = GetPropertyCollection(workbook, scope);
            int count = properties.Count;
            for (var index = 1; index <= count; index++)
            {
                ExcelDocumentProperties.Property? property = null;
                try
                {
                    property = properties.Item(index);
                    target.Add(ReadProperty(property, scope));
                }
                finally
                {
                    property?.Dispose();
                }
            }
        }
        finally
        {
            properties?.Dispose();
        }
    }

    private static ExcelDocumentProperties GetPropertyCollection(Excel.Workbook workbook, DocumentPropertyScope scope)
    {
        return ExcelDocumentProperties.Open(workbook, scope == DocumentPropertyScope.BuiltIn);
    }

    private static ExcelDocumentProperties.Property GetRequiredProperty(
        Excel.Workbook workbook,
        string propertyName,
        DocumentPropertyScope scope)
    {
        ExcelDocumentProperties? properties = null;
        try
        {
            properties = GetPropertyCollection(workbook, scope);
            return TryGetProperty(properties, propertyName)
                ?? throw new InvalidOperationException(
                    $"{GetScopeName(scope)} document property '{propertyName}' was not found.");
        }
        finally
        {
            properties?.Dispose();
        }
    }

    private static ExcelDocumentProperties.Property? TryGetProperty(ExcelDocumentProperties properties, string propertyName)
    {
        int count = properties.Count;
        for (var index = 1; index <= count; index++)
        {
            ExcelDocumentProperties.Property? property = null;
            try
            {
                property = properties.Item(index);
                string currentName = property.Name;
                if (string.Equals(currentName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    var result = property;
                    property = null;
                    return result;
                }
            }
            finally
            {
                property?.Dispose();
            }
        }

        return null;
    }

    private static DocumentPropertyInfo ReadProperty(ExcelDocumentProperties.Property property, DocumentPropertyScope scope) =>
        ReadDocumentProperty(() => property.Name, () => (int)property.Type, () => property.Value, scope);

    internal static DocumentPropertyInfo ReadDocumentProperty(
        Func<string> readName, Func<int> readType, Func<object?> readValue, DocumentPropertyScope scope)
    {
        int typeCode = readType();
        var result = new DocumentPropertyInfo
        {
            Name = readName(),
            ValueType = typeCode switch
            {
                1 => "number",
                2 => "boolean",
                3 => "date",
                4 => "string",
                5 => "floating-point",
                _ => $"unknown-{typeCode}"
            },
            Scope = GetScopeName(scope)
        };
        object? value;
        try
        {
            value = readValue();
        }
        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex) ||
            (scope == DocumentPropertyScope.BuiltIn && ex.HResult == BuiltInPropertyValueFailure))
        {
            // Office documents errors for undefined built-in values but exposes no defined flag.
            // E_FAIL alone cannot distinguish that state from a failed Value getter.
            result.ValueAvailable = false;
            result.ValueError =
                $"Could not read document property '{result.Name}' Value (0x{ex.HResult:X8}): {ex.Message}. " +
                (scope == DocumentPropertyScope.BuiltIn
                    ? "The built-in value may be undefined or its getter may have failed; these states cannot be distinguished."
                    : "The Value member is unavailable.");
            return result;
        }
        result.Value = Convert.ToString(value, CultureInfo.InvariantCulture);
        return result;
    }

    private static string GetFormatName(Excel.XlFileFormat format)
    {
        return format switch
        {
            Excel.XlFileFormat.xlOpenXMLWorkbook => "xlsx",
            Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled => "xlsm",
            Excel.XlFileFormat.xlExcel12 => "xlsb",
            Excel.XlFileFormat.xlExcel8 => "xls",
            Excel.XlFileFormat.xlCSV => "csv",
            _ => format.ToString()
        };
    }

    private static string GetScopeName(DocumentPropertyScope scope) =>
        scope == DocumentPropertyScope.BuiltIn ? "built-in" : "custom";

    private static void ValidatePropertyName(string propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            throw new ArgumentException("Document property name cannot be empty.", nameof(propertyName));
        }
    }
}
