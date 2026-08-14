using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    private const int MsoPropertyTypeString = 4;

    /// <inheritdoc />
    public WorkbookInfoResult GetInfo(IExcelBatch batch)
    {
        return batch.Execute((context, _) =>
        {
            var fileFormat = context.Book.FileFormat;
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
                ReadOnly = context.Book.ReadOnly,
                HasPassword = context.Book.HasPassword,
                WriteReserved = context.Book.WriteReserved
            };
        });
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
            dynamic? property = null;
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
                ComUtilities.Release(ref property!);
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
            dynamic? properties = null;
            dynamic? property = null;
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

                    property = properties.Add(propertyName, false, MsoPropertyTypeString, value);
                }
                else
                {
                    property.Value = value;
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
                ComUtilities.Release(ref property!);
                ComUtilities.Release(ref properties!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteDocumentProperty(IExcelBatch batch, string propertyName)
    {
        ValidatePropertyName(propertyName);
        return batch.Execute((context, _) =>
        {
            dynamic? property = null;
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
                ComUtilities.Release(ref property!);
            }
        });
    }

    private static void AddProperties(
        Excel.Workbook workbook,
        DocumentPropertyScope scope,
        List<DocumentPropertyInfo> target)
    {
        dynamic? properties = null;
        try
        {
            properties = GetPropertyCollection(workbook, scope);
            var count = Convert.ToInt32(properties.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic? property = null;
                try
                {
                    property = properties.Item(index);
                    target.Add(ReadProperty(property, scope));
                }
                finally
                {
                    ComUtilities.Release(ref property!);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref properties!);
        }
    }

    private static dynamic GetPropertyCollection(Excel.Workbook workbook, DocumentPropertyScope scope)
    {
        // PIA gap: Built-in and custom document properties are Office.DocumentProperties from office.dll.
        return scope == DocumentPropertyScope.BuiltIn
            ? workbook.BuiltinDocumentProperties
            : workbook.CustomDocumentProperties;
    }

    private static dynamic GetRequiredProperty(
        Excel.Workbook workbook,
        string propertyName,
        DocumentPropertyScope scope)
    {
        dynamic? properties = null;
        try
        {
            properties = GetPropertyCollection(workbook, scope);
            return TryGetProperty(properties, propertyName)
                ?? throw new InvalidOperationException(
                    $"{GetScopeName(scope)} document property '{propertyName}' was not found.");
        }
        finally
        {
            ComUtilities.Release(ref properties!);
        }
    }

    private static dynamic? TryGetProperty(dynamic properties, string propertyName)
    {
        var count = Convert.ToInt32(properties.Count, CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            dynamic? property = null;
            try
            {
                property = properties.Item(index);
                var currentName = Convert.ToString(property.Name, CultureInfo.InvariantCulture);
                if (string.Equals(currentName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    var result = property;
                    property = null;
                    return result;
                }
            }
            finally
            {
                ComUtilities.Release(ref property!);
            }
        }

        return null;
    }

    private static DocumentPropertyInfo ReadProperty(dynamic property, DocumentPropertyScope scope)
    {
        object? value;
        try
        {
            value = property.Value;
        }
        catch (COMException)
        {
            value = null;
        }

        var typeCode = Convert.ToInt32(property.Type, CultureInfo.InvariantCulture);
        return new DocumentPropertyInfo
        {
            Name = Convert.ToString(property.Name, CultureInfo.InvariantCulture) ?? string.Empty,
            Value = Convert.ToString(value, CultureInfo.InvariantCulture),
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
