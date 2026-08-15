using System.Xml;
using System.Xml.Linq;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.XmlMap;

/// <summary>
/// Implements deterministic XML map operations through the typed Excel PIA.
/// </summary>
public sealed class XmlMapCommands : IXmlMapCommands
{
    /// <inheritdoc />
    public XmlMapListResult List(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMaps? maps = null;
            try
            {
                maps = ctx.Book.XmlMaps;
                var result = new XmlMapListResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };

                for (var index = 1; index <= maps.Count; index++)
                {
                    Excel.XmlMap? map = null;
                    try
                    {
                        map = maps.Item[index];
                        result.Maps.Add(new XmlMapInfo
                        {
                            Name = map.Name,
                            RootElementName = map.RootElementName,
                            IsExportable = map.IsExportable
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref map);
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref maps);
            }
        });
    }

    /// <inheritdoc />
    public XmlMapAddResult Add(
        IExcelBatch batch,
        string schema,
        string? rootElementName = null,
        string? mapName = null)
    {
        ValidateSchema(schema);

        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMaps? maps = null;
            Excel.XmlMap? map = null;
            try
            {
                maps = ctx.Book.XmlMaps;
                map = maps.Add(
                    schema,
                    string.IsNullOrWhiteSpace(rootElementName) ? Type.Missing : rootElementName);

                if (!string.IsNullOrWhiteSpace(mapName))
                {
                    map.Name = mapName;
                }

                return new XmlMapAddResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    MapName = map.Name,
                    RootElementName = map.RootElementName
                };
            }
            finally
            {
                ComUtilities.Release(ref map);
                ComUtilities.Release(ref maps);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult MapRange(
        IExcelBatch batch,
        string mapName,
        string sheetName,
        string rangeAddress,
        string xpath,
        string? selectionNamespace = null,
        bool repeating = false)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMap? map = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.XPath? rangeXPath = null;
            try
            {
                map = FindXmlMap(ctx.Book, mapName)
                    ?? throw new InvalidOperationException($"XML map '{mapName}' not found.");
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                range = sheet.Range[rangeAddress];
                rangeXPath = range.XPath;
                rangeXPath.SetValue(
                    map,
                    xpath,
                    string.IsNullOrWhiteSpace(selectionNamespace) ? Type.Missing : selectionNamespace,
                    repeating);

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "map-range",
                    Message = $"Mapped '{sheetName}!{rangeAddress}' to '{xpath}' in XML map '{mapName}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref rangeXPath);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref map);
            }
        });
    }

    /// <inheritdoc />
    public XmlMapImportResult ImportXml(
        IExcelBatch batch,
        string xmlData,
        string? mapName = null,
        string? sheetName = null,
        string startCell = "A1",
        bool overwrite = true)
    {
        ValidateXmlData(xmlData);

        if (string.IsNullOrWhiteSpace(mapName) && string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException(
                "sheetName is required when mapName is not provided.",
                nameof(sheetName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMap? map = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? destination = null;
            try
            {
                Excel.XlXmlImportResult importStatus;
                if (!string.IsNullOrWhiteSpace(mapName))
                {
                    map = FindXmlMap(ctx.Book, mapName)
                        ?? throw new InvalidOperationException($"XML map '{mapName}' not found.");
                    importStatus = map.ImportXml(xmlData, overwrite);
                }
                else
                {
                    sheet = ComUtilities.FindSheet(ctx.Book, sheetName!)
                        ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                    destination = sheet.Range[startCell];
                    importStatus = ctx.Book.XmlImportXml(xmlData, out map, overwrite, destination);
                }

                EnsureImportSucceeded(importStatus);

                return new XmlMapImportResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    MapName = map.Name,
                    ImportStatus = importStatus.ToString(),
                    SheetName = string.IsNullOrWhiteSpace(mapName) ? sheetName : null,
                    StartCell = string.IsNullOrWhiteSpace(mapName) ? startCell : null
                };
            }
            finally
            {
                ComUtilities.Release(ref destination);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref map);
            }
        });
    }

    /// <inheritdoc />
    public XmlMapExportResult ExportXml(IExcelBatch batch, string mapName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMap? map = null;
            try
            {
                map = FindXmlMap(ctx.Book, mapName)
                    ?? throw new InvalidOperationException($"XML map '{mapName}' not found.");

                var exportStatus = map.ExportXml(out var xmlData);
                if (exportStatus != Excel.XlXmlExportResult.xlXmlExportSuccess)
                {
                    throw new InvalidOperationException(
                        $"Excel could not export XML map '{mapName}': {exportStatus}.");
                }

                return new XmlMapExportResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    MapName = map.Name,
                    XmlData = xmlData,
                    ExportStatus = exportStatus.ToString()
                };
            }
            finally
            {
                ComUtilities.Release(ref map);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string mapName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.XmlMap? map = null;
            try
            {
                map = FindXmlMap(ctx.Book, mapName)
                    ?? throw new InvalidOperationException($"XML map '{mapName}' not found.");
                map.Delete();

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "delete",
                    Message = $"Deleted XML map '{mapName}'."
                };
            }
            finally
            {
                ComUtilities.Release(ref map);
            }
        });
    }

    private static Excel.XmlMap? FindXmlMap(Excel.Workbook workbook, string mapName)
    {
        Excel.XmlMaps? maps = null;
        try
        {
            maps = workbook.XmlMaps;
            for (var index = 1; index <= maps.Count; index++)
            {
                Excel.XmlMap? map = null;
                try
                {
                    map = maps.Item[index];
                    var isMatch = string.Equals(map.Name, mapName, StringComparison.OrdinalIgnoreCase);
                    if (isMatch)
                    {
                        var result = map;
                        map = null;
                        return result;
                    }

                    ComUtilities.Release(ref map);
                }
                finally
                {
                    ComUtilities.Release(ref map);
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref maps);
        }
    }

    private static void ValidateSchema(string schema)
    {
        var document = ValidateXml(schema, nameof(schema));
        var hasExternalDependency = document
            .Descendants()
            .Any(element =>
                element.Name.NamespaceName == "http://www.w3.org/2001/XMLSchema" &&
                element.Name.LocalName is "import" or "include" or "redefine");

        if (hasExternalDependency)
        {
            throw new ArgumentException(
                "XSD import/include/redefine is not supported because it can load an external schema.",
                nameof(schema));
        }
    }

    private static void ValidateXmlData(string xmlData)
    {
        const string xmlSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        var document = ValidateXml(xmlData, nameof(xmlData));
        var hasExternalSchemaLocation = document
            .Descendants()
            .Attributes()
            .Any(attribute =>
                attribute.Name.NamespaceName == xmlSchemaInstanceNamespace &&
                attribute.Name.LocalName is "schemaLocation" or "noNamespaceSchemaLocation");

        if (hasExternalSchemaLocation)
        {
            throw new ArgumentException(
                "XML schema location attributes (xsi:schemaLocation and xsi:noNamespaceSchemaLocation) are not supported because Excel can load an external schema.",
                nameof(xmlData));
        }
    }

    private static XDocument ValidateXml(string xml, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(xml))
        {
            throw new ArgumentException("XML content cannot be empty.", parameterName);
        }

        using var stringReader = new StringReader(xml);
        using var xmlReader = XmlReader.Create(
            stringReader,
            new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
        return XDocument.Load(xmlReader, LoadOptions.None);
    }

    private static void EnsureImportSucceeded(Excel.XlXmlImportResult importStatus)
    {
        if (importStatus != Excel.XlXmlImportResult.xlXmlImportSuccess)
        {
            throw new InvalidOperationException($"Excel XML import failed: {importStatus}.");
        }
    }
}
