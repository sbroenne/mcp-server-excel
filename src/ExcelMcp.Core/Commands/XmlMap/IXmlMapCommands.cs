using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.XmlMap;

/// <summary>
/// Manage workbook XML maps and exchange XML data without interactive dialogs.
///
/// SECURITY: Schemas and XML data are parsed from supplied content with DTD processing
/// disabled. XSD import/include/redefine dependencies and XML schema-location attributes
/// are rejected to prevent implicit network or file access.
///
/// IMPORT MODES: Provide mapName to import into existing mapped cells. Omit mapName and
/// provide sheetName plus startCell to let Excel create a map and XML table at that destination.
/// </summary>
[ServiceCategory("xmlmap", "XmlMap")]
[McpTool("xmlmap", Title = "XML Map Operations", Destructive = true, Category = "data",
    Description = "Manage workbook XML maps and exchange XML data without interactive dialogs. Schemas and XML data are parsed from supplied content with DTD processing disabled; external XSD dependencies and XML schema-location attributes are rejected. IMPORT MODES: provide map_name to import into existing mapped cells, or omit map_name and provide sheet_name plus start_cell to let Excel create a map and XML table.")]
public interface IXmlMapCommands
{
    /// <summary>
    /// Lists all XML maps in the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    [ServiceAction("list")]
    XmlMapListResult List(IExcelBatch batch);

    /// <summary>
    /// Adds an XML map from an inline XSD schema or schema file content.
    /// External XSD import/include dependencies are rejected.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="schema">XSD schema content. Public callers must supply either inline schema or a readable schemaFile, not both.</param>
    /// <param name="rootElementName">Optional root element when the schema has multiple roots</param>
    /// <param name="mapName">Optional name to assign to the created map</param>
    [ServiceAction("add")]
    XmlMapAddResult Add(
        IExcelBatch batch,
        [RequiredParameter, FileOrValue] string schema,
        string? rootElementName = null,
        string? mapName = null);

    /// <summary>
    /// Maps a worksheet range to an XPath in an existing XML map.
    /// Set repeating=true to create a repeating XML list mapping.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mapName">XML map name</param>
    /// <param name="sheetName">Worksheet containing the target range</param>
    /// <param name="rangeAddress">Target cell or single-column range</param>
    /// <param name="xpath">XPath to map</param>
    /// <param name="selectionNamespace">Optional namespace declarations used by prefixed XPath expressions</param>
    /// <param name="repeating">Whether to create a repeating XML list mapping</param>
    [ServiceAction("map-range")]
    OperationResult MapRange(
        IExcelBatch batch,
        [RequiredParameter] string mapName,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter] string xpath,
        string? selectionNamespace = null,
        bool repeating = false);

    /// <summary>
    /// Imports XML content without opening an Excel dialog.
    /// Provide mapName to update existing mapped cells. Otherwise provide sheetName and
    /// optionally startCell so Excel creates an XML map and table at the destination.
    /// XML schema-location attributes are rejected to prevent implicit external access.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="xmlData">XML data. Public callers must supply either inline xmlData or a readable xmlDataFile, not both.</param>
    /// <param name="mapName">Existing XML map name; omit for automatic mapping</param>
    /// <param name="sheetName">Destination worksheet for automatic mapping</param>
    /// <param name="startCell">Top-left destination cell for automatic mapping</param>
    /// <param name="overwrite">Whether imported XML may overwrite mapped cells</param>
    [ServiceAction("import-xml")]
    XmlMapImportResult ImportXml(
        IExcelBatch batch,
        [RequiredParameter, FileOrValue] string xmlData,
        string? mapName = null,
        string? sheetName = null,
        string startCell = "A1",
        bool overwrite = true);

    /// <summary>
    /// Exports mapped cell values to an XML string.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mapName">XML map name</param>
    [ServiceAction("export-xml")]
    XmlMapExportResult ExportXml(IExcelBatch batch, [RequiredParameter] string mapName);

    /// <summary>
    /// Deletes an XML map. Existing cell data remains in the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="mapName">XML map name</param>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, [RequiredParameter] string mapName);
}
