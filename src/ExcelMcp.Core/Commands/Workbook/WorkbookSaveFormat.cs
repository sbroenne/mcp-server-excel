namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// File formats supported by workbook Save As.
/// </summary>
public enum WorkbookSaveFormat
{
    /// <summary>Infer the format from the output file extension.</summary>
    Auto,

    /// <summary>Excel Open XML workbook (.xlsx).</summary>
    Xlsx,

    /// <summary>Excel macro-enabled Open XML workbook (.xlsm).</summary>
    Xlsm,

    /// <summary>Excel binary workbook (.xlsb).</summary>
    Xlsb,

    /// <summary>Excel 97-2003 workbook (.xls).</summary>
    Xls
}
