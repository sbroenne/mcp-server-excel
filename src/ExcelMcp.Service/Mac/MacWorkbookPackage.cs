using System.IO.Compression;
using System.Text;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacWorkbookPackage
{
    public static void Create(string filePath, bool macroEnabled)
    {
        var temporaryPath = Path.Combine(
            Path.GetDirectoryName(filePath) ?? throw new ArgumentException("Workbook path has no parent directory.", nameof(filePath)),
            $".excelmcp-{Guid.NewGuid():N}.tmp");
        try
        {
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Create))
            {
                Write(archive, "[Content_Types].xml",
                    $"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{(macroEnabled ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>""");
                Write(archive, "_rels/.rels",
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""");
                Write(archive, "xl/workbook.xml",
                    """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
                Write(archive, "xl/_rels/workbook.xml.rels",
                    """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""");
                Write(archive, "xl/worksheets/sheet1.xml",
                    """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>""");
            }

            File.Move(temporaryPath, filePath);
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }

    private static void Write(ZipArchive archive, string path, string content)
    {
        using var writer = new StreamWriter(
            archive.CreateEntry(path, CompressionLevel.Optimal).Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }
}
