using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public enum PowerQueryFixtureKind
{
    ConnectionOnly,
    WorksheetLoaded
}

internal sealed record PowerQueryFixture(
    string WorkbookPath,
    string ProvenanceManifestPath);

internal sealed record PowerQueryFixtureAudit(
    IReadOnlyList<string> PackageParts,
    IReadOnlyList<string> Errors,
    string Sha256);

internal sealed class PowerQueryFixtureManifest
{
    public string Provenance { get; set; } = string.Empty;
    public string QueryName { get; set; } = string.Empty;
    public string LoadKind { get; set; } = string.Empty;
    public string MFormula { get; set; } = string.Empty;
    public string[] Specifications { get; set; } = [];
    public string[] AuthorshipAssertions { get; set; } = [];
    public string WorkbookSha256 { get; set; } = string.Empty;
}

internal static class PowerQueryFixtureFactory
{
    public const string QueryName = "LiteralRows";
    public const string WorksheetName = "Literal Output";
    public const string LiteralM =
        """
        let
            Source = #table(
                type table [Item = text, Amount = Int64.Type],
                {{"Alpha", 10}, {"Beta", 20}})
        in
            Source
        """;

    private const string DataMashupNamespace = "http://schemas.microsoft.com/DataMashup";
    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly DateTimeOffset PackageTimestamp =
        new(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private static readonly JsonSerializerOptions ManifestJsonOptions = new()
    {
        WriteIndented = true
    };
    private static readonly string[] RequiredContentTypeOverrides =
    [
        "/xl/workbook.xml",
        "/xl/worksheets/sheet1.xml",
        "/xl/connections.xml",
        "/customXml/itemProps1.xml"
    ];

    public static PowerQueryFixture Create(string directory, PowerQueryFixtureKind kind)
    {
        Directory.CreateDirectory(directory);
        var stem = kind == PowerQueryFixtureKind.ConnectionOnly
            ? "literal-query-connection-only"
            : "literal-query-worksheet-loaded";
        var workbookPath = Path.Combine(directory, $"{stem}.xlsx");
        var manifestPath = Path.Combine(directory, $"{stem}.provenance.json");
        var entries = BuildWorkbookParts(kind);

        using (var archive = ZipFile.Open(workbookPath, ZipArchiveMode.Create))
        {
            foreach (var entry in entries.OrderBy(item => item.Key, StringComparer.Ordinal))
            {
                WriteEntry(archive, entry.Key, entry.Value);
            }
        }

        var audit = Audit(workbookPath, kind);
        if (audit.Errors.Count > 0)
        {
            throw new InvalidDataException(
                $"Generated Power Query fixture failed package audit: {string.Join("; ", audit.Errors)}");
        }

        var manifest = new PowerQueryFixtureManifest
        {
            Provenance = "repository-authored",
            QueryName = QueryName,
            LoadKind = kind.ToString(),
            MFormula = LiteralM,
            Specifications =
            [
                "MS-QDEFF",
                "ECMA-376",
                "ISO 29500"
            ],
            AuthorshipAssertions =
            [
                "The workbook package, XML, identifiers, formula, values, and expected output are generated from repository source.",
                "No external workbook, research fixture, customer workbook, or copied package bytes are read by this factory."
            ],
            WorkbookSha256 = audit.Sha256
        };
        File.WriteAllText(
            manifestPath,
            JsonSerializer.Serialize(manifest, ManifestJsonOptions),
            new UTF8Encoding(false));
        return new PowerQueryFixture(workbookPath, manifestPath);
    }

    public static PowerQueryFixtureAudit Audit(
        string workbookPath,
        PowerQueryFixtureKind kind)
    {
        var errors = new List<string>();
        using var archive = ZipFile.OpenRead(workbookPath);
        var parts = archive.Entries.Select(entry => entry.FullName)
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToArray();
        var required = new List<string>
        {
            "[Content_Types].xml",
            "_rels/.rels",
            "customXml/item1.xml",
            "customXml/itemProps1.xml",
            "customXml/_rels/item1.xml.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/worksheets/sheet1.xml",
            "xl/connections.xml",
            "xl/styles.xml",
            "xl/theme/theme1.xml"
        };
        if (kind == PowerQueryFixtureKind.WorksheetLoaded)
        {
            required.AddRange(
            [
                "xl/worksheets/_rels/sheet1.xml.rels",
                "xl/tables/table1.xml",
                "xl/tables/_rels/table1.xml.rels",
                "xl/queryTables/queryTable1.xml"
            ]);
        }

        foreach (var requiredPart in required.Where(part => !parts.Contains(part, StringComparer.Ordinal)))
        {
            errors.Add($"Missing required package part '{requiredPart}'.");
        }

        AuditContentTypes(archive, kind, errors);
        AuditTheme(archive, errors);
        AuditRelationships(archive, kind, errors);
        AuditDataMashup(archive, errors);
        AuditLoadGraph(archive, kind, errors);

        var hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(workbookPath))).ToLowerInvariant();
        return new PowerQueryFixtureAudit(parts, errors, hash);
    }

    private static Dictionary<string, byte[]> BuildWorkbookParts(PowerQueryFixtureKind kind)
    {
        var isLoaded = kind == PowerQueryFixtureKind.WorksheetLoaded;
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["[Content_Types].xml"] = XmlBytes(ContentTypes(isLoaded)),
            ["_rels/.rels"] = XmlBytes(
                $"""
                <Relationships xmlns="{RelationshipsNamespace}">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """),
            ["xl/workbook.xml"] = XmlBytes(
                $"""
                <workbook xmlns="{SpreadsheetNamespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>
                  <sheets><sheet name="{WorksheetName}" sheetId="1" r:id="rId1"/></sheets>
                  <calcPr calcId="191029" fullCalcOnLoad="1"/>
                </workbook>
                """),
            ["xl/_rels/workbook.xml.rels"] = XmlBytes(
                $"""
                <Relationships xmlns="{RelationshipsNamespace}">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
                  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/>
                  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>
                </Relationships>
                """),
            ["xl/worksheets/sheet1.xml"] = XmlBytes(Worksheet(isLoaded)),
            ["xl/connections.xml"] = XmlBytes(Connections()),
            ["xl/styles.xml"] = XmlBytes(Styles()),
            ["xl/theme/theme1.xml"] = XmlBytes(Theme()),
            ["customXml/item1.xml"] = XmlBytes(
                $"""<DataMashup xmlns="{DataMashupNamespace}">{Convert.ToBase64String(BuildDataMashup(isLoaded))}</DataMashup>"""),
            ["customXml/itemProps1.xml"] = XmlBytes(
                """
                <ds:datastoreItem ds:itemID="{41C5B18A-71F1-4B72-A973-57CBDF77D40B}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
                  <ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/></ds:schemaRefs>
                </ds:datastoreItem>
                """),
            ["customXml/_rels/item1.xml.rels"] = XmlBytes(
                $"""
                <Relationships xmlns="{RelationshipsNamespace}">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>
                </Relationships>
                """)
        };

        if (isLoaded)
        {
            parts["xl/worksheets/_rels/sheet1.xml.rels"] = XmlBytes(
                $"""
                <Relationships xmlns="{RelationshipsNamespace}">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
                </Relationships>
                """);
            parts["xl/tables/table1.xml"] = XmlBytes(
                $"""
                <table xmlns="{SpreadsheetNamespace}" id="1" name="LiteralRowsTable" displayName="LiteralRowsTable" ref="A1:B3" tableType="queryTable" totalsRowShown="0">
                  <autoFilter ref="A1:B3"/>
                  <tableColumns count="2">
                    <tableColumn id="1" name="Item" queryTableFieldId="1"/>
                    <tableColumn id="2" name="Amount" queryTableFieldId="2"/>
                  </tableColumns>
                  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
                </table>
                """);
            parts["xl/tables/_rels/table1.xml.rels"] = XmlBytes(
                $"""
                <Relationships xmlns="{RelationshipsNamespace}">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/>
                </Relationships>
                """);
            parts["xl/queryTables/queryTable1.xml"] = XmlBytes(
                $"""
                <queryTable xmlns="{SpreadsheetNamespace}" name="LiteralRowsTable" connectionId="1" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1">
                  <queryTableRefresh nextId="3" minimumVersion="0">
                    <queryTableFields count="2">
                      <queryTableField id="1" name="Item" tableColumnId="1"/>
                      <queryTableField id="2" name="Amount" tableColumnId="2"/>
                    </queryTableFields>
                  </queryTableRefresh>
                </queryTable>
                """);
        }

        return parts;
    }

    private static string ContentTypes(bool loaded)
    {
        var loadOverrides = loaded
            ? """
                <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
                <Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>
              """
            : string.Empty;
        return
            $"""
             <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
               <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
               <Default Extension="xml" ContentType="application/xml"/>
               <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
               <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
               <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
               <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
               <Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>
               <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
             {loadOverrides}
             </Types>
             """;
    }

    private static string Worksheet(bool loaded) =>
        loaded
            ?
            $"""
             <worksheet xmlns="{SpreadsheetNamespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
               <dimension ref="A1:B3"/>
               <sheetViews><sheetView workbookViewId="0"/></sheetViews>
               <sheetFormatPr defaultRowHeight="15"/>
               <sheetData>
                 <row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row>
                 <row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>10</v></c></row>
                 <row r="3"><c r="A3" t="inlineStr"><is><t>Beta</t></is></c><c r="B3"><v>20</v></c></row>
               </sheetData>
               <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
               <tableParts count="1"><tablePart r:id="rId1"/></tableParts>
             </worksheet>
             """
            :
            $"""
             <worksheet xmlns="{SpreadsheetNamespace}">
               <sheetViews><sheetView workbookViewId="0"/></sheetViews>
               <sheetFormatPr defaultRowHeight="15"/>
               <sheetData/>
               <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
             </worksheet>
             """;

    private static string Connections() =>
        $"""
         <connections xmlns="{SpreadsheetNamespace}">
           <connection id="1" name="Query - {QueryName}" description="Repository-authored literal Power Query fixture" type="5" refreshedVersion="7" background="0" saveData="1">
             <dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={QueryName};Extended Properties=&quot;&quot;" command="SELECT * FROM [{QueryName}]" commandType="2"/>
           </connection>
         </connections>
         """;

    private static byte[] BuildDataMashup(bool loaded)
    {
        var packageParts = CreateZip(
            ("[Content_Types].xml", XmlBytes(
                """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="xml" ContentType="text/xml"/>
                  <Default Extension="m" ContentType="text/plain"/>
                </Types>
                """)),
            ("Config/Package.xml", XmlBytes(
                $"""
                <Package xmlns="{DataMashupNamespace}">
                  <Version>2.120.0.0</Version>
                  <MinVersion>2.120.0.0</MinVersion>
                  <Culture>en-US</Culture>
                </Package>
                """)),
            ("Formulas/Section1.m", Encoding.UTF8.GetBytes($"section Section1;\r\nshared {QueryName} = {LiteralM};\r\n")));
        var permissions = XmlBytes(
            $"""
            <PermissionList xmlns="{DataMashupNamespace}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
              <CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>
              <FirewallEnabled>true</FirewallEnabled>
              <WorkbookGroupType xsi:nil="true"/>
            </PermissionList>
            """);
        var metadataXml = XmlBytes(
            $"""
            <LocalPackageMetadataFile xmlns="{DataMashupNamespace}">
              <Items>
                <Item>
                  <ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/{QueryName}</ItemPath></ItemLocation>
                  <StableEntries>
                    <Entry Type="QueryID" Value="s8b3bd182-1d62-4606-8daa-40c35c8a973a"/>
                    <Entry Type="IsPrivate" Value="l0"/>
                    <Entry Type="FillEnabled" Value="l{(loaded ? "1" : "0")}"/>
                    <Entry Type="FillToDataModelEnabled" Value="l0"/>
                    <Entry Type="AddedToDataModel" Value="l0"/>
                    <Entry Type="FillTarget" Value="s{(loaded ? "LiteralRowsTable" : string.Empty)}"/>
                    <Entry Type="RecoveryTargetSheet" Value="s{(loaded ? WorksheetName : string.Empty)}"/>
                    <Entry Type="RecoveryTargetRow" Value="l1"/>
                    <Entry Type="RecoveryTargetColumn" Value="l1"/>
                    <Entry Type="FillColumnNames" Value="sItem,Amount"/>
                    <Entry Type="FillCount" Value="l2"/>
                    <Entry Type="FillStatus" Value="sComplete"/>
                  </StableEntries>
                </Item>
                <Item>
                  <ItemLocation><ItemType>AllFormulas</ItemType><ItemPath/></ItemLocation>
                  <StableEntries><Entry Type="IsRelationshipDetectionEnabled" Value="l0"/></StableEntries>
                </Item>
              </Items>
            </LocalPackageMetadataFile>
            """);
        var metadata = BuildMetadata(metadataXml, CreateZip());

        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(0u);
        WriteSized(writer, packageParts);
        WriteSized(writer, permissions);
        WriteSized(writer, metadata);
        WriteSized(writer, []);
        return stream.ToArray();
    }

    private static byte[] BuildMetadata(byte[] xml, byte[] content)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(0u);
        writer.Write(xml.Length);
        writer.Write(xml);
        writer.Write(content.Length);
        writer.Write(content);
        return stream.ToArray();
    }

    private static byte[] CreateZip(params (string Path, byte[] Content)[] entries)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in entries.OrderBy(item => item.Path, StringComparer.Ordinal))
            {
                WriteEntry(archive, entry.Path, entry.Content);
            }
        }
        return stream.ToArray();
    }

    private static void AuditContentTypes(
        ZipArchive archive,
        PowerQueryFixtureKind kind,
        List<string> errors)
    {
        var document = LoadXml(archive, "[Content_Types].xml", errors);
        if (document is null)
        {
            return;
        }

        XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        var overrides = document.Root?.Elements(contentTypes + "Override")
            .Select(element => (string?)element.Attribute("PartName"))
            .Where(path => path is not null)
            .ToHashSet(StringComparer.Ordinal)
            ?? [];
        foreach (var required in RequiredContentTypeOverrides.Where(
                     required => !overrides.Contains(required)))
        {
            errors.Add($"Missing content type override '{required}'.");
        }

        if (kind == PowerQueryFixtureKind.WorksheetLoaded
            && (!overrides.Contains("/xl/tables/table1.xml")
                || !overrides.Contains("/xl/queryTables/queryTable1.xml")))
        {
            errors.Add("Worksheet-loaded fixture lacks table or QueryTable content type.");
        }
    }

    private static void AuditTheme(ZipArchive archive, List<string> errors)
    {
        var theme = LoadXml(archive, "xl/theme/theme1.xml", errors);
        if (theme is null) return;
        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var matrix = theme.Root?.Element(drawing + "themeElements")?.Element(drawing + "fmtScheme");
        foreach (var name in new[] { "fillStyleLst", "lnStyleLst", "effectStyleLst", "bgFillStyleLst" })
        {
            if ((matrix?.Element(drawing + name)?.Elements().Count() ?? 0) < 3)
            {
                errors.Add($"Theme style list '{name}' requires at least three entries.");
            }
        }
    }

    private static void AuditRelationships(
        ZipArchive archive,
        PowerQueryFixtureKind kind,
        List<string> errors)
    {
        foreach (var relationshipPart in archive.Entries.Where(entry =>
                     entry.FullName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            var document = LoadXml(archive, relationshipPart.FullName, errors);
            if (document is null)
            {
                continue;
            }

            XNamespace relationships = RelationshipsNamespace;
            foreach (var relationship in document.Root?.Elements(relationships + "Relationship")
                         ?? [])
            {
                if ((string?)relationship.Attribute("TargetMode") == "External")
                {
                    errors.Add($"External relationship is forbidden in '{relationshipPart.FullName}'.");
                }
            }
        }

        if (kind == PowerQueryFixtureKind.WorksheetLoaded)
        {
            AssertRelationship(
                archive,
                "xl/worksheets/_rels/sheet1.xml.rels",
                "/table",
                "../tables/table1.xml",
                errors);
            AssertRelationship(
                archive,
                "xl/tables/_rels/table1.xml.rels",
                "/queryTable",
                "../queryTables/queryTable1.xml",
                errors);
        }
    }

    private static void AssertRelationship(
        ZipArchive archive,
        string part,
        string typeSuffix,
        string target,
        List<string> errors)
    {
        var document = LoadXml(archive, part, errors);
        XNamespace relationships = RelationshipsNamespace;
        var found = document?.Root?.Elements(relationships + "Relationship")
            .Any(element =>
                ((string?)element.Attribute("Type"))?.EndsWith(typeSuffix, StringComparison.Ordinal) == true
                && (string?)element.Attribute("Target") == target) == true;
        if (!found)
        {
            errors.Add($"Relationship '{typeSuffix}' to '{target}' is missing from '{part}'.");
        }
    }

    private static void AuditDataMashup(ZipArchive archive, List<string> errors)
    {
        var document = LoadXml(archive, "customXml/item1.xml", errors);
        if (document?.Root?.Name != XName.Get("DataMashup", DataMashupNamespace))
        {
            errors.Add("Custom XML part does not contain the MS-QDEFF DataMashup root.");
            return;
        }

        try
        {
            _ = Convert.FromBase64String(document.Root.Value);
        }
        catch (FormatException)
        {
            errors.Add("DataMashup content is not valid base64.");
        }
    }

    private static void AuditLoadGraph(
        ZipArchive archive,
        PowerQueryFixtureKind kind,
        List<string> errors)
    {
        var connections = LoadXml(archive, "xl/connections.xml", errors);
        XNamespace spreadsheet = SpreadsheetNamespace;
        var connection = connections?.Root?.Element(spreadsheet + "connection");
        var dbPr = connection?.Element(spreadsheet + "dbPr");
        if ((string?)connection?.Attribute("id") != "1"
            || !((string?)dbPr?.Attribute("connection") ?? string.Empty)
                .Contains($"Location={QueryName}", StringComparison.Ordinal))
        {
            errors.Add("Workbook connection does not identify the literal query.");
        }

        if (kind == PowerQueryFixtureKind.WorksheetLoaded)
        {
            var table = LoadXml(archive, "xl/tables/table1.xml", errors);
            var queryTable = LoadXml(archive, "xl/queryTables/queryTable1.xml", errors);
            if ((string?)table?.Root?.Attribute("ref") != "A1:B3")
            {
                errors.Add("Worksheet table range is not A1:B3.");
            }
            if ((string?)table?.Root?.Attribute("tableType") != "queryTable")
            {
                errors.Add("Worksheet tableType must be queryTable for its QueryTable relationship.");
            }
            if ((string?)queryTable?.Root?.Attribute("connectionId") != "1")
            {
                errors.Add("QueryTable is not bound to connection 1.");
            }
        }
    }

    private static XDocument? LoadXml(
        ZipArchive archive,
        string path,
        List<string> errors)
    {
        var entry = archive.GetEntry(path);
        if (entry is null)
        {
            errors.Add($"Missing XML part '{path}'.");
            return null;
        }

        try
        {
            using var stream = entry.Open();
            return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        }
        catch (Exception exception) when (exception is InvalidDataException or System.Xml.XmlException)
        {
            errors.Add($"Invalid XML part '{path}': {exception.Message}");
            return null;
        }
    }

    private static byte[] XmlBytes(string xml) =>
        new UTF8Encoding(false).GetBytes($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n{xml}");

    private static void WriteSized(BinaryWriter writer, byte[] value)
    {
        writer.Write(value.Length);
        writer.Write(value);
    }

    private static void WriteEntry(ZipArchive archive, string path, byte[] content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        entry.LastWriteTime = PackageTimestamp;
        using var target = entry.Open();
        target.Write(content);
    }

    private static string Styles() =>
        $"""
         <styleSheet xmlns="{SpreadsheetNamespace}">
           <fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>
           <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
           <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
           <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
           <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
           <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
           <dxfs count="0"/>
           <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
         </styleSheet>
         """;

    private static string Theme() =>
        """
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Repository Fixture">
          <a:themeElements>
            <a:clrScheme name="Repository Fixture">
              <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
              <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
              <a:dk2><a:srgbClr val="44546A"/></a:dk2>
              <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
              <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
              <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
              <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
              <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
              <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
              <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
              <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
              <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
            </a:clrScheme>
            <a:fontScheme name="Repository Fixture">
              <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
              <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
            </a:fontScheme>
            <a:fmtScheme name="Repository Fixture">
              <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
              </a:fillStyleLst>
              <a:lnStyleLst>
                <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
                <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
              </a:lnStyleLst>
              <a:effectStyleLst>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
                <a:effectStyle><a:effectLst/></a:effectStyle>
              </a:effectStyleLst>
              <a:bgFillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
              </a:bgFillStyleLst>
            </a:fmtScheme>
          </a:themeElements>
        </a:theme>
        """;
}
