using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacPowerQueryDefinition(string Name, string Formula);
internal sealed record MacPowerQueryWorksheetLoad(string SheetName, string Connection);

internal static class MacPowerQueryPackage
{
    private const string DataMashupNamespace = "http://schemas.microsoft.com/DataMashup";
    private const string DefaultPermissions =
        "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
        "<PermissionList xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n" +
        "  <CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n" +
        "  <FirewallEnabled>true</FirewallEnabled>\r\n" +
        "  <WorkbookGroupType xsi:nil=\"true\" />\r\n" +
        "</PermissionList>";

    public static IReadOnlyList<MacPowerQueryDefinition> ReadQueries(string workbookPath)
    {
        MashupPackage mashup;
        try
        {
            mashup = ReadMashup(workbookPath);
        }
        catch (MacPowerQueryPackageNotFoundException)
        {
            return [];
        }
        return ReadFormulaSections(mashup.PackageParts)
            .SelectMany(section => ParseQueries(section.Content))
            .Select(query => new MacPowerQueryDefinition(query.Name, query.Formula))
            .ToArray();
    }

    public static void UpdateQuery(string workbookPath, string queryName, string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(queryName);
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);

        var mashup = ReadMashup(workbookPath);
        var sections = ReadFormulaSections(mashup.PackageParts);
        FormulaSection? targetSection = null;
        ParsedQuery? targetQuery = null;
        foreach (var section in sections)
        {
            foreach (var query in ParseQueries(section.Content).Where(candidate =>
                         string.Equals(candidate.Name, queryName, StringComparison.OrdinalIgnoreCase)))
            {
                if (targetQuery is not null)
                {
                    throw new InvalidDataException(
                        $"Power Query identity '{queryName}' is ambiguous.");
                }

                targetSection = section;
                targetQuery = query;
            }
        }

        if (targetSection is null || targetQuery is null)
        {
            throw new InvalidOperationException($"Power Query '{queryName}' was not found.");
        }

        var updatedSection = string.Concat(
            targetSection.Content.AsSpan(0, targetQuery.FormulaStart),
            formula,
            targetSection.Content.AsSpan(targetQuery.FormulaStart + targetQuery.FormulaLength));
        var updatedPackageParts = ReplaceZipEntry(
            mashup.PackageParts,
            targetSection.Path,
            new UTF8Encoding(false).GetBytes(updatedSection));
        var updatedRoot = WriteRoot(
            mashup.Version,
            updatedPackageParts,
            new UTF8Encoding(false).GetBytes(DefaultPermissions),
            mashup.Metadata,
            mashup.PermissionBindings);
        WriteMashupTransactionally(workbookPath, mashup.CustomXmlPath, updatedRoot);
    }

    public static IReadOnlyList<MacPowerQueryWorksheetLoad> ReadWorksheetLoads(
        string workbookPath)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        if (archive.Entries.Any(entry =>
                entry.FullName.StartsWith("xl/model/", StringComparison.OrdinalIgnoreCase)))
        {
            throw new NotSupportedException(
                "Power Query package reads on macOS cannot yet map Data Model loads " +
                "to individual queries. No load state was returned.");
        }

        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        var connectionsEntry = archive.GetEntry("xl/connections.xml");
        if (workbookEntry is null || connectionsEntry is null)
        {
            return [];
        }

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRelationships =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var workbook = LoadXml(workbookEntry);
        var workbookRelationships = ReadRelationships(archive, "xl/workbook.xml");
        var connections = LoadXml(connectionsEntry)
            .Root?
            .Elements(spreadsheet + "connection")
            .Select(connection => new
            {
                Id = (string?)connection.Attribute("id"),
                Value = (string?)connection.Element(spreadsheet + "dbPr")?.Attribute("connection")
            })
            .Where(connection =>
                !string.IsNullOrWhiteSpace(connection.Id)
                && !string.IsNullOrWhiteSpace(connection.Value))
            .ToDictionary(
                connection => connection.Id!,
                connection => connection.Value!,
                StringComparer.Ordinal)
            ?? new Dictionary<string, string>(StringComparer.Ordinal);

        var loads = new List<MacPowerQueryWorksheetLoad>();
        foreach (var sheet in workbook.Descendants(spreadsheet + "sheet"))
        {
            var sheetName = (string?)sheet.Attribute("name");
            var relationshipId = (string?)sheet.Attribute(officeRelationships + "id");
            if (string.IsNullOrWhiteSpace(sheetName)
                || string.IsNullOrWhiteSpace(relationshipId)
                || !workbookRelationships.TryGetValue(relationshipId, out var sheetRelationship)
                || !sheetRelationship.Type.EndsWith("/worksheet", StringComparison.Ordinal))
            {
                continue;
            }

            var sheetRelationships = ReadRelationships(archive, sheetRelationship.Target);
            foreach (var tableRelationship in sheetRelationships.Values.Where(relationship =>
                         relationship.Type.EndsWith("/table", StringComparison.Ordinal)))
            {
                var tableRelationships = ReadRelationships(archive, tableRelationship.Target);
                foreach (var queryRelationship in tableRelationships.Values.Where(relationship =>
                             relationship.Type.EndsWith("/queryTable", StringComparison.Ordinal)))
                {
                    var queryTableEntry = archive.GetEntry(queryRelationship.Target)
                        ?? throw new InvalidDataException(
                            $"Power Query table relationship targets missing part '{queryRelationship.Target}'.");
                    var queryTable = LoadXml(queryTableEntry);
                    var connectionId = (string?)queryTable.Root?.Attribute("connectionId");
                    if (!string.IsNullOrWhiteSpace(connectionId)
                        && connections.TryGetValue(connectionId, out var connection))
                    {
                        loads.Add(new MacPowerQueryWorksheetLoad(sheetName, connection));
                    }
                }
            }
        }

        return loads;
    }

    private static MashupPackage ReadMashup(string workbookPath)
    {
        using var archive = ZipFile.OpenRead(workbookPath);
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("customXml/", StringComparison.Ordinal)
                || !entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var stream = entry.Open();
            XDocument document;
            try
            {
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }
            catch (XmlException)
            {
                continue;
            }

            var element = document.Root?.DescendantsAndSelf()
                .FirstOrDefault(candidate =>
                    candidate.Name.LocalName == "DataMashup"
                    && candidate.Name.NamespaceName == DataMashupNamespace);
            if (element is null)
            {
                continue;
            }

            byte[] rootBytes;
            try
            {
                rootBytes = Convert.FromBase64String(element.Value.Trim());
            }
            catch (FormatException ex)
            {
                throw new InvalidDataException(
                    $"Power Query DataMashup in '{entry.FullName}' is not valid base64.",
                    ex);
            }

            return ParseRoot(entry.FullName, rootBytes);
        }

        throw new MacPowerQueryPackageNotFoundException(
            "The workbook does not contain a Power Query DataMashup package.");
    }

    internal sealed class MacPowerQueryPackageNotFoundException(string message)
        : InvalidOperationException(message);

    private static MashupPackage ParseRoot(string customXmlPath, byte[] rootBytes)
    {
        var reader = new SpanReader(rootBytes);
        var version = reader.ReadUInt32();
        var packageParts = reader.ReadSizedBytes("PackageParts");
        var permissions = reader.ReadSizedBytes("Permissions");
        var metadata = reader.ReadSizedBytes("Metadata");
        var permissionBindings = reader.ReadSizedBytes("PermissionBindings");
        reader.EnsureComplete();
        return new MashupPackage(
            customXmlPath,
            version,
            packageParts,
            permissions,
            metadata,
            permissionBindings);
    }

    private static List<FormulaSection> ReadFormulaSections(byte[] packageParts)
    {
        using var stream = new MemoryStream(packageParts, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sections = new List<FormulaSection>();
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("Formulas/", StringComparison.Ordinal)
                || !entry.FullName.EndsWith(".m", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            using var reader = new StreamReader(
                entry.Open(),
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true);
            sections.Add(new FormulaSection(entry.FullName, reader.ReadToEnd()));
        }

        if (sections.Count == 0)
        {
            throw new InvalidDataException(
                "The Power Query package does not contain a formula section.");
        }

        return sections;
    }

    private static List<ParsedQuery> ParseQueries(string section)
    {
        var result = new List<ParsedQuery>();
        var position = 0;
        while ((position = FindKeyword(section, "shared", position)) >= 0)
        {
            var cursor = SkipTrivia(section, position + "shared".Length);
            var (name, afterName) = ReadIdentifier(section, cursor);
            cursor = SkipTrivia(section, afterName);
            if (cursor >= section.Length || section[cursor] != '=')
            {
                throw new InvalidDataException(
                    $"Power Query formula section has no '=' after shared query '{name}'.");
            }

            var formulaStart = SkipTrivia(section, cursor + 1);
            var formulaEnd = FindFormulaTerminator(section, formulaStart);
            var trimmedEnd = formulaEnd;
            while (trimmedEnd > formulaStart && char.IsWhiteSpace(section[trimmedEnd - 1]))
            {
                trimmedEnd--;
            }

            result.Add(new ParsedQuery(
                name,
                section[formulaStart..trimmedEnd],
                formulaStart,
                trimmedEnd - formulaStart));
            position = formulaEnd + 1;
        }

        return result;
    }

    private static int FindKeyword(string text, string keyword, int start)
    {
        var position = start;
        while (position < text.Length)
        {
            position = SkipTrivia(text, position);
            if (position >= text.Length)
            {
                return -1;
            }

            if (text[position] == '"')
            {
                position = SkipQuotedText(text, position);
                continue;
            }

            if (IsKeywordAt(text, keyword, position))
            {
                return position;
            }

            position++;
        }

        return -1;
    }

    private static (string Name, int Next) ReadIdentifier(string text, int position)
    {
        if (position + 1 < text.Length && text[position] == '#' && text[position + 1] == '"')
        {
            var builder = new StringBuilder();
            position += 2;
            while (position < text.Length)
            {
                if (text[position] == '"')
                {
                    if (position + 1 < text.Length && text[position + 1] == '"')
                    {
                        builder.Append('"');
                        position += 2;
                        continue;
                    }

                    return (builder.ToString(), position + 1);
                }

                builder.Append(text[position++]);
            }

            throw new InvalidDataException("Power Query quoted identifier is unterminated.");
        }

        var start = position;
        while (position < text.Length
            && !char.IsWhiteSpace(text[position])
            && text[position] != '=')
        {
            position++;
        }

        if (position == start)
        {
            throw new InvalidDataException("Power Query shared declaration has no query name.");
        }

        return (text[start..position], position);
    }

    private static int FindFormulaTerminator(string text, int position)
    {
        var round = 0;
        var square = 0;
        var curly = 0;
        while (position < text.Length)
        {
            if (position + 1 < text.Length && text[position] == '/' && text[position + 1] == '/')
            {
                position = SkipLineComment(text, position + 2);
                continue;
            }
            if (position + 1 < text.Length && text[position] == '/' && text[position + 1] == '*')
            {
                position = SkipBlockComment(text, position + 2);
                continue;
            }
            if (text[position] == '"')
            {
                position = SkipQuotedText(text, position);
                continue;
            }

            switch (text[position])
            {
                case '(':
                    round++;
                    break;
                case ')':
                    round--;
                    break;
                case '[':
                    square++;
                    break;
                case ']':
                    square--;
                    break;
                case '{':
                    curly++;
                    break;
                case '}':
                    curly--;
                    break;
                case ';' when round == 0 && square == 0 && curly == 0:
                    return position;
            }

            if (round < 0 || square < 0 || curly < 0)
            {
                throw new InvalidDataException("Power Query formula has unbalanced delimiters.");
            }
            position++;
        }

        throw new InvalidDataException("Power Query shared formula is missing its terminating ';'.");
    }

    private static int SkipTrivia(string text, int position)
    {
        while (position < text.Length)
        {
            if (char.IsWhiteSpace(text[position]))
            {
                position++;
                continue;
            }
            if (position + 1 < text.Length && text[position] == '/' && text[position + 1] == '/')
            {
                position = SkipLineComment(text, position + 2);
                continue;
            }
            if (position + 1 < text.Length && text[position] == '/' && text[position + 1] == '*')
            {
                position = SkipBlockComment(text, position + 2);
                continue;
            }
            break;
        }
        return position;
    }

    private static int SkipQuotedText(string text, int position)
    {
        position++;
        while (position < text.Length)
        {
            if (text[position] == '"')
            {
                if (position + 1 < text.Length && text[position + 1] == '"')
                {
                    position += 2;
                    continue;
                }
                return position + 1;
            }
            position++;
        }
        throw new InvalidDataException("Power Query string is unterminated.");
    }

    private static int SkipLineComment(string text, int position)
    {
        while (position < text.Length && text[position] is not ('\r' or '\n'))
        {
            position++;
        }
        return position;
    }

    private static int SkipBlockComment(string text, int position)
    {
        var depth = 1;
        while (position < text.Length && depth > 0)
        {
            if (position + 1 < text.Length && text[position] == '/' && text[position + 1] == '*')
            {
                depth++;
                position += 2;
            }
            else if (position + 1 < text.Length && text[position] == '*' && text[position + 1] == '/')
            {
                depth--;
                position += 2;
            }
            else
            {
                position++;
            }
        }
        if (depth != 0)
        {
            throw new InvalidDataException("Power Query block comment is unterminated.");
        }
        return position;
    }

    private static bool IsKeywordAt(string text, string keyword, int position)
    {
        if (position + keyword.Length > text.Length
            || !text.AsSpan(position, keyword.Length).Equals(
                keyword,
                StringComparison.Ordinal))
        {
            return false;
        }

        var beforeIsIdentifier = position > 0 && IsIdentifierCharacter(text[position - 1]);
        var after = position + keyword.Length;
        var afterIsIdentifier = after < text.Length && IsIdentifierCharacter(text[after]);
        return !beforeIsIdentifier && !afterIsIdentifier;
    }

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value is '_' or '.';

    private static byte[] ReplaceZipEntry(byte[] package, string path, byte[] content)
    {
        using var stream = new MemoryStream(package.Length + content.Length);
        stream.Write(package);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(path)
                ?? throw new InvalidDataException(
                    $"Power Query formula section '{path}' disappeared during update.");
            var timestamp = entry.LastWriteTime;
            var attributes = entry.ExternalAttributes;
            entry.Delete();
            var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
            replacement.LastWriteTime = timestamp;
            replacement.ExternalAttributes = attributes;
            using var target = replacement.Open();
            target.Write(content);
        }
        return stream.ToArray();
    }

    private static Dictionary<string, PackageRelationship> ReadRelationships(
        ZipArchive archive,
        string sourcePart)
    {
        var slash = sourcePart.LastIndexOf('/');
        var directory = slash >= 0 ? sourcePart[..slash] : string.Empty;
        var fileName = slash >= 0 ? sourcePart[(slash + 1)..] : sourcePart;
        var relationshipPath = string.IsNullOrEmpty(directory)
            ? $"_rels/{fileName}.rels"
            : $"{directory}/_rels/{fileName}.rels";
        var entry = archive.GetEntry(relationshipPath);
        if (entry is null)
        {
            return new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
        }

        XNamespace relationships =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        return LoadXml(entry)
            .Root?
            .Elements(relationships + "Relationship")
            .Select(relationship => new PackageRelationship(
                (string?)relationship.Attribute("Id") ?? string.Empty,
                (string?)relationship.Attribute("Type") ?? string.Empty,
                ResolvePackageTarget(
                    sourcePart,
                    (string?)relationship.Attribute("Target") ?? string.Empty)))
            .Where(relationship =>
                !string.IsNullOrWhiteSpace(relationship.Id)
                && !string.IsNullOrWhiteSpace(relationship.Target))
            .ToDictionary(
                relationship => relationship.Id,
                relationship => relationship,
                StringComparer.Ordinal)
            ?? new Dictionary<string, PackageRelationship>(StringComparer.Ordinal);
    }

    private static string ResolvePackageTarget(string sourcePart, string target)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            return string.Empty;
        }

        var sourceUri = new Uri($"https://excelmcp.invalid/{sourcePart}");
        var resolved = new Uri(sourceUri, target.Replace('\\', '/'));
        if (!string.Equals(resolved.Host, sourceUri.Host, StringComparison.Ordinal))
        {
            return string.Empty;
        }

        return Uri.UnescapeDataString(resolved.AbsolutePath).TrimStart('/');
    }

    private static XDocument LoadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static byte[] WriteRoot(
        uint version,
        byte[] packageParts,
        byte[] permissions,
        byte[] metadata,
        byte[] permissionBindings)
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true);
        writer.Write(version);
        WriteSizedBytes(writer, packageParts);
        WriteSizedBytes(writer, permissions);
        WriteSizedBytes(writer, metadata);
        WriteSizedBytes(writer, permissionBindings);
        return stream.ToArray();
    }

    private static void WriteMashupTransactionally(
        string workbookPath,
        string customXmlPath,
        byte[] root)
    {
        var directory = Path.GetDirectoryName(workbookPath)
            ?? throw new ArgumentException("Workbook path has no parent directory.", nameof(workbookPath));
        var temporaryPath = Path.Combine(directory, $".excelmcp-pq-{Guid.NewGuid():N}.tmp");
        try
        {
            File.Copy(workbookPath, temporaryPath);
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry(customXmlPath)
                    ?? throw new InvalidDataException(
                        $"Power Query custom XML part '{customXmlPath}' disappeared during update.");
                XDocument document;
                using (var source = entry.Open())
                {
                    document = XDocument.Load(source, LoadOptions.PreserveWhitespace);
                }
                var element = document.Root?.DescendantsAndSelf()
                    .Single(candidate =>
                        candidate.Name.LocalName == "DataMashup"
                        && candidate.Name.NamespaceName == DataMashupNamespace);
                if (element is null)
                {
                    throw new InvalidDataException(
                        $"Power Query custom XML part '{customXmlPath}' no longer contains DataMashup.");
                }
                element.Value = Convert.ToBase64String(root);

                var timestamp = entry.LastWriteTime;
                var attributes = entry.ExternalAttributes;
                entry.Delete();
                var replacement = archive.CreateEntry(customXmlPath, CompressionLevel.Optimal);
                replacement.LastWriteTime = timestamp;
                replacement.ExternalAttributes = attributes;
                using var target = replacement.Open();
                using var xmlWriter = XmlWriter.Create(
                    target,
                    new XmlWriterSettings
                    {
                        Encoding = new UTF8Encoding(false),
                        Indent = false,
                        CloseOutput = false
                    });
                document.Save(xmlWriter);
            }

            File.Move(temporaryPath, workbookPath, overwrite: true);
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }

    private static void WriteSizedBytes(BinaryWriter writer, byte[] value)
    {
        writer.Write(value.Length);
        writer.Write(value);
    }

    private sealed record MashupPackage(
        string CustomXmlPath,
        uint Version,
        byte[] PackageParts,
        byte[] Permissions,
        byte[] Metadata,
        byte[] PermissionBindings);

    private sealed record FormulaSection(string Path, string Content);

    private sealed record PackageRelationship(string Id, string Type, string Target);

    private sealed record ParsedQuery(
        string Name,
        string Formula,
        int FormulaStart,
        int FormulaLength);

    private ref struct SpanReader(ReadOnlySpan<byte> source)
    {
        private readonly ReadOnlySpan<byte> _source = source;
        private int _position;

        public uint ReadUInt32()
        {
            EnsureAvailable(sizeof(uint), "version");
            var result = BinaryPrimitives.ReadUInt32LittleEndian(_source[_position..]);
            _position += sizeof(uint);
            return result;
        }

        public byte[] ReadSizedBytes(string fieldName)
        {
            EnsureAvailable(sizeof(int), $"{fieldName} length");
            var length = BinaryPrimitives.ReadInt32LittleEndian(_source[_position..]);
            _position += sizeof(int);
            if (length < 0)
            {
                throw new InvalidDataException(
                    $"Power Query DataMashup {fieldName} length is negative.");
            }
            EnsureAvailable(length, fieldName);
            var result = _source.Slice(_position, length).ToArray();
            _position += length;
            return result;
        }

        public void EnsureComplete()
        {
            if (_position != _source.Length)
            {
                throw new InvalidDataException(
                    "Power Query DataMashup contains trailing bytes.");
            }
        }

        private readonly void EnsureAvailable(int length, string fieldName)
        {
            if (length > _source.Length - _position)
            {
                throw new InvalidDataException(
                    $"Power Query DataMashup {fieldName} exceeds the available data.");
            }
        }
    }
}
