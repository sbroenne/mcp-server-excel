using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Utility class for validating file access and locking status.
/// Provides OS-level file lock detection and IRM/AIP-encryption detection before Excel COM operations.
/// </summary>
public static class FileAccessValidator
{
    // OLE2 Compound Document Format signature.
    // IRM/AIP-protected Excel files are stored as OLE2 containers with an EncryptedPackage
    // stream instead of the standard ZIP-based Office Open XML format.
    private static ReadOnlySpan<byte> Ole2Signature => [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private const string DataSpacesStorage = "\u0006DataSpaces";
    private const string DataSpaceMapStream = "DataSpaceMap";
    private const string DataSpaceInfoStorage = "DataSpaceInfo";
    private static readonly HashSet<string> IrmDataSpaceNames =
        new(StringComparer.Ordinal)
        {
            "\tDRMDataSpace",
            "DRMEncryptedDataSpace"
        };

    /// <summary>
    /// Detects if the file is IRM/AIP-protected by checking for both the OLE2 compound
    /// document signature and its structured data-space map and matching definition.
    /// Ordinary password-encrypted OOXML also uses OLE2 data spaces and must not be
    /// classified as IRM.
    /// Legacy .xls files are always OLE2 by design and are excluded.
    /// IRM-protected files must be opened as read-only with Excel visible so the user can
    /// authenticate through the Information Rights Management credential prompt.
    /// </summary>
    /// <param name="filePath">The file path to inspect.</param>
    /// <returns>
    /// <c>true</c> if the file has the OLE2 Compound Document header, uses a modern
    /// OOXML extension (.xlsx, .xlsm, .xlsb), and maps protected content to the legacy
    /// <c>\tDRMDataSpace</c> or modern <c>DRMEncryptedDataSpace</c> definition;
    /// <c>false</c> for legacy .xls files (always OLE2), standard ZIP-based files, or
    /// if the file cannot be read.
    /// </returns>
    public static bool IsIrmProtected(string filePath)
    {
        if (!File.Exists(filePath))
            return false;

        // Legacy .xls/.xlt files are always OLE2 compound documents by design.
        // They are NOT IRM-protected just because they have an OLE2 header.
        var ext = Path.GetExtension(filePath);
        if (ext.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
            ext.Equals(".xlt", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            Span<byte> signature = stackalloc byte[8];
            using (var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite))
            {
                if (stream.Read(signature) < signature.Length
                    || !signature.SequenceEqual(Ole2Signature))
                {
                    return false;
                }
            }

            using var compoundFile = OleCompoundFileReader.Open(filePath);
            if (!compoundFile.TryReadStream(
                    $"{DataSpacesStorage}/{DataSpaceMapStream}",
                    out var map))
            {
                return false;
            }

            foreach (var dataSpaceName in ReadDataSpaceNames(map))
            {
                if (!IrmDataSpaceNames.Contains(dataSpaceName)
                    || !compoundFile.TryReadStream(
                        $"{DataSpacesStorage}/{DataSpaceInfoStorage}/{dataSpaceName}",
                        out var definition))
                {
                    continue;
                }

                if (IsValidDataSpaceDefinition(definition))
                {
                    return true;
                }
            }

            return false;
        }
        catch (InvalidDataException)
        {
            return false;
        }
        catch (OverflowException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            // Cannot read → treat as not IRM so normal error handling takes over
            return false;
        }
    }

    private static List<string> ReadDataSpaceNames(ReadOnlySpan<byte> map)
    {
        const int maximumMapEntries = 1024;
        if (map.Length < 8
            || BinaryPrimitives.ReadUInt32LittleEndian(map) != 8)
        {
            return [];
        }

        var entryCount = BinaryPrimitives.ReadUInt32LittleEndian(map[4..]);
        if (entryCount > maximumMapEntries)
        {
            return [];
        }

        var names = new List<string>(checked((int)entryCount));
        var offset = 8;
        for (var entryIndex = 0u; entryIndex < entryCount; entryIndex++)
        {
            if (!TryReadUInt32(map, ref offset, out var entryLength)
                || entryLength < 12
                || entryLength > int.MaxValue)
            {
                return [];
            }

            var entryStart = offset - sizeof(uint);
            var entryEnd = checked(entryStart + (int)entryLength);
            if (entryEnd > map.Length
                || !TryReadUInt32(map, ref offset, out var componentCount)
                || componentCount is 0 or > 64)
            {
                return [];
            }

            var hasStreamReference = false;
            for (var componentIndex = 0u; componentIndex < componentCount; componentIndex++)
            {
                if (!TryReadUInt32(map[..entryEnd], ref offset, out var componentType)
                    || !TryReadUnicodeLpP4(map[..entryEnd], ref offset, out _))
                {
                    return [];
                }

                hasStreamReference |= componentType == 0;
            }

            if (!TryReadUnicodeLpP4(map[..entryEnd], ref offset, out var dataSpaceName)
                || offset != entryEnd)
            {
                return [];
            }

            if (hasStreamReference)
            {
                names.Add(dataSpaceName);
            }
        }

        return names;
    }

    private static bool IsValidDataSpaceDefinition(ReadOnlySpan<byte> definition)
    {
        const int maximumTransformReferences = 64;
        if (definition.Length < 8
            || BinaryPrimitives.ReadUInt32LittleEndian(definition) != 8)
        {
            return false;
        }

        var transformCount = BinaryPrimitives.ReadUInt32LittleEndian(definition[4..]);
        if (transformCount is 0 or > maximumTransformReferences)
        {
            return false;
        }

        var offset = 8;
        for (var index = 0u; index < transformCount; index++)
        {
            if (!TryReadUnicodeLpP4(definition, ref offset, out var transformName)
                || string.IsNullOrWhiteSpace(transformName))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryReadUnicodeLpP4(
        ReadOnlySpan<byte> bytes,
        ref int offset,
        out string value)
    {
        value = string.Empty;
        if (!TryReadUInt32(bytes, ref offset, out var byteLength)
            || byteLength > int.MaxValue
            || byteLength % 2 != 0
            || byteLength > bytes.Length - offset)
        {
            return false;
        }

        value = Encoding.Unicode.GetString(
            bytes.Slice(offset, checked((int)byteLength)));
        offset += checked((int)byteLength);
        var alignedOffset = checked((offset + 3) & ~3);
        if (alignedOffset > bytes.Length)
        {
            return false;
        }

        offset = alignedOffset;
        return true;
    }

    private static bool TryReadUInt32(
        ReadOnlySpan<byte> bytes,
        ref int offset,
        out uint value)
    {
        value = 0;
        if (offset < 0 || offset > bytes.Length - sizeof(uint))
        {
            return false;
        }

        value = BinaryPrimitives.ReadUInt32LittleEndian(bytes[offset..]);
        offset += sizeof(uint);
        return true;
    }

    /// <summary>
    /// Validates a standard OOXML workbook package without launching Excel.
    /// Encrypted IRM/AIP compound documents require interactive Excel validation
    /// and are intentionally not accepted by this preflight validator.
    /// </summary>
    public static bool HasValidWorkbookContainer(string filePath)
    {
        const string contentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        const string packageRelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        const string spreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string strictSpreadsheetNamespace =
            "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string officeDocumentRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        const string strictOfficeDocumentRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
        const string worksheetRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        const string strictWorksheetRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet";
        const string chartSheetRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
        const string strictChartSheetRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
        const string dialogSheetRelationship =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
        const string strictDialogSheetRelationship =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
        const string macroSheetRelationship =
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
        const string internationalMacroSheetRelationship =
            "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
        const string macroSheetNamespace =
            "http://schemas.microsoft.com/office/excel/2006/main";

        try
        {
            using var stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var contentTypes = LoadXmlPart(archive, "[Content_Types].xml");
            var relationships = LoadXmlPart(archive, "_rels/.rels");
            var workbook = LoadXmlPart(archive, "xl/workbook.xml");
            var workbookRelationships = LoadXmlPart(archive, "xl/_rels/workbook.xml.rels");
            if (contentTypes?.Root?.Name
                    != XName.Get("Types", contentTypesNamespace)
                || relationships?.Root?.Name
                    != XName.Get("Relationships", packageRelationshipsNamespace)
                || workbookRelationships?.Root?.Name
                    != XName.Get("Relationships", packageRelationshipsNamespace)
                || workbook?.Root?.Name.LocalName != "workbook"
                || workbook.Root.Name.NamespaceName is not (
                    spreadsheetNamespace or strictSpreadsheetNamespace))
            {
                return false;
            }

            var hasWorkbookContentType = contentTypes.Root.Elements()
                .Any(element =>
                    element.Name == XName.Get("Override", contentTypesNamespace)
                    && string.Equals(
                        element.Attribute("PartName")?.Value,
                        "/xl/workbook.xml",
                        StringComparison.OrdinalIgnoreCase)
                    && element.Attribute("ContentType")?.Value is
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                        or "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
            var hasWorkbookRelationship = relationships.Root.Elements()
                .Any(element =>
                    element.Name == XName.Get("Relationship", packageRelationshipsNamespace)
                    && element.Attribute("Type")?.Value is
                        officeDocumentRelationship or strictOfficeDocumentRelationship
                    && string.Equals(
                        element.Attribute("Target")?.Value.TrimStart('/'),
                        "xl/workbook.xml",
                        StringComparison.OrdinalIgnoreCase));
            if (!hasWorkbookContentType || !hasWorkbookRelationship)
            {
                return false;
            }

            var workbookNamespace = workbook.Root.Name.NamespaceName;
            var relationshipAttributeNamespace =
                workbookNamespace == strictSpreadsheetNamespace
                    ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                    : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var sheets = workbook.Root
                .Element(XName.Get("sheets", workbookNamespace))?
                .Elements(XName.Get("sheet", workbookNamespace))
                .ToList();
            if (sheets is not { Count: > 0 })
            {
                return false;
            }

            var relationshipElements = workbookRelationships.Root
                .Elements(XName.Get("Relationship", packageRelationshipsNamespace))
                .ToList();
            foreach (var sheet in sheets)
            {
                var relationshipId = sheet
                    .Attribute(XName.Get("id", relationshipAttributeNamespace))?
                    .Value;
                var sheetRelationship = relationshipElements.FirstOrDefault(element =>
                    string.Equals(
                        element.Attribute("Id")?.Value,
                        relationshipId,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        element.Attribute("TargetMode")?.Value,
                        "External",
                        StringComparison.OrdinalIgnoreCase));
                var sheetRelationshipType =
                    sheetRelationship?.Attribute("Type")?.Value;
                var sheetPart = sheetRelationshipType switch
                {
                    worksheetRelationship or strictWorksheetRelationship =>
                        new SheetPartDescriptor(
                            "worksheet",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                            workbookNamespace),
                    chartSheetRelationship or strictChartSheetRelationship =>
                        new SheetPartDescriptor(
                            "chartsheet",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
                            workbookNamespace),
                    dialogSheetRelationship or strictDialogSheetRelationship =>
                        new SheetPartDescriptor(
                            "dialogsheet",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
                            workbookNamespace),
                    macroSheetRelationship =>
                        new SheetPartDescriptor(
                            "macrosheet",
                            "application/vnd.ms-excel.macrosheet+xml",
                            macroSheetNamespace),
                    internationalMacroSheetRelationship =>
                        new SheetPartDescriptor(
                            "macrosheet",
                            "application/vnd.ms-excel.intlmacrosheet+xml",
                            macroSheetNamespace),
                    _ => null
                };
                var sheetTarget = sheetRelationship?.Attribute("Target")?.Value
                    .Replace('\\', '/')
                    .TrimStart('/');
                if (sheetPart == null
                    || string.IsNullOrWhiteSpace(sheetTarget)
                    || sheetTarget.Split('/').Contains("..", StringComparer.Ordinal))
                {
                    return false;
                }

                var sheetEntryPath = sheetTarget.StartsWith(
                    "xl/",
                    StringComparison.OrdinalIgnoreCase)
                    ? sheetTarget
                    : $"xl/{sheetTarget}";
                var hasSheetContentType = contentTypes.Root.Elements()
                    .Any(element =>
                        element.Name == XName.Get("Override", contentTypesNamespace)
                        && string.Equals(
                            element.Attribute("PartName")?.Value.TrimStart('/'),
                            sheetEntryPath,
                            StringComparison.OrdinalIgnoreCase)
                        && string.Equals(
                            element.Attribute("ContentType")?.Value,
                            sheetPart.ContentType,
                            StringComparison.Ordinal));
                if (!hasSheetContentType
                    || !HasExpectedXmlRoot(
                        archive,
                        sheetEntryPath,
                        sheetPart.RootElement,
                        sheetPart.RootNamespace))
                {
                    return false;
                }
            }

            return true;
        }
        catch (InvalidDataException)
        {
            return false;
        }
        catch (IOException)
        {
            return false;
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
        catch (XmlException)
        {
            return false;
        }
    }

    private static XDocument? LoadXmlPart(ZipArchive archive, string name)
    {
        const long maximumMetadataPartBytes = 4 * 1024 * 1024;
        var entry = archive.GetEntry(name);
        if (entry == null || entry.Length > maximumMetadataPartBytes)
        {
            return null;
        }

        using var part = entry.Open();
        using var reader = XmlReader.Create(part, CreateSafeXmlReaderSettings(maximumMetadataPartBytes));
        return XDocument.Load(reader, LoadOptions.None);
    }

    private static bool HasExpectedXmlRoot(
        ZipArchive archive,
        string name,
        string expectedLocalName,
        string expectedNamespace)
    {
        const long maximumRootProbeCharacters = 1024 * 1024;
        var entry = archive.GetEntry(name);
        if (entry == null)
        {
            return false;
        }

        using var part = entry.Open();
        using var reader = XmlReader.Create(
            part,
            CreateSafeXmlReaderSettings(maximumRootProbeCharacters));
        reader.MoveToContent();
        return reader.LocalName == expectedLocalName
            && reader.NamespaceURI == expectedNamespace;
    }

    private static XmlReaderSettings CreateSafeXmlReaderSettings(long maximumCharacters) =>
        new()
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = maximumCharacters
        };

    private sealed record SheetPartDescriptor(
        string RootElement,
        string ContentType,
        string RootNamespace);

    /// <summary>
    /// Validates that a file is not locked by attempting to open it with exclusive access.
    /// Throws InvalidOperationException if file is locked or inaccessible.
    /// This is a fast OS-level check that doesn't require launching Excel.
    /// </summary>
    /// <param name="filePath">The file path to validate</param>
    /// <exception cref="InvalidOperationException">Thrown when file is locked or inaccessible</exception>
    public static void ValidateFileNotLocked(string filePath)
    {
        try
        {
            using var lockTest = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            // File is NOT locked - close and proceed
        }
        catch (IOException ioEx)
        {
            // File is locked by another process (most likely already open in Excel)
            throw CreateFileLockedError(filePath, ioEx);
        }
        catch (UnauthorizedAccessException uaEx)
        {
            // File access denied (permissions issue or file is locked)
            throw new InvalidOperationException(
                $"Cannot access '{Path.GetFileName(filePath)}'. " +
                "The file may be read-only, you may lack permissions, or it's locked by another process. " +
                "Please verify file permissions and close any applications using this file.",
                uaEx);
        }
    }

    /// <summary>
    /// Creates a standardized InvalidOperationException for file-locked scenarios.
    /// Provides consistent error messages across the codebase.
    /// </summary>
    /// <param name="filePath">The file path that is locked</param>
    /// <param name="innerException">The underlying exception that triggered the error</param>
    /// <returns>A user-friendly InvalidOperationException with guidance</returns>
    public static InvalidOperationException CreateFileLockedError(string filePath, Exception innerException)
    {
        return new InvalidOperationException(
            $"Cannot open '{Path.GetFileName(filePath)}'. " +
            "The file is already open in Excel or another process is using it. " +
            "Please close the file before running automation commands. " +
            "ExcelMcp requires exclusive access to workbooks during operations.",
            innerException);
    }
}
