using System.Buffers.Binary;
using System.Text;

namespace Sbroenne.ExcelMcp.ComInterop;

internal sealed class OleCompoundFileReader : IDisposable
{
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FatSector = 0xFFFFFFFD;
    private const uint DifatSector = 0xFFFFFFFC;
    private const string DataSpacesStorage = "\u0006DataSpaces";
    private const string DataSpaceMapStream = "DataSpaceMap";
    private const string DataSpaceInfoStorage = "DataSpaceInfo";
    private const string LegacyDrmDataSpace = "\tDRMDataSpace";
    private const string ModernDrmDataSpace = "DRMEncryptedDataSpace";
    private readonly FileStream _stream;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint _miniStreamCutoff;
    private readonly uint _firstMiniFatSector;
    private readonly uint _miniFatSectorCount;
    private readonly uint[] _fat;
    private readonly DirectoryEntry[] _directoryEntries;
    private readonly Dictionary<string, DirectoryEntry> _entriesByPath =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly DirectoryEntry _rootEntry;
    private readonly ushort _majorVersion;
    private byte[]? _miniStream;
    private uint[]? _miniFat;

    private OleCompoundFileReader(FileStream stream)
    {
        _stream = stream;
        Span<byte> header = stackalloc byte[512];
        stream.ReadExactly(header);
        ReadOnlySpan<byte> signature =
            [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
        if (!header[..8].SequenceEqual(signature))
        {
            throw new InvalidDataException("The file is not an OLE compound file.");
        }

        _majorVersion = BinaryPrimitives.ReadUInt16LittleEndian(header[26..]);
        var sectorShift = BinaryPrimitives.ReadUInt16LittleEndian(header[30..]);
        var miniSectorShift = BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
        if (BinaryPrimitives.ReadUInt16LittleEndian(header[28..]) != 0xFFFE
            || _majorVersion is not (3 or 4)
            || sectorShift != (_majorVersion == 3 ? 9 : 12)
            || miniSectorShift != 6)
        {
            throw new InvalidDataException("The OLE compound header is invalid.");
        }

        _sectorSize = 1 << sectorShift;
        _miniSectorSize = 1 << miniSectorShift;
        var sectorCount = GetSectorCount(stream.Length, _sectorSize);
        var fatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(header[44..]);
        var firstDirectorySector = BinaryPrimitives.ReadUInt32LittleEndian(header[48..]);
        _miniStreamCutoff = BinaryPrimitives.ReadUInt32LittleEndian(header[56..]);
        _firstMiniFatSector = BinaryPrimitives.ReadUInt32LittleEndian(header[60..]);
        _miniFatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(header[64..]);
        var firstDifatSector = BinaryPrimitives.ReadUInt32LittleEndian(header[68..]);
        var difatSectorCount = BinaryPrimitives.ReadUInt32LittleEndian(header[72..]);

        var fatSectorIds = ReadFatSectorIds(
            header,
            fatSectorCount,
            firstDifatSector,
            difatSectorCount,
            sectorCount);
        _fat = ReadFat(fatSectorIds, sectorCount);
        var directoryBytes = ReadRegularChain(firstDirectorySector, maximumSectors: sectorCount);
        _directoryEntries = ReadDirectoryEntries(directoryBytes);
        _rootEntry = _directoryEntries.FirstOrDefault(entry => entry.Type == 5)
            ?? throw new InvalidDataException("The OLE root storage is missing.");
        IndexRelevantChildren(
            _rootEntry.Child,
            DirectorySearchScope.Root,
            new HashSet<int>());
    }

    internal static OleCompoundFileReader Open(string filePath)
    {
        var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite);
        try
        {
            return new OleCompoundFileReader(stream);
        }

        catch
        {
            stream.Dispose();
            throw;
        }
    }

    internal static int GetSectorCount(long streamLength, int sectorSize)
    {
        if (sectorSize <= 0
            || streamLength < sectorSize
            || streamLength % sectorSize != 0)
        {
            throw new InvalidDataException("The OLE compound file length is invalid.");
        }

        var sectorCount = streamLength / sectorSize - 1;
        if (sectorCount > int.MaxValue)
        {
            throw new InvalidDataException(
                "The OLE compound file length exceeds the inspection limit.");
        }

        return (int)sectorCount;
    }

    internal bool TryReadStream(string path, out byte[] contents)
    {
        contents = [];
        if (!_entriesByPath.TryGetValue(path, out var entry)
            || entry.Type != 2
            || entry.StreamSize < 0
            || entry.StreamSize > int.MaxValue)
        {
            return false;
        }

        contents = entry.StreamSize < _miniStreamCutoff
            ? ReadMiniStream(entry)
            : ReadRegularStream(entry);
        return true;
    }

    private List<uint> ReadFatSectorIds(
        ReadOnlySpan<byte> header,
        uint fatSectorCount,
        uint firstDifatSector,
        uint difatSectorCount,
        int sectorCount)
    {
        if (fatSectorCount > sectorCount || difatSectorCount > sectorCount)
        {
            throw new InvalidDataException("The OLE FAT metadata is invalid.");
        }

        var ids = new List<uint>(checked((int)fatSectorCount));
        for (var offset = 76; offset < 512 && ids.Count < fatSectorCount; offset += sizeof(uint))
        {
            AddFatSectorId(
                BinaryPrimitives.ReadUInt32LittleEndian(header[offset..]),
                ids,
                sectorCount);
        }

        var difatSector = firstDifatSector;
        var visited = new HashSet<uint>();
        for (var index = 0u; index < difatSectorCount && ids.Count < fatSectorCount; index++)
        {
            if (!IsRegularSector(difatSector, sectorCount) || !visited.Add(difatSector))
            {
                throw new InvalidDataException("The OLE DIFAT chain is invalid.");
            }

            var sector = ReadSector(difatSector);
            var entryCount = _sectorSize / sizeof(uint) - 1;
            for (var entryIndex = 0; entryIndex < entryCount && ids.Count < fatSectorCount; entryIndex++)
            {
                AddFatSectorId(
                    BinaryPrimitives.ReadUInt32LittleEndian(
                        sector.AsSpan(entryIndex * sizeof(uint))),
                    ids,
                    sectorCount);
            }

            difatSector = BinaryPrimitives.ReadUInt32LittleEndian(
                sector.AsSpan(entryCount * sizeof(uint)));
        }

        if (ids.Count != fatSectorCount)
        {
            throw new InvalidDataException("The OLE FAT sector list is incomplete.");
        }

        return ids;
    }

    private static void AddFatSectorId(uint id, List<uint> ids, int sectorCount)
    {
        if (id == FreeSector)
        {
            return;
        }

        if (!IsRegularSector(id, sectorCount))
        {
            throw new InvalidDataException("The OLE FAT contains an invalid sector.");
        }

        ids.Add(id);
    }

    private uint[] ReadFat(List<uint> fatSectorIds, int sectorCount)
    {
        var entries = new List<uint>(fatSectorIds.Count * (_sectorSize / sizeof(uint)));
        foreach (var sectorId in fatSectorIds)
        {
            var sector = ReadSector(sectorId);
            for (var offset = 0; offset < sector.Length; offset += sizeof(uint))
            {
                entries.Add(BinaryPrimitives.ReadUInt32LittleEndian(sector.AsSpan(offset)));
            }
        }

        if (entries.Count < sectorCount)
        {
            throw new InvalidDataException("The OLE FAT does not cover the file.");
        }

        return [.. entries];
    }

    private DirectoryEntry[] ReadDirectoryEntries(byte[] bytes)
    {
        if (bytes.Length % 128 != 0)
        {
            throw new InvalidDataException("The OLE directory stream is invalid.");
        }

        var entries = new List<DirectoryEntry>(bytes.Length / 128);
        for (var index = 0; index < bytes.Length / 128; index++)
        {
            var entry = bytes.AsSpan(index * 128, 128);
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(entry[64..]);
            var type = entry[66];
            string name;
            if (type == 0)
            {
                name = string.Empty;
            }
            else
            {
                if (nameLength is < 2 or > 64 || nameLength % 2 != 0)
                {
                    throw new InvalidDataException("An OLE directory name is invalid.");
                }

                name = Encoding.Unicode.GetString(entry[..(nameLength - 2)]);
            }

            var rawSize = BinaryPrimitives.ReadInt64LittleEndian(entry[120..]);
            var streamSize = _majorVersion == 3
                ? unchecked((uint)rawSize)
                : rawSize;
            entries.Add(new DirectoryEntry(
                index,
                name,
                type,
                BinaryPrimitives.ReadInt32LittleEndian(entry[68..]),
                BinaryPrimitives.ReadInt32LittleEndian(entry[72..]),
                BinaryPrimitives.ReadInt32LittleEndian(entry[76..]),
                BinaryPrimitives.ReadUInt32LittleEndian(entry[116..]),
                streamSize));
        }

        return [.. entries];
    }

    private void IndexRelevantChildren(
        int entryId,
        DirectorySearchScope scope,
        HashSet<int> visited)
    {
        var pending = new Stack<DirectoryTraversal>();
        pending.Push(new DirectoryTraversal(
            entryId,
            scope,
            DirectoryTraversalPhase.Enter));
        while (pending.TryPop(out var traversal))
        {
            if (traversal.EntryId == -1)
            {
                continue;
            }

            var entry = GetDirectoryEntry(traversal.EntryId);
            if (traversal.Phase == DirectoryTraversalPhase.Enter)
            {
                if (!visited.Add(traversal.EntryId))
                {
                    throw new InvalidDataException("The OLE directory tree contains a cycle.");
                }

                pending.Push(traversal with { Phase = DirectoryTraversalPhase.Visit });
                pending.Push(new DirectoryTraversal(
                    entry.LeftSibling,
                    traversal.Scope,
                    DirectoryTraversalPhase.Enter));
                continue;
            }

            if (traversal.Phase == DirectoryTraversalPhase.Visit)
            {
                pending.Push(new DirectoryTraversal(
                    entry.RightSibling,
                    traversal.Scope,
                    DirectoryTraversalPhase.Enter));

                var childScope = GetChildSearchScope(traversal.Scope, entry);
                var relevantPath = GetRelevantStreamPath(traversal.Scope, entry);
                if (relevantPath != null
                    && !_entriesByPath.TryAdd(relevantPath, entry))
                {
                    throw new InvalidDataException("The OLE directory contains duplicate paths.");
                }

                if (childScope is { } relevantChildScope)
                {
                    pending.Push(new DirectoryTraversal(
                        entry.Child,
                        relevantChildScope,
                        DirectoryTraversalPhase.Enter));
                }

                continue;
            }
        }
    }

    private static DirectorySearchScope? GetChildSearchScope(
        DirectorySearchScope scope,
        DirectoryEntry entry)
    {
        if (entry.Type != 1)
        {
            return null;
        }

        return scope switch
        {
            DirectorySearchScope.Root
                when entry.Name.Equals(DataSpacesStorage, StringComparison.OrdinalIgnoreCase) =>
                DirectorySearchScope.DataSpaces,
            DirectorySearchScope.DataSpaces
                when entry.Name.Equals(DataSpaceInfoStorage, StringComparison.OrdinalIgnoreCase) =>
                DirectorySearchScope.DataSpaceInfo,
            _ => null
        };
    }

    private static string? GetRelevantStreamPath(
        DirectorySearchScope scope,
        DirectoryEntry entry)
    {
        if (entry.Type != 2)
        {
            return null;
        }

        return scope switch
        {
            DirectorySearchScope.DataSpaces
                when entry.Name.Equals(DataSpaceMapStream, StringComparison.OrdinalIgnoreCase) =>
                $"{DataSpacesStorage}/{DataSpaceMapStream}",
            DirectorySearchScope.DataSpaceInfo
                when entry.Name.Equals(LegacyDrmDataSpace, StringComparison.OrdinalIgnoreCase) =>
                $"{DataSpacesStorage}/{DataSpaceInfoStorage}/{LegacyDrmDataSpace}",
            DirectorySearchScope.DataSpaceInfo
                when entry.Name.Equals(ModernDrmDataSpace, StringComparison.OrdinalIgnoreCase) =>
                $"{DataSpacesStorage}/{DataSpaceInfoStorage}/{ModernDrmDataSpace}",
            _ => null
        };
    }

    private DirectoryEntry GetDirectoryEntry(int entryId)
    {
        if ((uint)entryId >= _directoryEntries.Length)
        {
            throw new InvalidDataException("The OLE directory references an invalid entry.");
        }

        return _directoryEntries[entryId];
    }

    private byte[] ReadRegularStream(DirectoryEntry entry)
    {
        if (entry.StreamSize < 0 || entry.StreamSize > int.MaxValue)
        {
            throw new InvalidDataException("An OLE stream size exceeds the inspection limit.");
        }

        var bytes = ReadRegularChain(
            entry.StartSector,
            maximumSectors: (int)((entry.StreamSize + _sectorSize - 1) / _sectorSize));
        if (bytes.Length < entry.StreamSize)
        {
            throw new InvalidDataException("An OLE stream chain is shorter than its declared size.");
        }

        return bytes[..checked((int)entry.StreamSize)];
    }

    private byte[] ReadRegularChain(uint startSector, int maximumSectors)
    {
        if (maximumSectors < 0)
        {
            throw new InvalidDataException("The OLE stream size is invalid.");
        }

        using var output = new MemoryStream();
        var sectorId = startSector;
        var visited = new HashSet<uint>();
        while (sectorId != EndOfChain)
        {
            if (sectorId >= _fat.Length
                || sectorId is FreeSector or FatSector or DifatSector
                || !visited.Add(sectorId)
                || visited.Count > maximumSectors)
            {
                throw new InvalidDataException("An OLE stream chain is invalid.");
            }

            output.Write(ReadSector(sectorId));
            sectorId = _fat[sectorId];
        }

        return output.ToArray();
    }

    private byte[] ReadMiniStream(DirectoryEntry entry)
    {
        EnsureMiniStreamLoaded();
        if (_miniFat == null || _miniStream == null)
        {
            throw new InvalidDataException("The OLE mini stream is unavailable.");
        }

        var expectedSectors = checked((int)((entry.StreamSize + _miniSectorSize - 1) / _miniSectorSize));
        using var output = new MemoryStream();
        var sectorId = entry.StartSector;
        var visited = new HashSet<uint>();
        while (sectorId != EndOfChain)
        {
            if (sectorId >= _miniFat.Length
                || !visited.Add(sectorId)
                || visited.Count > expectedSectors)
            {
                throw new InvalidDataException("An OLE mini stream chain is invalid.");
            }

            var offset = checked((long)sectorId * _miniSectorSize);
            if (offset + _miniSectorSize > _miniStream.Length)
            {
                throw new InvalidDataException("An OLE mini stream sector is outside the root stream.");
            }

            output.Write(_miniStream, checked((int)offset), _miniSectorSize);
            sectorId = _miniFat[sectorId];
        }

        var bytes = output.ToArray();
        if (bytes.Length < entry.StreamSize)
        {
            throw new InvalidDataException("An OLE mini stream is shorter than its declared size.");
        }

        return bytes[..checked((int)entry.StreamSize)];
    }

    private void EnsureMiniStreamLoaded()
    {
        if (_miniFat != null)
        {
            return;
        }

        if (_miniFatSectorCount == 0
            || _firstMiniFatSector == EndOfChain
            || _rootEntry.StreamSize <= 0)
        {
            _miniFat = [];
            _miniStream = [];
            return;
        }

        var miniFatBytes = ReadRegularChain(
            _firstMiniFatSector,
            checked((int)_miniFatSectorCount));
        var miniFat = new uint[miniFatBytes.Length / sizeof(uint)];
        for (var index = 0; index < miniFat.Length; index++)
        {
            miniFat[index] = BinaryPrimitives.ReadUInt32LittleEndian(
                miniFatBytes.AsSpan(index * sizeof(uint)));
        }

        _miniFat = miniFat;
        _miniStream = ReadRegularStream(_rootEntry);
    }

    private byte[] ReadSector(uint sectorId)
    {
        var offset = checked(((long)sectorId + 1) * _sectorSize);
        if (offset < _sectorSize || offset + _sectorSize > _stream.Length)
        {
            throw new InvalidDataException("An OLE sector is outside the file.");
        }

        var bytes = new byte[_sectorSize];
        _stream.Position = offset;
        _stream.ReadExactly(bytes);
        return bytes;
    }

    private static bool IsRegularSector(uint sectorId, int sectorCount) =>
        sectorId < sectorCount;

    public void Dispose()
    {
        _stream.Dispose();
    }

    private sealed record DirectoryEntry(
        int Id,
        string Name,
        byte Type,
        int LeftSibling,
        int RightSibling,
        int Child,
        uint StartSector,
        long StreamSize);

    private readonly record struct DirectoryTraversal(
        int EntryId,
        DirectorySearchScope Scope,
        DirectoryTraversalPhase Phase);

    private enum DirectoryTraversalPhase
    {
        Enter,
        Visit
    }

    private enum DirectorySearchScope
    {
        Root,
        DataSpaces,
        DataSpaceInfo
    }
}
