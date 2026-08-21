using System.Buffers.Binary;
using System.Text;

namespace Sbroenne.ExcelMcp.Tests.Helpers;

public static class OleDataSpaceTestFile
{
    private const int SectorSize = 512;
    private const int StreamSize = 4096;
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint FreeSector = 0xFFFFFFFF;
    private const uint FatSector = 0xFFFFFFFD;

    public static string Write(
        string path,
        string dataSpaceName,
        string? definitionName = null)
    {
        File.WriteAllBytes(
            path,
            Build(dataSpaceName, definitionName ?? dataSpaceName));
        return path;
    }

    public static string WriteDeepDirectory(string path, int entryCount)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(entryCount);
        var directorySectorCount = (entryCount + 1 + 3) / 4;
        var fatSectorCount = 1;
        while (fatSectorCount * (SectorSize / sizeof(uint))
               < directorySectorCount + fatSectorCount)
        {
            fatSectorCount++;
        }

        if (fatSectorCount > 109)
        {
            throw new ArgumentOutOfRangeException(
                nameof(entryCount),
                "The deterministic fixture supports only header DIFAT entries.");
        }

        var bytes = new byte[
            SectorSize * (1 + directorySectorCount + fatSectorCount)];
        var header = bytes.AsSpan(0, SectorSize);
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 3);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteUInt32LittleEndian(header[44..], checked((uint)fatSectorCount));
        BinaryPrimitives.WriteUInt32LittleEndian(header[48..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header[60..], EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        for (var offset = 76; offset < SectorSize; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[offset..], FreeSector);
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(
                header[(76 + index * sizeof(uint))..],
                checked((uint)(directorySectorCount + index)));
        }

        var directory = bytes.AsSpan(
            SectorSize,
            directorySectorCount * SectorSize);
        WriteDirectoryEntry(directory, 0, "Root Entry", 5, child: 1);
        for (var index = 1; index <= entryCount; index++)
        {
            WriteDirectoryEntry(
                directory,
                index,
                $"Entry{index:D6}",
                1,
                leftSibling: index == entryCount ? -1 : index + 1);
        }

        var fatEntries = new uint[fatSectorCount * (SectorSize / sizeof(uint))];
        Array.Fill(fatEntries, FreeSector);
        for (var index = 0; index < directorySectorCount; index++)
        {
            fatEntries[index] = index == directorySectorCount - 1
                ? EndOfChain
                : checked((uint)(index + 1));
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            fatEntries[directorySectorCount + index] = FatSector;
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            var fatBytes = bytes.AsSpan(
                SectorSize * (1 + directorySectorCount + index),
                SectorSize);
            for (var entry = 0; entry < SectorSize / sizeof(uint); entry++)
            {
                BinaryPrimitives.WriteUInt32LittleEndian(
                    fatBytes[(entry * sizeof(uint))..],
                    fatEntries[index * (SectorSize / sizeof(uint)) + entry]);
            }
        }

        File.WriteAllBytes(path, bytes);
        return path;
    }

    public static string WriteDeepNestedStorages(string path, int entryCount)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(entryCount);
        var directorySectorCount = (entryCount + 1 + 3) / 4;
        var fatSectorCount = 1;
        while (fatSectorCount * (SectorSize / sizeof(uint))
               < directorySectorCount + fatSectorCount)
        {
            fatSectorCount++;
        }

        if (fatSectorCount > 109)
        {
            throw new ArgumentOutOfRangeException(
                nameof(entryCount),
                "The deterministic fixture supports only header DIFAT entries.");
        }

        var bytes = new byte[
            SectorSize * (1 + directorySectorCount + fatSectorCount)];
        var header = bytes.AsSpan(0, SectorSize);
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 3);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteUInt32LittleEndian(header[44..], checked((uint)fatSectorCount));
        BinaryPrimitives.WriteUInt32LittleEndian(header[48..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header[60..], EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        for (var offset = 76; offset < SectorSize; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[offset..], FreeSector);
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            BinaryPrimitives.WriteUInt32LittleEndian(
                header[(76 + index * sizeof(uint))..],
                checked((uint)(directorySectorCount + index)));
        }

        var directory = bytes.AsSpan(
            SectorSize,
            directorySectorCount * SectorSize);
        WriteDirectoryEntry(directory, 0, "Root Entry", 5, child: 1);
        for (var index = 1; index <= entryCount; index++)
        {
            WriteDirectoryEntry(
                directory,
                index,
                $"NestedStorage{index:D18}",
                1,
                child: index == entryCount ? -1 : index + 1);
        }

        var fatEntries = new uint[fatSectorCount * (SectorSize / sizeof(uint))];
        Array.Fill(fatEntries, FreeSector);
        for (var index = 0; index < directorySectorCount; index++)
        {
            fatEntries[index] = index == directorySectorCount - 1
                ? EndOfChain
                : checked((uint)(index + 1));
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            fatEntries[directorySectorCount + index] = FatSector;
        }
        for (var index = 0; index < fatSectorCount; index++)
        {
            var fatBytes = bytes.AsSpan(
                SectorSize * (1 + directorySectorCount + index),
                SectorSize);
            for (var entry = 0; entry < SectorSize / sizeof(uint); entry++)
            {
                BinaryPrimitives.WriteUInt32LittleEndian(
                    fatBytes[(entry * sizeof(uint))..],
                    fatEntries[index * (SectorSize / sizeof(uint)) + entry]);
            }
        }

        File.WriteAllBytes(path, bytes);
        return path;
    }

    public static string WriteOversizedVersion4RootMiniStream(string path)
    {
        const int version4SectorSize = 4096;
        var bytes = new byte[version4SectorSize * 4];
        var header = bytes.AsSpan(0, 512);
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 4);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 12);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteUInt32LittleEndian(header[40..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header[44..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header[48..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header[60..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header[64..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        for (var offset = 76; offset < 512; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[offset..], FreeSector);
        }
        BinaryPrimitives.WriteUInt32LittleEndian(header[76..], 2);

        var directory = bytes.AsSpan(version4SectorSize, version4SectorSize);
        WriteDirectoryEntry(
            directory,
            0,
            "Root Entry",
            5,
            child: 1,
            streamSize: 9_000_000_000_000);
        WriteDirectoryEntry(directory, 1, "\u0006DataSpaces", 1, child: 2);
        WriteDirectoryEntry(
            directory,
            2,
            "DataSpaceMap",
            2,
            startSector: 0,
            streamSize: 64);

        var fat = bytes.AsSpan(version4SectorSize * 3, version4SectorSize);
        for (var offset = 0; offset < fat.Length; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(fat[offset..], FreeSector);
        }
        BinaryPrimitives.WriteUInt32LittleEndian(fat, EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(fat[sizeof(uint)..], EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(fat[(2 * sizeof(uint))..], FatSector);

        File.WriteAllBytes(path, bytes);
        return path;
    }

    private static byte[] Build(string dataSpaceName, string definitionName)
    {
        const int directorySectorCount = 2;
        const int mapStartSector = directorySectorCount;
        const int definitionStartSector = mapStartSector + StreamSize / SectorSize;
        const int fatSectorIndex = definitionStartSector + StreamSize / SectorSize;
        var bytes = new byte[SectorSize * (fatSectorIndex + 2)];
        var header = bytes.AsSpan(0, SectorSize);

        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(header);
        BinaryPrimitives.WriteUInt16LittleEndian(header[24..], 0x003E);
        BinaryPrimitives.WriteUInt16LittleEndian(header[26..], 3);
        BinaryPrimitives.WriteUInt16LittleEndian(header[28..], 0xFFFE);
        BinaryPrimitives.WriteUInt16LittleEndian(header[30..], 9);
        BinaryPrimitives.WriteUInt16LittleEndian(header[32..], 6);
        BinaryPrimitives.WriteUInt32LittleEndian(header[44..], 1);
        BinaryPrimitives.WriteUInt32LittleEndian(header[48..], 0);
        BinaryPrimitives.WriteUInt32LittleEndian(header[56..], 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(header[60..], EndOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(header[68..], EndOfChain);
        for (var offset = 76; offset < SectorSize; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(header[offset..], FreeSector);
        }
        BinaryPrimitives.WriteUInt32LittleEndian(header[76..], (uint)fatSectorIndex);

        var directory = bytes.AsSpan(SectorSize, directorySectorCount * SectorSize);
        WriteDirectoryEntry(directory, 0, "Root Entry", 5, child: 1);
        WriteDirectoryEntry(directory, 1, "\u0006DataSpaces", 1, child: 2);
        WriteDirectoryEntry(directory, 2, "DataSpaceInfo", 1, rightSibling: 3, child: 4);
        WriteDirectoryEntry(
            directory,
            3,
            "DataSpaceMap",
            2,
            startSector: mapStartSector,
            streamSize: StreamSize);
        WriteDirectoryEntry(
            directory,
            4,
            definitionName,
            2,
            startSector: definitionStartSector,
            streamSize: StreamSize);

        BuildDataSpaceMap(dataSpaceName).CopyTo(Sector(bytes, mapStartSector));
        BuildDataSpaceDefinition().CopyTo(Sector(bytes, definitionStartSector));

        var fat = Sector(bytes, fatSectorIndex);
        for (var offset = 0; offset < fat.Length; offset += sizeof(uint))
        {
            BinaryPrimitives.WriteUInt32LittleEndian(fat[offset..], FreeSector);
        }
        WriteChain(fat, 0, directorySectorCount);
        WriteChain(fat, mapStartSector, StreamSize / SectorSize);
        WriteChain(fat, definitionStartSector, StreamSize / SectorSize);
        BinaryPrimitives.WriteUInt32LittleEndian(
            fat[(fatSectorIndex * sizeof(uint))..],
            FatSector);

        return bytes;
    }

    private static byte[] BuildDataSpaceMap(string dataSpaceName)
    {
        using var entryStream = new MemoryStream();
        using (var entry = new BinaryWriter(entryStream, Encoding.Unicode, leaveOpen: true))
        {
            entry.Write(1u);
            entry.Write(0u);
            WriteUnicodeString(entry, "EncryptedPackage");
            WriteUnicodeString(entry, dataSpaceName);
        }

        using var mapStream = new MemoryStream();
        using (var map = new BinaryWriter(mapStream, Encoding.Unicode, leaveOpen: true))
        {
            map.Write(8u);
            map.Write(1u);
            map.Write(checked((uint)(entryStream.Length + sizeof(uint))));
            map.Write(entryStream.ToArray());
        }

        return mapStream.ToArray();
    }

    private static byte[] BuildDataSpaceDefinition()
    {
        using var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, Encoding.Unicode, leaveOpen: true);
        writer.Write(8u);
        writer.Write(1u);
        WriteUnicodeString(writer, "DRMTransform");
        return stream.ToArray();
    }

    private static void WriteUnicodeString(BinaryWriter writer, string value)
    {
        var encoded = Encoding.Unicode.GetBytes(value);
        writer.Write(checked((uint)encoded.Length));
        writer.Write(encoded);
        var padding = (4 - encoded.Length % 4) % 4;
        writer.Write(new byte[padding]);
    }

    private static void WriteDirectoryEntry(
        Span<byte> directory,
        int index,
        string name,
        byte type,
        int leftSibling = -1,
        int rightSibling = -1,
        int child = -1,
        int startSector = -2,
        long streamSize = 0)
    {
        var entry = directory.Slice(index * 128, 128);
        var encodedName = Encoding.Unicode.GetBytes(name + "\0");
        encodedName.CopyTo(entry);
        BinaryPrimitives.WriteUInt16LittleEndian(entry[64..], checked((ushort)encodedName.Length));
        entry[66] = type;
        entry[67] = 1;
        BinaryPrimitives.WriteInt32LittleEndian(entry[68..], leftSibling);
        BinaryPrimitives.WriteInt32LittleEndian(entry[72..], rightSibling);
        BinaryPrimitives.WriteInt32LittleEndian(entry[76..], child);
        BinaryPrimitives.WriteInt32LittleEndian(entry[116..], startSector);
        BinaryPrimitives.WriteInt64LittleEndian(entry[120..], streamSize);
    }

    private static void WriteChain(Span<byte> fat, int startSector, int sectorCount)
    {
        for (var index = 0; index < sectorCount; index++)
        {
            var next = index == sectorCount - 1
                ? EndOfChain
                : checked((uint)(startSector + index + 1));
            BinaryPrimitives.WriteUInt32LittleEndian(
                fat[((startSector + index) * sizeof(uint))..],
                next);
        }
    }

    private static Span<byte> Sector(byte[] bytes, int sectorIndex) =>
        bytes.AsSpan(SectorSize * (sectorIndex + 1), SectorSize);
}
