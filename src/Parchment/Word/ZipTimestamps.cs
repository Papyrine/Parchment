/// <summary>
/// Pins every zip entry timestamp in a finished package to a stable date, in place.
/// </summary>
/// <remarks>
/// Zip entry timestamps come from the wall clock at the format's 2-second resolution. Entries
/// cloned from the registration snapshot keep their original stamps through a save, but a part
/// added during the render — settings, a numbering part for a list, an image — is stamped with
/// now, and <c>[Content_Types].xml</c> is re-stamped on every save. So the byte-identical
/// guarantee held only when two renders landed inside the same 2-second quantum.
/// <para>
/// This rewrites only the four timestamp bytes in each local file header and central directory
/// record. Nothing else moves: no entry reordering, no recompression, no content rewrite.
/// <c>DeterministicIoPackaging.DeterministicPackage</c> was tried first and is deliberately not
/// used — it is built for snapshot normalization and rewrites entry content (png chunks, core
/// properties), which loses document content the render put there on purpose.
/// </para>
/// </remarks>
static class ZipTimestamps
{
    // 2020-01-01 00:00:00, matching DeterministicIoPackaging's StableDate so packages normalized
    // by either read the same. DOS date packs as year-1980 << 9 | month << 5 | day.
    internal static readonly DateTime StableDate = new(2020, 1, 1);
    const ushort stableDosDate = (2020 - 1980) << 9 | 1 << 5 | 1;
    const ushort stableDosTime = 0;

    const uint centralDirectorySignature = 0x02014b50;
    const uint localHeaderSignature = 0x04034b50;
    const uint endOfCentralDirectorySignature = 0x06054b50;

    public static void Pin(MemoryStream stream)
    {
        if (!stream.TryGetBuffer(out var segment) ||
            segment.Array == null)
        {
            throw new InvalidOperationException("Zip timestamp pinning needs an exposable MemoryStream buffer.");
        }

        var buffer = segment.Array;
        var length = segment.Offset + (int) stream.Length;

        var endOfCentralDirectory = FindEndOfCentralDirectory(buffer, segment.Offset, length);
        var entryCount = ReadUInt16(buffer, endOfCentralDirectory + 10);
        var offset = segment.Offset + checked((int) ReadUInt32(buffer, endOfCentralDirectory + 16));

        // 0xFFFF entries / 0xFFFFFFFF offset are the zip64 escape values. A docx never gets close,
        // and silently skipping entries would silently break the guarantee.
        if (entryCount == ushort.MaxValue)
        {
            throw new InvalidOperationException("Zip64 archives are not supported.");
        }

        for (var i = 0; i < entryCount; i++)
        {
            if (ReadUInt32(buffer, offset) != centralDirectorySignature)
            {
                throw new InvalidOperationException($"Malformed central directory at offset {offset}.");
            }

            WriteUInt16(buffer, offset + 12, stableDosTime);
            WriteUInt16(buffer, offset + 14, stableDosDate);

            var localHeader = segment.Offset + checked((int) ReadUInt32(buffer, offset + 42));
            if (ReadUInt32(buffer, localHeader) != localHeaderSignature)
            {
                throw new InvalidOperationException($"Malformed local header at offset {localHeader}.");
            }

            WriteUInt16(buffer, localHeader + 10, stableDosTime);
            WriteUInt16(buffer, localHeader + 12, stableDosDate);

            var nameLength = ReadUInt16(buffer, offset + 28);
            var extraLength = ReadUInt16(buffer, offset + 30);
            var commentLength = ReadUInt16(buffer, offset + 32);
            offset += 46 + nameLength + extraLength + commentLength;
        }
    }

    static int FindEndOfCentralDirectory(byte[] buffer, int start, int length)
    {
        // The record is 22 bytes plus an optional trailing comment of up to 64k. These packages
        // carry no comment, so the first candidate is the match; the scan is bounded anyway.
        var lowest = Math.Max(start, length - 22 - ushort.MaxValue);
        for (var i = length - 22; i >= lowest; i--)
        {
            if (ReadUInt32(buffer, i) == endOfCentralDirectorySignature)
            {
                return i;
            }
        }

        throw new InvalidOperationException("End-of-central-directory record not found.");
    }

    static ushort ReadUInt16(byte[] buffer, int index) =>
        (ushort) (buffer[index] | buffer[index + 1] << 8);

    static uint ReadUInt32(byte[] buffer, int index) =>
        (uint) (buffer[index] | buffer[index + 1] << 8 | buffer[index + 2] << 16 | buffer[index + 3] << 24);

    static void WriteUInt16(byte[] buffer, int index, ushort value)
    {
        buffer[index] = (byte) value;
        buffer[index + 1] = (byte) (value >> 8);
    }
}
