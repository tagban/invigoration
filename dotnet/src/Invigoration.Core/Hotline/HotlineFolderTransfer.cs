using System.Buffers.Binary;
using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>One item as it arrives during a folder download, so callers can report progress per file.</summary>
public readonly record struct HotlineFolderItem(IReadOnlyList<string> RelativePath, bool IsFolder, long Size)
{
    public string Name => RelativePath.Count > 0 ? RelativePath[^1] : "";
}

/// <summary>
/// Downloading a whole folder. Unlike a single file — which is just "connect and read" — a folder
/// transfer is a conversation: the server describes one item, the client says what to do with it
/// (send it, resume it, or skip to the next), and that repeats until every item is accounted for.
///
/// The transfer connection is the same port+1 socket a file download uses, but its header carries
/// two extra fields (a type of 1, and an opening folder action) — see
/// <see cref="HotlineFileTransfer"/> for the plain-file version.
///
/// Per github.com/fogWraith/Hotline/blob/main/Docs/Protocol/Hotline.md.
/// </summary>
public static class HotlineFolderTransfer
{
    /// <summary>What the client tells the server to do with the item it just described.</summary>
    public enum FolderAction : ushort
    {
        SendFile = 1,
        ResumeFile = 2,
        NextFile = 3,
    }

    /// <summary>A folder item's type code, as sent in the item header. Folders are "fldr", same as in a directory listing.</summary>
    private const string FolderTypeCode = HotlineFileEntry.FolderType;

    /// <summary>
    /// Downloads a folder and everything under it into <paramref name="destinationDirectory"/>,
    /// recreating the server's own subfolder structure. Returns how many files were written.
    ///
    /// <paramref name="itemCount"/> comes from the DownloadFolder reply (Folder Item Count, 220) —
    /// the server doesn't otherwise say when it's finished, so the loop counts items rather than
    /// waiting for an end marker.
    /// </summary>
    public static async Task<int> DownloadAsync(
        string host,
        int serverPort,
        uint referenceNumber,
        int itemCount,
        string destinationDirectory,
        IProgress<HotlineFolderItem>? itemProgress = null,
        IProgress<HotlineTransferProgress>? byteProgress = null,
        CancellationToken ct = default)
    {
        using var client = new TcpClient();
        await client.ConnectAsync(host, HotlineFileTransfer.TransferPort(serverPort), ct).ConfigureAwait(false);
        await using var stream = client.GetStream();

        await stream.WriteAsync(FolderTransferHeader(referenceNumber), ct).ConfigureAwait(false);

        var filesWritten = 0;
        for (var i = 0; i < itemCount; i++)
        {
            var item = await ReadItemHeaderAsync(stream, ct).ConfigureAwait(false);
            if (item is null)
            {
                break;
            }

            var (relativePath, isFolder) = item.Value;
            var localPath = Path.Combine(destinationDirectory, Path.Combine([.. SafePath(relativePath)]));

            if (isFolder)
            {
                Directory.CreateDirectory(localPath);
                itemProgress?.Report(new HotlineFolderItem(relativePath, true, 0));
                await SendActionAsync(stream, FolderAction.NextFile, ct).ConfigureAwait(false);
                continue;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(localPath) ?? destinationDirectory);
            await SendActionAsync(stream, FolderAction.SendFile, ct).ConfigureAwait(false);

            // The server answers a Send File with the flattened file's length, then the file itself.
            var sizeBytes = new byte[4];
            await stream.ReadExactlyAsync(sizeBytes, ct).ConfigureAwait(false);
            var flattenedSize = BinaryPrimitives.ReadUInt32BigEndian(sizeBytes);

            var partial = localPath + ".part";
            long written;
            await using (var file = File.Create(partial))
            {
                written = await HotlineFileTransfer.ReadFlattenedFileAsync(stream, file, flattenedSize, byteProgress, ct)
                    .ConfigureAwait(false);
            }

            File.Move(partial, localPath, overwrite: true);
            filesWritten++;
            itemProgress?.Report(new HotlineFolderItem(relativePath, false, written));
        }

        return filesWritten;
    }

    /// <summary>
    /// The opening header: the same 16 bytes a file transfer sends, but with the type set to 1 and
    /// a first folder action ("next file") appended — that action is what starts the item loop.
    /// </summary>
    private static byte[] FolderTransferHeader(uint referenceNumber)
    {
        var header = new byte[18];
        "HTXF"u8.CopyTo(header.AsSpan(0));
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), referenceNumber);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(8), 0); // data size: nothing being sent up
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(12), 1); // type: folder transfer
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(14), 0); // reserved
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(16), (ushort)FolderAction.NextFile);
        return header;
    }

    private static async Task SendActionAsync(Stream stream, FolderAction action, CancellationToken ct)
    {
        var bytes = new byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, (ushort)action);
        await stream.WriteAsync(bytes, ct).ConfigureAwait(false);
        await stream.FlushAsync(ct).ConfigureAwait(false);
    }

    /// <summary>
    /// One item header: a length, a type code, then the item's path relative to the folder being
    /// downloaded (the same packed path encoding requests use). Null when the connection ends,
    /// which is how a server that sends fewer items than it promised is handled.
    /// </summary>
    private static async Task<(IReadOnlyList<string> Path, bool IsFolder)?> ReadItemHeaderAsync(Stream stream, CancellationToken ct)
    {
        var prefix = new byte[4];
        try
        {
            await stream.ReadExactlyAsync(prefix, ct).ConfigureAwait(false);
        }
        catch (EndOfStreamException)
        {
            return null;
        }

        var headerSize = BinaryPrimitives.ReadUInt16BigEndian(prefix);
        var typeCode = BinaryPrimitives.ReadUInt16BigEndian(prefix.AsSpan(2));

        // headerSize counts the type field and the path that follows it; the two bytes of the size
        // field itself are not included.
        var remaining = headerSize >= 2 ? headerSize - 2 : 0;
        var pathBytes = new byte[remaining];
        if (remaining > 0)
        {
            await stream.ReadExactlyAsync(pathBytes, ct).ConfigureAwait(false);
        }

        // The type is sent as a 2-byte code rather than the 4-character "fldr" a listing uses; a
        // real server sends 1 for a folder here. Both spellings are accepted since servers differ.
        var isFolder = typeCode == 1 ||
            (pathBytes.Length >= 4 && Encoding.ASCII.GetString(pathBytes, 0, 4) == FolderTypeCode);

        return (HotlineNewsPath.Decode(pathBytes), isFolder);
    }

    /// <summary>
    /// Strips anything that could write outside the chosen folder — a server is not trusted to
    /// send well-behaved names, and "../../etc" in a path component would otherwise land wherever
    /// it liked. Empty or dot-only components are dropped.
    /// </summary>
    public static IEnumerable<string> SafePath(IReadOnlyList<string> components)
    {
        foreach (var component in components)
        {
            var cleaned = component.Trim();
            if (cleaned is "" or "." or "..")
            {
                continue;
            }

            foreach (var invalid in Path.GetInvalidFileNameChars())
            {
                cleaned = cleaned.Replace(invalid, '_');
            }

            if (cleaned.Length > 0)
            {
                yield return cleaned;
            }
        }
    }
}
