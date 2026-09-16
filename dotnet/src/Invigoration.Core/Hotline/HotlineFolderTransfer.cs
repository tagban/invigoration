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

    /// <summary>Everything in a folder, flattened to the list the server will be walked through — folders first so they exist before the files that go in them.</summary>
    public static IReadOnlyList<(IReadOnlyList<string> Path, bool IsFolder, long Size)> Enumerate(string localDirectory)
    {
        var root = new DirectoryInfo(localDirectory);
        var items = new List<(IReadOnlyList<string> Path, bool IsFolder, long Size)>();

        void Walk(DirectoryInfo directory, List<string> prefix)
        {
            foreach (var sub in directory.GetDirectories().OrderBy(d => d.Name, StringComparer.OrdinalIgnoreCase))
            {
                List<string> path = [.. prefix, sub.Name];
                items.Add((path, true, 0));
                Walk(sub, path);
            }

            foreach (var file in directory.GetFiles().OrderBy(f => f.Name, StringComparer.OrdinalIgnoreCase))
            {
                items.Add(([.. prefix, file.Name], false, file.Length));
            }
        }

        Walk(root, []);
        return items;
    }

    /// <summary>
    /// Uploads a folder and everything under it. The mirror of the download: the server drives,
    /// asking for the next item and then deciding whether it wants that item's contents, and this
    /// answers until every item has been offered.
    ///
    /// Resume (action 2) is answered by sending the whole file from the start — the resume block
    /// is read and discarded. Partial resumes aren't implemented; re-sending is correct, just not
    /// the most economical thing a client could do.
    /// </summary>
    public static async Task<int> UploadAsync(
        string host,
        int serverPort,
        uint referenceNumber,
        string localDirectory,
        IProgress<HotlineFolderItem>? itemProgress = null,
        IProgress<HotlineTransferProgress>? byteProgress = null,
        CancellationToken ct = default)
    {
        var items = Enumerate(localDirectory);
        using var client = new TcpClient();
        await client.ConnectAsync(host, HotlineFileTransfer.TransferPort(serverPort), ct).ConfigureAwait(false);
        await using var stream = client.GetStream();

        // Same 16 bytes a file upload sends, with the type marking this a folder. Unlike a
        // download there's no opening action here — the server speaks first.
        var header = new byte[16];
        "HTXF"u8.CopyTo(header.AsSpan(0));
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), referenceNumber);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(8), 0);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(12), 1);
        await stream.WriteAsync(header, ct).ConfigureAwait(false);

        var sent = 0;
        var next = 0;

        while (next < items.Count)
        {
            if (await ReadActionAsync(stream, ct).ConfigureAwait(false) is not { } action)
            {
                break;
            }

            if (action != FolderAction.NextFile)
            {
                // Only "next file" is meaningful before an item has been offered.
                continue;
            }

            var (path, isFolder, size) = items[next++];
            await stream.WriteAsync(ItemDescriptor(path, isFolder), ct).ConfigureAwait(false);
            await stream.FlushAsync(ct).ConfigureAwait(false);

            if (isFolder)
            {
                itemProgress?.Report(new HotlineFolderItem(path, true, 0));
                continue;
            }

            if (await ReadActionAsync(stream, ct).ConfigureAwait(false) is not { } response)
            {
                break;
            }

            if (response == FolderAction.ResumeFile)
            {
                // Read past the resume block, then send the whole file anyway.
                var resumeSize = new byte[2];
                await stream.ReadExactlyAsync(resumeSize, ct).ConfigureAwait(false);
                var length = BinaryPrimitives.ReadUInt16BigEndian(resumeSize);
                if (length > 0)
                {
                    await stream.ReadExactlyAsync(new byte[length], ct).ConfigureAwait(false);
                }
            }
            else if (response != FolderAction.SendFile)
            {
                // The server skipped this one.
                continue;
            }

            var localPath = Path.Combine(localDirectory, Path.Combine([.. path]));
            await using (var file = File.OpenRead(localPath))
            {
                await HotlineFileTransfer.WriteFlattenedFileAsync(stream, path[^1], file, size, byteProgress, ct)
                    .ConfigureAwait(false);
            }

            sent++;
            itemProgress?.Report(new HotlineFolderItem(path, false, size));
        }

        await stream.FlushAsync(ct).ConfigureAwait(false);
        return sent;
    }

    private static async Task<FolderAction?> ReadActionAsync(Stream stream, CancellationToken ct)
    {
        var action = new byte[2];
        try
        {
            await stream.ReadExactlyAsync(action, ct).ConfigureAwait(false);
        }
        catch (EndOfStreamException)
        {
            return null;
        }

        return (FolderAction)BinaryPrimitives.ReadUInt16BigEndian(action);
    }

    /// <summary>
    /// One item offered to the server: a length, whether it's a folder, then its path relative to
    /// the folder being uploaded. Note this is NOT the same packing a request's FilePath uses —
    /// the is-folder flag sits between the length and the component count.
    /// </summary>
    private static byte[] ItemDescriptor(IReadOnlyList<string> path, bool isFolder)
    {
        var names = path.Select(Encoding.UTF8.GetBytes).ToArray();
        var bodySize = 2 + 2 + names.Sum(n => 3 + n.Length);
        var buffer = new byte[2 + bodySize];

        BinaryPrimitives.WriteUInt16BigEndian(buffer, (ushort)bodySize);
        BinaryPrimitives.WriteUInt16BigEndian(buffer.AsSpan(2), isFolder ? (ushort)1 : (ushort)0);
        BinaryPrimitives.WriteUInt16BigEndian(buffer.AsSpan(4), (ushort)names.Length);

        var offset = 6;
        foreach (var name in names)
        {
            buffer[offset] = 0;
            buffer[offset + 1] = 0;
            buffer[offset + 2] = (byte)Math.Min(name.Length, byte.MaxValue);
            name.AsSpan(0, buffer[offset + 2]).CopyTo(buffer.AsSpan(offset + 3));
            offset += 3 + buffer[offset + 2];
        }

        return buffer;
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
