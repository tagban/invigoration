using System.Buffers.Binary;
using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>Progress of a running transfer, reported as bytes move. <paramref name="Total"/> is 0 when the server didn't say how big it is.</summary>
public readonly record struct HotlineTransferProgress(long Transferred, long Total)
{
    public double? Fraction => Total > 0 ? Math.Clamp((double)Transferred / Total, 0, 1) : null;
}

/// <summary>
/// Moves the actual bytes of a file. Hotline does this on a second, short-lived connection to
/// port+1 rather than over the transaction connection: the main connection asks for the file and
/// gets back a reference number, and that number is the ticket this connection presents. One
/// connection per transfer, closed when it finishes.
///
/// The payload is a "flattened file object" — a small header, then forks. The one that matters is
/// DATA (the file's actual contents); INFO carries the classic Mac type/creator codes and the
/// name, and MACR a resource fork that modern files don't have. Downloads keep only DATA and skip
/// the rest, which is why a download can't simply be piped to disk byte-for-byte.
/// </summary>
public static class HotlineFileTransfer
{
    private static readonly byte[] TransferMagic = "HTXF"u8.ToArray();
    private static readonly byte[] FlatFileMagic = "FILP"u8.ToArray();
    private static readonly byte[] InfoFork = "INFO"u8.ToArray();
    private static readonly byte[] DataFork = "DATA"u8.ToArray();

    /// <summary>The transfer port is always the transaction port plus one.</summary>
    public static int TransferPort(int serverPort) => serverPort + 1;

    /// <summary>
    /// Downloads one file's contents into <paramref name="destination"/>, given the reference
    /// number the server handed back from a DownloadFile request. Returns the number of bytes
    /// written — the DATA fork's size, which is smaller than the transfer size the server quoted
    /// (that includes the headers and every other fork).
    /// </summary>
    public static async Task<long> DownloadAsync(
        string host,
        int serverPort,
        uint referenceNumber,
        Stream destination,
        IProgress<HotlineTransferProgress>? progress = null,
        CancellationToken ct = default)
    {
        using var client = new TcpClient();
        await client.ConnectAsync(host, TransferPort(serverPort), ct).ConfigureAwait(false);
        await using var stream = client.GetStream();

        await stream.WriteAsync(TransferHeader(referenceNumber, dataSize: 0), ct).ConfigureAwait(false);
        return await ReadFlattenedFileAsync(stream, destination, expectedTotal: 0, progress, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads one flattened file object off an already-open transfer stream and writes its DATA
    /// fork to <paramref name="destination"/>, returning how many bytes that was. Shared by a
    /// single-file download and by each item of a folder download, which differ only in how the
    /// connection got here.
    /// </summary>
    /// <param name="expectedTotal">The flattened size when the caller was told it up front (folder items are), else 0 — used only to make progress reporting meaningful.</param>
    internal static async Task<long> ReadFlattenedFileAsync(
        Stream stream,
        Stream destination,
        uint expectedTotal,
        IProgress<HotlineTransferProgress>? progress,
        CancellationToken ct)
    {
        var header = new byte[24];
        await stream.ReadExactlyAsync(header, ct).ConfigureAwait(false);
        if (!header.AsSpan(0, 4).SequenceEqual(FlatFileMagic))
        {
            throw new InvalidDataException("The server didn't send a Hotline file — the transfer may have expired or been denied.");
        }

        var forkCount = BinaryPrimitives.ReadUInt16BigEndian(header.AsSpan(22));
        long written = 0;

        for (var i = 0; i < forkCount; i++)
        {
            var forkHeader = new byte[16];
            await stream.ReadExactlyAsync(forkHeader, ct).ConfigureAwait(false);
            var forkSize = BinaryPrimitives.ReadUInt32BigEndian(forkHeader.AsSpan(12));

            if (forkHeader.AsSpan(0, 4).SequenceEqual(DataFork))
            {
                written = await CopyExactlyAsync(stream, destination, forkSize, progress, ct).ConfigureAwait(false);
            }
            else
            {
                // INFO, MACR, or anything a newer server adds: read past it so the next fork header
                // lands where it should. Skipping by seeking isn't an option on a socket.
                await CopyExactlyAsync(stream, Stream.Null, forkSize, progress: null, ct).ConfigureAwait(false);
            }
        }

        await destination.FlushAsync(ct).ConfigureAwait(false);
        return written;
    }

    /// <summary>
    /// Uploads <paramref name="source"/> as a flattened file, given the reference number from an
    /// UploadFile request. The INFO fork names the file; the type and creator codes are the
    /// generic ones a modern (non-Mac) file gets, since there's nothing better to claim.
    /// </summary>
    public static async Task UploadAsync(
        string host,
        int serverPort,
        uint referenceNumber,
        string fileName,
        Stream source,
        long length,
        IProgress<HotlineTransferProgress>? progress = null,
        CancellationToken ct = default)
    {
        var info = BuildInfoFork(fileName);
        // Flat-file header + two fork headers + both forks — the size the server is told to expect.
        var totalSize = 24 + 16 + info.Length + 16 + length;

        using var client = new TcpClient();
        await client.ConnectAsync(host, TransferPort(serverPort), ct).ConfigureAwait(false);
        await using var stream = client.GetStream();

        await stream.WriteAsync(TransferHeader(referenceNumber, (uint)totalSize), ct).ConfigureAwait(false);
        await WriteFlattenedFileAsync(stream, fileName, source, length, progress, ct, withLengthPrefix: false).ConfigureAwait(false);
    }

    /// <summary>Total bytes a flattened file takes on the wire, so a folder upload can tell the server up front how much is coming.</summary>
    public static long FlattenedSize(string fileName, long length) => 24 + 16 + BuildInfoFork(fileName).Length + 16 + length;

    /// <summary>
    /// Writes one flattened file object to an already-open transfer stream. Shared by a single-file
    /// upload and by each file of a folder upload — the only difference is that a folder item is
    /// prefixed with its own 4-byte length, which is what <paramref name="withLengthPrefix"/>
    /// controls.
    /// </summary>
    internal static async Task WriteFlattenedFileAsync(
        Stream stream,
        string fileName,
        Stream source,
        long length,
        IProgress<HotlineTransferProgress>? progress,
        CancellationToken ct,
        bool withLengthPrefix = true)
    {
        var info = BuildInfoFork(fileName);

        if (withLengthPrefix)
        {
            var size = new byte[4];
            BinaryPrimitives.WriteUInt32BigEndian(size, (uint)(24 + 16 + info.Length + 16 + length));
            await stream.WriteAsync(size, ct).ConfigureAwait(false);
        }

        var header = new byte[24];
        FlatFileMagic.CopyTo(header, 0);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(4), 1); // version
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(22), 2); // INFO + DATA
        await stream.WriteAsync(header, ct).ConfigureAwait(false);

        await stream.WriteAsync(ForkHeader(InfoFork, (uint)info.Length), ct).ConfigureAwait(false);
        await stream.WriteAsync(info, ct).ConfigureAwait(false);
        await stream.WriteAsync(ForkHeader(DataFork, (uint)length), ct).ConfigureAwait(false);
        await CopyExactlyAsync(source, stream, (uint)length, progress, ct).ConfigureAwait(false);
        await stream.FlushAsync(ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Fetches the server banner over its own transfer connection. Unlike a file, a banner is the
    /// raw image bytes with no flattened-file wrapper around them — so this reads the stream
    /// straight through rather than looking for forks.
    ///
    /// The transfer type is 2 (a banner), where a file is 0 and a folder is 1.
    /// </summary>
    public static async Task<byte[]?> DownloadBannerAsync(
        string host,
        int serverPort,
        uint referenceNumber,
        uint expectedSize,
        CancellationToken ct = default)
    {
        try
        {
            using var client = new TcpClient();
            await client.ConnectAsync(host, TransferPort(serverPort), ct).ConfigureAwait(false);
            await using var stream = client.GetStream();

            var header = new byte[16];
            TransferMagic.CopyTo(header, 0);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), referenceNumber);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(8), 0);
            BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(12), 2); // type: banner
            await stream.WriteAsync(header, ct).ConfigureAwait(false);

            using var banner = new MemoryStream();
            if (expectedSize > 0)
            {
                await CopyExactlyAsync(stream, banner, expectedSize, progress: null, ct).ConfigureAwait(false);
            }
            else
            {
                // A server that didn't say how big it is sends until it closes.
                await stream.CopyToAsync(banner, ct).ConfigureAwait(false);
            }

            return banner.Length > 0 ? banner.ToArray() : null;
        }
        catch (Exception ex) when (ex is IOException or SocketException or EndOfStreamException)
        {
            // A banner is decoration — never worth failing a login over.
            return null;
        }
    }

    /// <summary>"HTXF", the reference number, how many bytes are about to be sent (0 when receiving), then four reserved bytes.</summary>
    private static byte[] TransferHeader(uint referenceNumber, uint dataSize)
    {
        var header = new byte[16];
        TransferMagic.CopyTo(header, 0);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), referenceNumber);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(8), dataSize);
        return header;
    }

    private static byte[] ForkHeader(byte[] forkType, uint size)
    {
        var header = new byte[16];
        forkType.CopyTo(header, 0);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(12), size);
        return header;
    }

    /// <summary>
    /// The INFO fork: platform and type/creator codes, timestamps, then the name and a comment.
    /// "AMAC"/"????" are what a non-Mac client has to claim — there's no real type to report for a
    /// file that came from a modern filesystem.
    /// </summary>
    private static byte[] BuildInfoFork(string fileName)
    {
        var name = Encoding.UTF8.GetBytes(fileName);
        var fork = new byte[72 + name.Length + 2];

        "AMAC"u8.CopyTo(fork.AsSpan(0));
        "????"u8.CopyTo(fork.AsSpan(4));  // type
        "????"u8.CopyTo(fork.AsSpan(8));  // creator
        // Bytes 12-71 are flags, the platform flags, and the created/modified dates. Zeroes are
        // accepted by every server tested against; a real date here would need Hotline's own
        // 8-byte date form, which servers don't use for anything on upload.
        BinaryPrimitives.WriteUInt16BigEndian(fork.AsSpan(70), (ushort)name.Length);
        name.CopyTo(fork.AsSpan(72));
        return fork;
    }

    /// <summary>Copies exactly <paramref name="count"/> bytes, reporting progress. Throws if the connection ends early, so a truncated download fails loudly instead of leaving a corrupt file.</summary>
    private static async Task<long> CopyExactlyAsync(
        Stream source,
        Stream destination,
        uint count,
        IProgress<HotlineTransferProgress>? progress,
        CancellationToken ct)
    {
        var buffer = new byte[81920];
        long remaining = count;
        long done = 0;

        while (remaining > 0)
        {
            var read = await source.ReadAsync(buffer.AsMemory(0, (int)Math.Min(buffer.Length, remaining)), ct).ConfigureAwait(false);
            if (read == 0)
            {
                throw new EndOfStreamException($"The transfer ended after {done} of {count} bytes.");
            }

            await destination.WriteAsync(buffer.AsMemory(0, read), ct).ConfigureAwait(false);
            remaining -= read;
            done += read;
            progress?.Report(new HotlineTransferProgress(done, count));
        }

        return done;
    }
}
