using System.Buffers.Binary;
using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>One server as listed by a Hotline tracker (the public/community-run directory service — hltracker.com is the well-known default, per the user's own request).</summary>
public sealed record HotlineTrackerServerEntry(string Address, ushort Port, ushort UserCount, string Name, string Description);

/// <summary>
/// Queries a Hotline tracker (HTRK protocol, TCP, default port 5498) for its list of registered
/// servers — this is deliberately the very first thing a user of this feature sees ("the primary
/// window starts with just the tracker"), so it's a plain static query rather than a stateful
/// connection like HotlineTransactionClient. Ported from Hotline-Navigator's tracker.rs.
/// </summary>
public static class HotlineTrackerClient
{
    public static async Task<IReadOnlyList<HotlineTrackerServerEntry>> QueryAsync(string host, int port = HotlineConstants.DefaultTrackerPort, CancellationToken ct = default)
    {
        using var client = new TcpClient();
        await client.ConnectAsync(host, port, ct).ConfigureAwait(false);
        using var stream = client.GetStream();

        var request = new byte[6];
        HotlineConstants.TrackerMagic.CopyTo(request, 0);
        BinaryPrimitives.WriteUInt16BigEndian(request.AsSpan(4), HotlineConstants.TrackerVersion);
        await stream.WriteAsync(request, ct).ConfigureAwait(false);

        var handshakeReply = await ReadExactAsync(stream, 6, ct).ConfigureAwait(false);
        if (handshakeReply is null || !handshakeReply.AsSpan(0, 4).SequenceEqual(HotlineConstants.TrackerMagic))
        {
            return [];
        }

        var results = new List<HotlineTrackerServerEntry>();

        // The tracker leaves the TCP connection open and idle once it's done (confirmed live
        // against the real hltracker.com), so "read batches until EOF" hangs forever — but a
        // large server list DOES arrive as more than one batch, so "read exactly one batch and
        // stop" (this method's second, still-wrong attempt) silently truncated the list. The
        // real shape (confirmed against Hotline-Navigator's actual read loop): the FIRST batch
        // header's server_count field is the TOTAL entry count across every batch, not this
        // batch's own count — that's server_count2. Keep reading batch headers, each contributing
        // its own server_count2 entries, until the running total reaches that first-seen total (a
        // capped iteration count guards against a malformed/hostile tracker never reaching it).
        var totalExpected = -1;
        var parsed = 0;
        var batchCount = 0;
        while ((totalExpected < 0 || parsed < totalExpected) && batchCount++ < 100)
        {
            var batchHeader = await ReadExactAsync(stream, 8, ct).ConfigureAwait(false);
            if (batchHeader is null)
            {
                break;
            }

            if (totalExpected < 0)
            {
                totalExpected = BinaryPrimitives.ReadUInt16BigEndian(batchHeader.AsSpan(4));
            }

            var countInBatch = BinaryPrimitives.ReadUInt16BigEndian(batchHeader.AsSpan(6));
            for (var i = 0; i < countInBatch; i++)
            {
                var entry = await ReadServerEntryAsync(stream, ct).ConfigureAwait(false);
                if (entry is null)
                {
                    return results;
                }

                results.Add(entry);
                parsed++;
            }
        }

        return results;
    }

    private static async Task<HotlineTrackerServerEntry?> ReadServerEntryAsync(NetworkStream stream, CancellationToken ct)
    {
        var fixedPart = await ReadExactAsync(stream, 10, ct).ConfigureAwait(false); // ip(4) + port(2) + users(2) + unused(2)
        if (fixedPart is null)
        {
            return null;
        }

        var address = $"{fixedPart[0]}.{fixedPart[1]}.{fixedPart[2]}.{fixedPart[3]}";
        var port = BinaryPrimitives.ReadUInt16BigEndian(fixedPart.AsSpan(4));
        var users = BinaryPrimitives.ReadUInt16BigEndian(fixedPart.AsSpan(6));

        var name = await ReadPascalStringAsync(stream, ct).ConfigureAwait(false);
        var description = await ReadPascalStringAsync(stream, ct).ConfigureAwait(false);
        if (name is null || description is null)
        {
            return null;
        }

        return new HotlineTrackerServerEntry(address, port, users, name, description);
    }

    /// <summary>
    /// Real Hotline servers/trackers encode these Pascal strings as Mac OS Roman, not UTF-8 or
    /// Latin-1 — decoded here as Latin-1 anyway, a known, accepted simplification (same one
    /// HotlineField.AsString makes): the two encodings agree on plain ASCII, which covers the
    /// overwhelming majority of real server names/descriptions, and pulling in a full Mac Roman
    /// code page table for the rare non-ASCII one is real extra complexity for no benefit to this
    /// app's actual users.
    /// </summary>
    private static async Task<string?> ReadPascalStringAsync(NetworkStream stream, CancellationToken ct)
    {
        var lengthByte = await ReadExactAsync(stream, 1, ct).ConfigureAwait(false);
        if (lengthByte is null)
        {
            return null;
        }

        var length = lengthByte[0];
        if (length == 0)
        {
            return "";
        }

        var data = await ReadExactAsync(stream, length, ct).ConfigureAwait(false);
        return data is null ? null : Encoding.Latin1.GetString(data);
    }

    private static async Task<byte[]?> ReadExactAsync(NetworkStream stream, int count, CancellationToken ct)
    {
        var buffer = new byte[count];
        var offset = 0;
        while (offset < count)
        {
            int read;
            try
            {
                read = await stream.ReadAsync(buffer.AsMemory(offset, count - offset), ct).ConfigureAwait(false);
            }
            catch (IOException)
            {
                return null;
            }

            if (read == 0)
            {
                return null; // remote closed mid-read (or cleanly, if offset is still 0 — either way, nothing more to parse)
            }

            offset += read;
        }

        return buffer;
    }
}
