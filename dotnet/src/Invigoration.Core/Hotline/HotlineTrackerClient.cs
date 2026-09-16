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
    /// <summary>Which protocol the last query actually used, and what the tracker offered — for showing the user whether they're on a modern tracker.</summary>
    public static ushort LastNegotiatedVersion { get; private set; }

    public static HotlineConstants.TrackerFeatures LastNegotiatedFeatures { get; private set; }

    /// <summary>
    /// Asks a tracker for its server list. Announces v3 and falls back on its answer: a v1/v2
    /// tracker ignores the extra bytes and replies with its own version, so offering the newer
    /// protocol costs nothing and gains IPv6 addresses, hostnames, and server-side search on a
    /// tracker that has them.
    /// </summary>
    /// <param name="search">Substring the tracker should filter on — only honoured by a v3 tracker that offers the query feature; ignored otherwise (the caller still filters locally).</param>
    /// <param name="limit">Maximum records to ask for, same caveat.</param>
    public static async Task<IReadOnlyList<HotlineTrackerServerEntry>> QueryAsync(
        string host,
        int port = HotlineConstants.DefaultTrackerPort,
        CancellationToken ct = default,
        string? search = null,
        ushort limit = 0)
    {
        // Ask for v3 first, then fall back to v1 on a BRAND-NEW connection — not by continuing on
        // the same socket. Confirmed live against hltracker.com (2026-09-15): it closes the
        // connection the moment it sees version 3 in the handshake, with or without the extra
        // feature bytes, contrary to the spec's claim that older trackers ignore what they don't
        // recognize. Same shape as the TRTP subversion fallback in HotlineTransactionClient, and
        // for the same reason: a rejected handshake leaves the socket unusable.
        if (await TryQueryAsync(host, port, HotlineConstants.TrackerVersion3, search, limit, ct).ConfigureAwait(false) is { } modern)
        {
            return modern;
        }

        return await TryQueryAsync(host, port, HotlineConstants.TrackerVersion, search, limit, ct).ConfigureAwait(false) ?? [];
    }

    /// <summary>One attempt at one protocol version. Null means the tracker refused this version — the caller should try an older one on a fresh connection — while an empty list means it answered and had nothing to list.</summary>
    private static async Task<IReadOnlyList<HotlineTrackerServerEntry>?> TryQueryAsync(
        string host,
        int port,
        ushort version,
        string? search,
        ushort limit,
        CancellationToken ct)
    {
        using var client = new TcpClient();
        await client.ConnectAsync(host, port, ct).ConfigureAwait(false);
        using var stream = client.GetStream();

        // v3 adds two feature bytes after the version; v1's handshake is just magic and version.
        var isVersion3 = version >= HotlineConstants.TrackerVersion3;
        var request = new byte[isVersion3 ? 8 : 6];
        HotlineConstants.TrackerMagic.CopyTo(request, 0);
        BinaryPrimitives.WriteUInt16BigEndian(request.AsSpan(4), version);
        if (isVersion3)
        {
            BinaryPrimitives.WriteUInt16BigEndian(request.AsSpan(6),
                (ushort)(HotlineConstants.TrackerFeatures.IPv6 | HotlineConstants.TrackerFeatures.Query));
        }

        await stream.WriteAsync(request, ct).ConfigureAwait(false);

        var handshakeReply = await ReadExactAsync(stream, 6, ct).ConfigureAwait(false);
        if (handshakeReply is null || !handshakeReply.AsSpan(0, 4).SequenceEqual(HotlineConstants.TrackerMagic))
        {
            // No reply at all, or something that isn't a tracker: refused.
            return null;
        }

        var negotiated = BinaryPrimitives.ReadUInt16BigEndian(handshakeReply.AsSpan(4));
        LastNegotiatedVersion = negotiated;
        LastNegotiatedFeatures = HotlineConstants.TrackerFeatures.None;

        if (negotiated >= HotlineConstants.TrackerVersion3)
        {
            // Only a v3 tracker sends the two feature bytes — reading them from a v1 tracker would
            // eat the first two bytes of its listing.
            var features = await ReadExactAsync(stream, 2, ct).ConfigureAwait(false);
            if (features is null)
            {
                return null;
            }

            LastNegotiatedFeatures = (HotlineConstants.TrackerFeatures)BinaryPrimitives.ReadUInt16BigEndian(features);
            return await ReadVersion3ListingAsync(stream, search, limit, ct).ConfigureAwait(false);
        }

        // A tracker that answered v1 to a v3 request is fine to keep reading from — it never saw a
        // version it objected to, it just spoke its own.

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

    /// <summary>
    /// The v3 listing: one request naming what we want, then one response carrying however many
    /// records matched. Unlike v1 — where records are fixed-shape and the batching is implicit —
    /// every v3 record is self-describing, so a tracker can add fields without breaking a client
    /// that doesn't know them (the trailing TLVs are read past here rather than parsed).
    /// </summary>
    private static async Task<IReadOnlyList<HotlineTrackerServerEntry>> ReadVersion3ListingAsync(
        NetworkStream stream,
        string? search,
        ushort limit,
        CancellationToken ct)
    {
        var canQuery = LastNegotiatedFeatures.HasFlag(HotlineConstants.TrackerFeatures.Query);
        var parameters = new List<byte[]>();
        if (canQuery && !string.IsNullOrWhiteSpace(search))
        {
            parameters.Add(Tlv(HotlineConstants.TrackerQuerySearchText, Encoding.UTF8.GetBytes(search)));
        }

        if (canQuery && limit > 0)
        {
            var value = new byte[2];
            BinaryPrimitives.WriteUInt16BigEndian(value, limit);
            parameters.Add(Tlv(HotlineConstants.TrackerQueryPageLimit, value));
        }

        var request = new byte[4 + parameters.Sum(p => p.Length)];
        BinaryPrimitives.WriteUInt16BigEndian(request, 0x0001); // list request
        BinaryPrimitives.WriteUInt16BigEndian(request.AsSpan(2), (ushort)parameters.Count);
        var offset = 4;
        foreach (var parameter in parameters)
        {
            parameter.CopyTo(request.AsSpan(offset));
            offset += parameter.Length;
        }

        await stream.WriteAsync(request, ct).ConfigureAwait(false);

        var header = await ReadExactAsync(stream, 10, ct).ConfigureAwait(false);
        if (header is null || BinaryPrimitives.ReadUInt16BigEndian(header) != 0x0001)
        {
            return [];
        }

        var recordCount = BinaryPrimitives.ReadUInt16BigEndian(header.AsSpan(8));
        var results = new List<HotlineTrackerServerEntry>(recordCount);
        for (var i = 0; i < recordCount; i++)
        {
            var entry = await ReadVersion3RecordAsync(stream, ct).ConfigureAwait(false);
            if (entry is null)
            {
                break;
            }

            results.Add(entry);
        }

        return results;
    }

    private static byte[] Tlv(ushort id, byte[] value)
    {
        var field = new byte[4 + value.Length];
        BinaryPrimitives.WriteUInt16BigEndian(field, id);
        BinaryPrimitives.WriteUInt16BigEndian(field.AsSpan(2), (ushort)value.Length);
        value.CopyTo(field.AsSpan(4));
        return field;
    }

    /// <summary>
    /// One v3 record. The address is tagged rather than assumed to be four bytes, which is what
    /// lets a tracker list an IPv6 server or one known only by hostname — neither of which v1 can
    /// express at all.
    /// </summary>
    private static async Task<HotlineTrackerServerEntry?> ReadVersion3RecordAsync(NetworkStream stream, CancellationToken ct)
    {
        var addressType = await ReadExactAsync(stream, 1, ct).ConfigureAwait(false);
        if (addressType is null)
        {
            return null;
        }

        string address;
        switch (addressType[0])
        {
            case 0x04:
                var v4 = await ReadExactAsync(stream, 4, ct).ConfigureAwait(false);
                if (v4 is null)
                {
                    return null;
                }

                address = new System.Net.IPAddress(v4).ToString();
                break;

            case 0x06:
                var v6 = await ReadExactAsync(stream, 16, ct).ConfigureAwait(false);
                if (v6 is null)
                {
                    return null;
                }

                address = new System.Net.IPAddress(v6).ToString();
                break;

            case 0x48:
                var hostname = await ReadUtf8StringAsync(stream, ct).ConfigureAwait(false);
                if (hostname is null)
                {
                    return null;
                }

                address = hostname;
                break;

            default:
                // An address kind this client doesn't know leaves the stream at an unknown offset;
                // there's no safe way to continue, so stop with what's already parsed.
                return null;
        }

        var fixedPart = await ReadExactAsync(stream, 4, ct).ConfigureAwait(false);
        if (fixedPart is null)
        {
            return null;
        }

        var port = BinaryPrimitives.ReadUInt16BigEndian(fixedPart);
        var users = BinaryPrimitives.ReadUInt16BigEndian(fixedPart.AsSpan(2));

        var name = await ReadUtf8StringAsync(stream, ct).ConfigureAwait(false);
        var description = await ReadUtf8StringAsync(stream, ct).ConfigureAwait(false);
        if (name is null || description is null)
        {
            return null;
        }

        // Trailing metadata this client doesn't need — read past it so the next record starts
        // where it should.
        var tlvCount = await ReadExactAsync(stream, 2, ct).ConfigureAwait(false);
        if (tlvCount is null)
        {
            return null;
        }

        for (var i = 0; i < BinaryPrimitives.ReadUInt16BigEndian(tlvCount); i++)
        {
            var tlvHeader = await ReadExactAsync(stream, 4, ct).ConfigureAwait(false);
            if (tlvHeader is null)
            {
                return null;
            }

            var length = BinaryPrimitives.ReadUInt16BigEndian(tlvHeader.AsSpan(2));
            if (length > 0 && await ReadExactAsync(stream, length, ct).ConfigureAwait(false) is null)
            {
                return null;
            }
        }

        return new HotlineTrackerServerEntry(address, port, users, name, description);
    }

    /// <summary>A v3 string: a 2-byte length then UTF-8 — the older format's 1-byte length and Mac Roman are both gone.</summary>
    private static async Task<string?> ReadUtf8StringAsync(NetworkStream stream, CancellationToken ct)
    {
        var lengthBytes = await ReadExactAsync(stream, 2, ct).ConfigureAwait(false);
        if (lengthBytes is null)
        {
            return null;
        }

        var length = BinaryPrimitives.ReadUInt16BigEndian(lengthBytes);
        if (length == 0)
        {
            return "";
        }

        var data = await ReadExactAsync(stream, length, ct).ConfigureAwait(false);
        return data is null ? null : Encoding.UTF8.GetString(data);
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
