using System.Buffers.Binary;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Tracker version negotiation. The client announces v3 and takes whatever the tracker answers
/// with, so the risk is in the split: reading the v3 feature bytes from a v1 tracker would eat the
/// start of its listing, and not reading them from a v3 one would desynchronize everything after.
/// Both directions are driven here against loopback stand-ins.
/// </summary>
[Collection("HotlineTracker")]
public class HotlineTrackerVersionTests
{
    private static byte[] HandshakeReply(ushort version, ushort? features = null)
    {
        var reply = new byte[features is null ? 6 : 8];
        "HTRK"u8.CopyTo(reply.AsSpan(0));
        BinaryPrimitives.WriteUInt16BigEndian(reply.AsSpan(4), version);
        if (features is { } f)
        {
            BinaryPrimitives.WriteUInt16BigEndian(reply.AsSpan(6), f);
        }

        return reply;
    }

    private static byte[] Utf8String(string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        var field = new byte[2 + bytes.Length];
        BinaryPrimitives.WriteUInt16BigEndian(field, (ushort)bytes.Length);
        bytes.CopyTo(field.AsSpan(2));
        return field;
    }

    private static byte[] PascalString(string value)
    {
        var bytes = Encoding.Latin1.GetBytes(value);
        var field = new byte[1 + bytes.Length];
        field[0] = (byte)bytes.Length;
        bytes.CopyTo(field.AsSpan(1));
        return field;
    }

    /// <summary>A v3 record: a tagged address, port and users, two UTF-8 strings, then metadata.</summary>
    private static byte[] Version3Record(byte addressType, byte[] address, ushort port, ushort users, string name, string description)
    {
        var body = new List<byte> { addressType };
        body.AddRange(address);
        var fixedPart = new byte[4];
        BinaryPrimitives.WriteUInt16BigEndian(fixedPart, port);
        BinaryPrimitives.WriteUInt16BigEndian(fixedPart.AsSpan(2), users);
        body.AddRange(fixedPart);
        body.AddRange(Utf8String(name));
        body.AddRange(Utf8String(description));
        body.AddRange([0, 0]); // no metadata
        return [.. body];
    }

    private static byte[] Version3Response(params byte[][] records)
    {
        var body = new List<byte>();
        var header = new byte[10];
        BinaryPrimitives.WriteUInt16BigEndian(header, 0x0001);
        BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(2), (uint)records.Sum(r => r.Length));
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(6), (ushort)records.Length);
        BinaryPrimitives.WriteUInt16BigEndian(header.AsSpan(8), (ushort)records.Length);
        body.AddRange(header);
        foreach (var record in records)
        {
            body.AddRange(record);
        }

        return [.. body];
    }

    private static (TcpListener Listener, int Port) StartListener()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        return (listener, ((IPEndPoint)listener.LocalEndpoint).Port);
    }

    [Fact]
    public async Task AVersion3Tracker_IsNegotiatedAndItsRecordsRead()
    {
        var (listener, port) = StartListener();
        var clientHandshake = new byte[8];
        try
        {
            var served = Task.Run(async () =>
            {
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();
                await stream.ReadExactlyAsync(clientHandshake);
                await stream.WriteAsync(HandshakeReply(3, features: 0x0003));

                // The client's listing request, then the listing.
                var requestHeader = new byte[4];
                await stream.ReadExactlyAsync(requestHeader);
                var fieldCount = BinaryPrimitives.ReadUInt16BigEndian(requestHeader.AsSpan(2));
                for (var i = 0; i < fieldCount; i++)
                {
                    var tlv = new byte[4];
                    await stream.ReadExactlyAsync(tlv);
                    await stream.ReadExactlyAsync(new byte[BinaryPrimitives.ReadUInt16BigEndian(tlv.AsSpan(2))]);
                }

                await stream.WriteAsync(Version3Response(
                    Version3Record(0x04, IPAddress.Parse("203.0.113.9").GetAddressBytes(), 5500, 12, "Modern Server", "v3 listed"),
                    Version3Record(0x48, Utf8String("hotline.example.org"), 5500, 3, "By Name", "hostname record")));
            });

            var servers = await HotlineTrackerClient.QueryAsync("127.0.0.1", port);
            await served.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Equal(3, HotlineTrackerClient.LastNegotiatedVersion);
            Assert.Equal(2, servers.Count);
            Assert.Equal("203.0.113.9", servers[0].Address);
            Assert.Equal("Modern Server", servers[0].Name);
            Assert.Equal((ushort)12, servers[0].UserCount);

            // A hostname record — something v1 can't express at all.
            Assert.Equal("hotline.example.org", servers[1].Address);

            // The client announced v3 with its feature bytes.
            Assert.Equal("HTRK", Encoding.ASCII.GetString(clientHandshake, 0, 4));
            Assert.Equal(3, BinaryPrimitives.ReadUInt16BigEndian(clientHandshake.AsSpan(4)));
            Assert.NotEqual(0, BinaryPrimitives.ReadUInt16BigEndian(clientHandshake.AsSpan(6)));
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>
    /// The compatibility case that matters: an old tracker answers with six bytes and its classic
    /// listing. Reading two more would swallow the first batch header.
    /// </summary>
    [Fact]
    public async Task AVersion1Tracker_StillWorks()
    {
        var (listener, port) = StartListener();
        try
        {
            var served = Task.Run(async () =>
            {
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();

                // A v1 tracker reads six bytes and ignores whatever else arrived.
                await stream.ReadExactlyAsync(new byte[6]);
                await stream.WriteAsync(HandshakeReply(1));

                var batchHeader = new byte[8];
                BinaryPrimitives.WriteUInt16BigEndian(batchHeader.AsSpan(4), 1); // total across batches
                BinaryPrimitives.WriteUInt16BigEndian(batchHeader.AsSpan(6), 1); // in this batch
                await stream.WriteAsync(batchHeader);

                var record = new byte[10];
                IPAddress.Parse("198.51.100.20").GetAddressBytes().CopyTo(record, 0);
                BinaryPrimitives.WriteUInt16BigEndian(record.AsSpan(4), 5500);
                BinaryPrimitives.WriteUInt16BigEndian(record.AsSpan(6), 7);
                await stream.WriteAsync(record);
                await stream.WriteAsync(PascalString("Classic Server"));
                await stream.WriteAsync(PascalString("still here"));
            });

            var servers = await HotlineTrackerClient.QueryAsync("127.0.0.1", port);
            await served.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Equal(1, HotlineTrackerClient.LastNegotiatedVersion);
            var server = Assert.Single(servers);
            Assert.Equal("198.51.100.20", server.Address);
            Assert.Equal("Classic Server", server.Name);
            Assert.Equal((ushort)7, server.UserCount);
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>
    /// The real-world case, confirmed live against hltracker.com (2026-09-15): a tracker that
    /// hangs up the moment it sees version 3, rather than answering with its own version as the
    /// spec claims older trackers do. The client must retry on a fresh connection — continuing on
    /// the closed one gets nothing — or the default tracker stops working entirely.
    /// </summary>
    [Fact]
    public async Task ATrackerThatRefusesVersion3Outright_IsRetriedAsVersion1()
    {
        var (listener, port) = StartListener();
        var versionsSeen = new List<ushort>();
        try
        {
            var served = Task.Run(async () =>
            {
                // First connection: sees v3 and hangs up without a word.
                using (var refused = await listener.AcceptTcpClientAsync())
                {
                    await using var refusedStream = refused.GetStream();
                    var handshake = new byte[8];
                    await refusedStream.ReadExactlyAsync(handshake);
                    versionsSeen.Add(BinaryPrimitives.ReadUInt16BigEndian(handshake.AsSpan(4)));
                }

                // Second connection: the client retries as v1, which works.
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();
                var v1Handshake = new byte[6];
                await stream.ReadExactlyAsync(v1Handshake);
                versionsSeen.Add(BinaryPrimitives.ReadUInt16BigEndian(v1Handshake.AsSpan(4)));
                await stream.WriteAsync(HandshakeReply(1));

                var batchHeader = new byte[8];
                BinaryPrimitives.WriteUInt16BigEndian(batchHeader.AsSpan(4), 1);
                BinaryPrimitives.WriteUInt16BigEndian(batchHeader.AsSpan(6), 1);
                await stream.WriteAsync(batchHeader);

                var record = new byte[10];
                IPAddress.Parse("192.0.2.5").GetAddressBytes().CopyTo(record, 0);
                BinaryPrimitives.WriteUInt16BigEndian(record.AsSpan(4), 5500);
                BinaryPrimitives.WriteUInt16BigEndian(record.AsSpan(6), 2);
                await stream.WriteAsync(record);
                await stream.WriteAsync(PascalString("Old Tracker Server"));
                await stream.WriteAsync(PascalString("refused v3"));
            });

            var servers = await HotlineTrackerClient.QueryAsync("127.0.0.1", port);
            await served.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Equal([3, 1], versionsSeen);
            Assert.Equal("Old Tracker Server", Assert.Single(servers).Name);
        }
        finally
        {
            listener.Stop();
        }
    }

    /// <summary>Search and a limit are only sent to a tracker that says it supports them — an older one would choke on unexpected parameters.</summary>
    [Fact]
    public async Task SearchTerms_AreOnlySentWhenTheTrackerOffersQuerying()
    {
        var (listener, port) = StartListener();
        try
        {
            ushort fieldCount = 0xFFFF;
            var served = Task.Run(async () =>
            {
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();
                await stream.ReadExactlyAsync(new byte[8]);

                // v3, but without the query feature bit.
                await stream.WriteAsync(HandshakeReply(3, features: 0x0001));

                var requestHeader = new byte[4];
                await stream.ReadExactlyAsync(requestHeader);
                fieldCount = BinaryPrimitives.ReadUInt16BigEndian(requestHeader.AsSpan(2));
                await stream.WriteAsync(Version3Response());
            });

            await HotlineTrackerClient.QueryAsync("127.0.0.1", port, search: "games", limit: 25);
            await served.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Equal(0, fieldCount);
        }
        finally
        {
            listener.Stop();
        }
    }

    [Fact]
    public async Task SearchTerms_AreSentWhenTheTrackerDoesOfferQuerying()
    {
        var (listener, port) = StartListener();
        try
        {
            var sentIds = new List<ushort>();
            var served = Task.Run(async () =>
            {
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();
                await stream.ReadExactlyAsync(new byte[8]);
                await stream.WriteAsync(HandshakeReply(3, features: 0x0003));

                var requestHeader = new byte[4];
                await stream.ReadExactlyAsync(requestHeader);
                for (var i = 0; i < BinaryPrimitives.ReadUInt16BigEndian(requestHeader.AsSpan(2)); i++)
                {
                    var tlv = new byte[4];
                    await stream.ReadExactlyAsync(tlv);
                    sentIds.Add(BinaryPrimitives.ReadUInt16BigEndian(tlv));
                    await stream.ReadExactlyAsync(new byte[BinaryPrimitives.ReadUInt16BigEndian(tlv.AsSpan(2))]);
                }

                await stream.WriteAsync(Version3Response());
            });

            await HotlineTrackerClient.QueryAsync("127.0.0.1", port, search: "games", limit: 25);
            await served.WaitAsync(TimeSpan.FromSeconds(10));

            Assert.Contains(HotlineConstants.TrackerQuerySearchText, sentIds);
            Assert.Contains(HotlineConstants.TrackerQueryPageLimit, sentIds);
        }
        finally
        {
            listener.Stop();
        }
    }
}

/// <summary>HotlineTrackerClient records the last negotiated version in static state, so these run one at a time.</summary>
[CollectionDefinition("HotlineTracker")]
public class HotlineTrackerCollection;
