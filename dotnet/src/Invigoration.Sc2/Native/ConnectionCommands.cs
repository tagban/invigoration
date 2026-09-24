using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Connection-slot keepalive records. Battle.net pings the client on Sunken and
/// expects a Pong straight back; the client should also send its own Ping about
/// every 30 seconds. See docs.bnet.cc's "StarCraft II chat on Sunken".
/// </summary>
public static class ConnectionCommands
{
    public const byte ConnectionSlot = 1;
    public const byte PingCommand = 10;
    public const byte PongCommand = 12;

    /// <summary>How often to send our own <see cref="Ping"/>.</summary>
    public static readonly TimeSpan PingInterval = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Our own ping: a presence bit of 1, then the timestamp's high and low
    /// 32 bits written straight after it, unaligned. The server echoes the
    /// timestamp back in its Pong, so any one steady microsecond clock will do.
    /// </summary>
    public static byte[] Ping(ulong timestampMicroseconds)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, PingCommand, ConnectionSlot);
        writer.Write(1, 1);
        writer.Write(timestampMicroseconds >> 32, 32);
        writer.Write(timestampMicroseconds & 0xFFFF_FFFF, 32);
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>
    /// The answer to a server Ping: the same presence bit and the same eight
    /// timestamp bytes, laid out exactly as <see cref="ConnectionRecordDecoder.DecodeKeepalive"/>
    /// read them (presence bit, byte-align, raw bytes), so the server gets back
    /// precisely what it sent.
    /// </summary>
    public static byte[] Pong(byte[]? timestamp)
    {
        if (timestamp is { Length: not 8 })
        {
            throw new ArgumentException("A Sunken ping timestamp is exactly 8 bytes.", nameof(timestamp));
        }

        var writer = new BitWriter();
        RoutingHeader.Encode(writer, PongCommand, ConnectionSlot);
        writer.Write(timestamp is null ? 0UL : 1UL, 1);
        if (timestamp is not null)
        {
            writer.WriteBytes(timestamp, aligned: true);
        }

        writer.Align();
        return writer.ToBytes();
    }
}
