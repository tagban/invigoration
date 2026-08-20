using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Every test vector here is reproduced verbatim from
/// ncarrillo/superiority's core/src/bsn/bits.rs unit tests, not invented —
/// this is the one part of the SC2 port we can verify against known-correct
/// output without a live server.
/// </summary>
public class BitWriterTests
{
    [Fact]
    public void CrossByteValue_UsesHighChunkFirst()
    {
        var writer = new BitWriter();

        writer.Write(0, 6);
        writer.Write(1, 1);
        writer.Write(14, 4);

        Assert.Equal(11, writer.Position);
        Assert.Equal(new byte[] { 0xc0, 0x06 }, writer.ToBytes());
    }

    [Theory]
    [InlineData(new byte[] { 0x40, 0x0c }, 0, 4)]
    [InlineData(new byte[] { 0x41, 0x05 }, 1, 5)]
    [InlineData(new byte[] { 0xc0, 0x06 }, 0, 14)]
    public void ObservedRoutingHeaders_Decode(byte[] data, byte expectedCommand, byte expectedSlot)
    {
        var reader = new BitReader(data);

        var header = RoutingHeader.Decode(reader);

        Assert.Equal(expectedCommand, header.CommandId);
        Assert.Equal(expectedSlot, header.ServiceSlot);
        Assert.Equal(11, header.BitCount);
    }

    [Fact]
    public void AlignedBytes_RoundTrip()
    {
        var writer = new BitWriter();
        writer.Write(5, 3);

        var skipped = writer.Align();

        Assert.Equal(5, skipped);
        writer.WriteBytes("abc"u8, aligned: true);
        var bytes = writer.ToBytes();

        var reader = new BitReader(bytes);
        Assert.Equal(5u, reader.Read(3));
        Assert.Equal(5, reader.Align());
        Assert.Equal("abc"u8.ToArray(), reader.ReadBytes(3, aligned: true));
    }

    [Fact]
    public void Write_ThenRead_RoundTripsArbitraryValues()
    {
        var writer = new BitWriter();
        writer.Write(0b101, 3);
        writer.Write(0x1234, 16);
        writer.Write(1, 1);
        var bytes = writer.ToBytes();

        var reader = new BitReader(bytes);
        Assert.Equal(0b101u, reader.Read(3));
        Assert.Equal(0x1234u, reader.Read(16));
        Assert.Equal(1u, reader.Read(1));
    }

    [Fact]
    public void WriteRaw_ThenReadRaw_RoundTrips()
    {
        var writer = new BitWriter();
        writer.Write(0b11, 2); // shift the raw copy off a byte boundary
        writer.WriteRaw([0b1011_0110], 6);
        var bytes = writer.ToBytes();

        var reader = new BitReader(bytes);
        Assert.Equal(0b11u, reader.Read(2));
        var raw = reader.ReadRaw(6);
        // ReadRaw returns the extracted bits packed LSB-first starting at bit 0 of a fresh buffer.
        var check = new BitReader(raw);
        Assert.Equal((ulong)(0b1011_0110 & 0x3f), ReadLsbFirst(check, 6));
    }

    private static ulong ReadLsbFirst(BitReader reader, int bitCount)
    {
        ulong value = 0;
        for (var i = 0; i < bitCount; i++)
        {
            value |= reader.Read(1) << i;
        }

        return value;
    }
}

public class FourCcTests
{
    [Fact]
    public void Encode_ShortString_LeftPadsWithZeroBytes()
    {
        Assert.Equal(0x00005332u, FourCc.Encode("S2"));
    }

    [Fact]
    public void Encode_FourCharString_UsesAllBytes()
    {
        // Standard big-endian byte packing for a full 4-char string, no left-pad ambiguity:
        // 'M'=0x4D 'c'=0x63 '6'=0x36 '4'=0x34.
        Assert.Equal(0x4d633634u, FourCc.Encode("Mc64"));
    }
}

public class ServiceHashTests
{
    [Theory]
    [InlineData("bnet.protocol.connection.ConnectionService", 0x65446991u)]
    [InlineData("bnet.protocol.authentication.AuthenticationServer", 0x0decfc01u)]
    [InlineData("bnet.protocol.game_utilities.GameUtilities", 0x3fc1274du)]
    public void Compute_MatchesReferenceVectors(string serviceName, uint expected)
    {
        Assert.Equal(expected, ServiceHash.Compute(serviceName));
    }
}
