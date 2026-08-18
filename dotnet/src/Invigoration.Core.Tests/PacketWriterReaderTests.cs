using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

public class PacketWriterReaderTests
{
    [Fact]
    public void ToBncsPacket_ProducesCorrectHeaderAndLength()
    {
        var packet = new PacketWriter()
            .WriteDword(0x12345678)
            .WriteNTString("abc")
            .ToBncsPacket(BncsPacketId.SID_AUTH_INFO);

        // FF, id, len(2 LE, includes 4-byte header), payload
        Assert.Equal(0xFF, packet[0]);
        Assert.Equal((byte)BncsPacketId.SID_AUTH_INFO, packet[1]);
        var length = (ushort)(packet[2] | (packet[3] << 8));
        Assert.Equal(packet.Length, length);
        Assert.Equal(4 + 4 + 4, packet.Length); // header + dword + "abc\0"
    }

    [Fact]
    public void ToBnlsPacket_ProducesCorrectHeaderAndLength()
    {
        var packet = new PacketWriter()
            .WriteByte(0x42)
            .ToBnlsPacket(BnlsPacketId.BNLS_HASHDATA);

        var length = (ushort)(packet[0] | (packet[1] << 8));
        Assert.Equal(packet.Length, length);
        Assert.Equal((byte)BnlsPacketId.BNLS_HASHDATA, packet[2]);
        Assert.Equal(3 + 1, packet.Length);
    }

    [Fact]
    public void ReadWriteRoundTrip_PreservesValues()
    {
        var writer = new PacketWriter()
            .WriteByte(0xAB)
            .WriteWord(0x1234)
            .WriteDword(0xDEADBEEF)
            .WriteNTString("hello")
            .WriteBytes([1, 2, 3]);

        var payload = writer.ToRealmPacket(0x01);
        // Strip the 3-byte realm header to get back the raw payload.
        var reader = new PacketReader(payload, offset: 3);

        Assert.Equal(0xAB, reader.ReadByte());
        Assert.Equal(0x1234, reader.ReadWord());
        Assert.Equal(0xDEADBEEFu, reader.ReadDword());
        Assert.Equal("hello", reader.ReadNTString());
        Assert.Equal([1, 2, 3], reader.ReadRaw(3));
    }

    [Fact]
    public void ReadBoolean_ReadsFullDwordAsNonZeroCheck()
    {
        var payload = new PacketWriter().WriteDword(0).WriteDword(1).ToRealmPacket(0);
        var reader = new PacketReader(payload, offset: 3);

        Assert.False(reader.ReadBoolean());
        Assert.True(reader.ReadBoolean());
    }

    [Fact]
    public void ReadFileTime_ReadsLowThenHighDword()
    {
        var payload = new PacketWriter().WriteDword(0x11111111).WriteDword(0x22222222).ToRealmPacket(0);
        var reader = new PacketReader(payload, offset: 3);

        var fileTime = reader.ReadFileTime();

        Assert.Equal(0x11111111u, fileTime.Low);
        Assert.Equal(0x22222222u, fileTime.High);
    }
}
