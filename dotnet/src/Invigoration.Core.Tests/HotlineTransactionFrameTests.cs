using System.Buffers.Binary;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Verifies HotlineTransactionFrame's encode/decode against the exact byte layout confirmed from
/// Hotline-Navigator's real source (transaction.rs) — 20-byte header (flags, is_reply, type, id,
/// error_code, total_size, data_size) then field_count + repeated (type, size, data) fields, all
/// big-endian.
/// </summary>
public class HotlineTransactionFrameTests
{
    [Fact]
    public void Encode_HeaderFieldsAreBigEndianInCorrectOffsets()
    {
        var frame = HotlineTransactionFrame.Create(HotlineTransactionType.SendChat, id: 0x01020304, new HotlineField(HotlineFieldType.Data, "hi"));

        var bytes = frame.Encode();

        Assert.Equal(0, bytes[0]); // flags
        Assert.Equal(0, bytes[1]); // is_reply = false
        Assert.Equal((ushort)HotlineTransactionType.SendChat, BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(2)));
        Assert.Equal(0x01020304u, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(4)));
        Assert.Equal(0u, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(8))); // error_code
        var bodySize = (uint)(2 /* field count */ + 4 /* field header */ + 2 /* "hi" */);
        Assert.Equal(bodySize, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(12))); // total_size
        Assert.Equal(bodySize, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(16))); // data_size
        Assert.Equal(20 + (int)bodySize, bytes.Length);
    }

    [Fact]
    public void CreateReply_SetsIsReplyAndErrorCode()
    {
        var frame = HotlineTransactionFrame.CreateReply(replyToId: 42, errorCode: 1000);

        var bytes = frame.Encode();

        Assert.Equal(1, bytes[1]); // is_reply = true
        Assert.Equal(42u, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(4)));
        Assert.Equal(1000u, BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(8)));
    }

    [Fact]
    public void EncodeThenDecode_RoundTripsFieldsExactly()
    {
        var original = HotlineTransactionFrame.Create(
            HotlineTransactionType.Login,
            id: 7,
            new HotlineField(HotlineFieldType.UserLogin, [0x9E, 0x8D]),
            new HotlineField(HotlineFieldType.UserIconId, (ushort)414),
            new HotlineField(HotlineFieldType.UserName, "TestBot"));

        var decoded = HotlineTransactionFrame.Decode(original.Encode());

        Assert.Equal(original.Type, decoded.Type);
        Assert.Equal(original.Id, decoded.Id);
        Assert.Equal(3, decoded.Fields.Count);
        Assert.Equal([0x9E, 0x8D], decoded.Field(HotlineFieldType.UserLogin)!.Data);
        Assert.Equal((ushort)414, decoded.Field(HotlineFieldType.UserIconId)!.AsUInt16());
        Assert.Equal("TestBot", decoded.Field(HotlineFieldType.UserName)!.AsString());
    }

    [Fact]
    public void TryGetFrameLength_ReturnsNullUntilHeaderFullyBuffered()
    {
        var full = HotlineTransactionFrame.Create(HotlineTransactionType.SendChat, 1, new HotlineField(HotlineFieldType.Data, "hi")).Encode();

        Assert.Null(HotlineTransactionFrame.TryGetFrameLength(full[..10]));
        Assert.Equal(full.Length, HotlineTransactionFrame.TryGetFrameLength(full[..20]));
    }

    [Fact]
    public void TryGetFrameLength_MatchesActualEncodedLength()
    {
        var frame = HotlineTransactionFrame.Create(HotlineTransactionType.GetUserNameList, 3);
        var bytes = frame.Encode();

        Assert.Equal(bytes.Length, HotlineTransactionFrame.TryGetFrameLength(bytes));
    }

    [Fact]
    public void HotlineField_UInt16Constructor_RoundTrips()
    {
        var field = new HotlineField(HotlineFieldType.UserIconId, (ushort)0xABCD);

        Assert.Equal((ushort)0xABCD, field.AsUInt16());
        Assert.Equal([0xAB, 0xCD], field.Data);
    }
}

public class HotlineTransactionClientTests
{
    [Fact]
    public void XorObfuscate_InvertsEveryByte()
    {
        var result = HotlineTransactionClient.XorObfuscate("AB");

        Assert.Equal(unchecked((byte)~(byte)'A'), result[0]);
        Assert.Equal(unchecked((byte)~(byte)'B'), result[1]);
    }

    [Fact]
    public void XorObfuscate_IsItsOwnInverse()
    {
        var original = "hunter2"u8.ToArray();
        var obfuscated = HotlineTransactionClient.XorObfuscate("hunter2");
        var deobfuscated = obfuscated.Select(b => (byte)~b).ToArray();

        Assert.Equal(original, deobfuscated);
    }
}

public class HotlineUserTests
{
    [Fact]
    public void Parse_DecodesUserNameWithInfoLayout()
    {
        // userId(2) + iconId(2) + flags(2) + nameLength(2) + name(N) — confirmed against
        // Hotline-Navigator's users.rs.
        List<byte> data = [0x00, 0x2A, 0x01, 0x9E, 0x00, 0x00, 0x00, 0x04, .. "Test"u8.ToArray()];

        var user = HotlineUser.Parse([.. data]);

        Assert.Equal(42, user.UserId);
        Assert.Equal(414, user.IconId);
        Assert.Equal(0, user.Flags);
        Assert.Equal("Test", user.Name);
    }

    [Fact]
    public void Parse_OlderServerWithNoLengthPrefix_FallsBackToRestOfDataAsName()
    {
        // Confirmed live against a real server (server.bigredh.com): no nameLength field at all —
        // userId(2) + iconId(2) + flags(2) + name (fills the rest of the field). A naive
        // newer-format read would treat "Te" (the name's own first two bytes) as a bogus
        // nameLength of 0x5465 = 21605 and throw trying to read that many bytes.
        List<byte> data = [0x00, 0x2A, 0x01, 0x9E, 0x00, 0x00, .. "Test"u8.ToArray()];

        var user = HotlineUser.Parse([.. data]);

        Assert.Equal(42, user.UserId);
        Assert.Equal(414, user.IconId);
        Assert.Equal("Test", user.Name);
    }
}
