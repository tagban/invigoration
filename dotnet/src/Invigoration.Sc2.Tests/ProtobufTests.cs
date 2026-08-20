using Invigoration.Sc2.Front;

namespace Invigoration.Sc2.Tests;

public class FrontFrameTests
{
    [Fact]
    public void Encode_ConnectRequestFrame_MatchesUpstreamVector()
    {
        var header = new Header
        {
            ServiceId = 0,
            MethodId = 1,
            Token = 0,
            Size = 2,
            ServiceHash = 0x65446991,
        };
        var body = new ConnectRequest { UseBindlessRpc = true }.Encode();

        var frame = FrontFrame.Encode(header, body);

        Assert.Equal("000d08001001180028025d916944651801", Convert.ToHexString(frame).ToLowerInvariant());
    }

    [Fact]
    public void Decode_UpstreamVector_RoundTripsHeaderAndBody()
    {
        var frame = Convert.FromHexString("000d08001001180028025d916944651801");

        var (header, body) = FrontFrame.Decode(frame);

        Assert.Equal(0u, header.ServiceId);
        Assert.Equal(1u, header.MethodId);
        Assert.Equal(0u, header.Token);
        Assert.Equal(2u, header.Size);
        Assert.Equal(0x65446991u, header.ServiceHash);

        var request = ConnectRequest.Decode(body);
        Assert.True(request.UseBindlessRpc);
    }
}

public class EntityIdTests
{
    [Fact]
    public void EncodeDecode_RoundTrips()
    {
        var id = new EntityId { High = 0x0102030405060708, Low = 0x1112131415161718 };

        var decoded = EntityId.Decode(id.Encode());

        Assert.Equal(id.High, decoded.High);
        Assert.Equal(id.Low, decoded.Low);
    }
}

public class AttributeTests
{
    [Fact]
    public void EncodeDecode_StringVariant_RoundTrips()
    {
        var attribute = new Front.Attribute { Name = "Client.Name", Value = new Variant { StringValue = "Tagban" } };

        var decoded = Front.Attribute.Decode(attribute.Encode());

        Assert.Equal("Client.Name", decoded.Name);
        Assert.Equal("Tagban", decoded.Value.StringValue);
    }

    [Fact]
    public void EncodeDecode_UintVariant_RoundTrips()
    {
        var attribute = new Front.Attribute { Name = "Client.Type", Value = new Variant { UintValue = 42 } };

        var decoded = Front.Attribute.Decode(attribute.Encode());

        Assert.Equal(42ul, decoded.Value.UintValue);
    }
}

public class LogonResultTests
{
    [Fact]
    public void Decode_WithBattleTagAndErrorCode_ParsesFields()
    {
        var accountId = new EntityId { High = 1, Low = 2 };
        var w = new Protobuf.ProtoWriter();
        w.WriteUInt32(1, 0);
        w.WriteBytesField(2, accountId.Encode());
        w.WriteString(7, "Tagban#1234");
        var data = w.ToArray();

        var result = LogonResult.Decode(data);

        Assert.Equal(0u, result.ErrorCode);
        Assert.Equal(1ul, result.AccountId!.High);
        Assert.Equal("Tagban#1234", result.BattleTag);
    }
}
