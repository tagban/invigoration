using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// Retail-captured vectors reproduced verbatim from ncarrillo/superiority's
/// core/src/native/decode.rs unit tests (relayed via research agent).
/// </summary>
public class ChatRecordDecoderTests
{
    [Fact]
    public void DecodeChatMessage_RetailVector_DecodesAtExactBoundary()
    {
        var packet = Convert.FromHexString(
            "4b0505012f0b04616e796f6f6e6520666f72206d75746174696f6e206f72206f6e65206d697373696f6e3f00");
        var reader = new BitReader(packet);
        var routing = RoutingHeader.Decode(reader);

        var message = ChatRecordDecoder.DecodeChatMessage(reader);

        Assert.Equal(ChatCommands.ChatSlot, routing.ServiceSlot);
        Assert.Equal(336, reader.Position - 11);
        Assert.Equal(0, message.ChannelIndex);
        Assert.Equal(2_623_867u, message.MemberHandle);
        Assert.Equal("anyoone for mutation or one mission?", message.Body);
    }

    [Fact]
    public void DecodeChatWhisper_RetailVector_DecodesPeerAndBody()
    {
        var packet = Convert.FromHexString(
            "5305414a682e0000000019034e656c736f6e54657374393123313435380100686f6c61");
        var reader = new BitReader(packet);
        RoutingHeader.Decode(reader);

        var whisper = ChatRecordDecoder.DecodeChatWhisper(reader);

        Assert.Equal(1, whisper.PeerRegion);
        Assert.Equal(FourCc.Encode("BSAp"), whisper.PeerProgramId);
        Assert.Equal(1u, whisper.PeerRealm);
        Assert.Equal("NelsonTest91#1458", whisper.PeerName);
        Assert.Equal("hola", whisper.Body);
    }
}
