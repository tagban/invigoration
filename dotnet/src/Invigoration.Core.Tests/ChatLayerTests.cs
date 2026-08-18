using Invigoration.Core.Chat;
using Invigoration.Core.Crypto;
using Invigoration.Core.Protocol;
using Invigoration.Core.Text;

namespace Invigoration.Core.Tests;

public class ChatEventParserTests
{
    [Fact]
    public void Parse_ExtractsEventIdFlagsPingUsernameAndText()
    {
        var frame = new PacketWriter()
            .WriteDword((uint)ChatEventType.Talk)
            .WriteDword(0x2) // flags: operator
            .WriteDword(42) // ping
            .WriteDword(0).WriteDword(0).WriteDword(0) // ip / account number / reg authority (unused)
            .WriteNTString("someuser")
            .WriteNTString("hello world")
            .ToBncsPacket(BncsPacketId.SID_CHATEVENT);

        var result = ChatEventParser.Parse(frame);

        Assert.Equal(ChatEventType.Talk, result.Type);
        Assert.Equal("someuser", result.Username);
        Assert.Equal(0x2u, result.Flags);
        Assert.Equal(42, result.Ping);
        Assert.Equal("hello world", result.Text);
    }
}

public class ChatColorFormatterTests
{
    [Fact]
    public void Parse_NoMarkers_ReturnsSingleSegmentInDefaultColor()
    {
        var segments = ChatColorFormatter.Parse("plain text", ChatColors.White);

        var segment = Assert.Single(segments);
        Assert.Equal(ChatColors.White, segment.Color);
        Assert.Equal("plain text", segment.Text);
    }

    [Fact]
    public void Parse_MarkerSwitchesColorForRemainderOfMessage()
    {
        var text = $"before{' '}rafter";

        var segments = ChatColorFormatter.Parse(text, ChatColors.White);

        Assert.Equal(2, segments.Count);
        Assert.Equal("before", segments[0].Text);
        Assert.Equal(ChatColors.White, segments[0].Color);
        Assert.Equal("after", segments[1].Text);
        Assert.Equal(RgbColor.FromWin32Bgr(0xFF), segments[1].Color); // red
    }
}

public class InvigCipherTests
{
    [Theory]
    [InlineData("hi!!")]
    [InlineData("test message!!")]
    public void EncryptThenDecrypt_RoundTrips(string text)
    {
        var encrypted = InvigCipher.Encrypt(text);
        var decrypted = InvigCipher.Decrypt(encrypted);

        Assert.Equal(text, decrypted);
    }

    [Fact]
    public void Encrypt_OddLength_DropsLastCharacter()
    {
        // "abc" has odd length (3); only "ab" should survive the round trip.
        var encrypted = InvigCipher.Encrypt("abc");
        var decrypted = InvigCipher.Decrypt(encrypted);

        Assert.Equal("ab", decrypted);
    }
}

public class HexCodecTests
{
    [Fact]
    public void StrToHex_ThenHexToStr_RoundTrips()
    {
        const string text = "Hello, Battle.net!";

        var hex = HexCodec.StrToHex(text);
        var result = HexCodec.HexToStr(hex);

        Assert.Equal(text, result);
    }

    [Fact]
    public void StrToHex_ProducesUppercaseTwoDigitPairs()
    {
        Assert.Equal("41", HexCodec.StrToHex("A"));
    }
}
