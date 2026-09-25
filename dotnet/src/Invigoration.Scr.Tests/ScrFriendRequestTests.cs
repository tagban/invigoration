using Invigoration.Scr;

namespace Invigoration.Scr.Tests;

public class ScrFriendRequestTests
{
    [Fact]
    public void SendWhisperWritesTheAccountAsFixed32()
    {
        var body = ScrWhispers.SendRequest(0x01020304, "hi");

        Assert.Equal("0D0403020112026869", Convert.ToHexString(body));
    }

    [Theory]
    [InlineData(ScrWhispers.WhisperReceivedMethod, false)]
    [InlineData(ScrWhispers.WhisperEchoReceivedMethod, true)]
    public void DecodesBattlenetWhispersAndTheirEchoes(uint method, bool outgoing)
    {
        var whisper = ScrWhispers.Decode(method, ScrWhispers.SendRequest(41, "hello"));

        Assert.Equal(new ScrWhisper(41, "hello", outgoing), whisper);
    }

    [Fact]
    public void IgnoresAWhisperWhoseAccountIsNotFixed32()
    {
        // {1: 41 as a varint, 2: "x"}: the client's own parser refuses this too.
        Assert.Null(ScrWhispers.Decode(ScrWhispers.WhisperReceivedMethod, Convert.FromHexString("0829120178")));
    }

    [Fact]
    public void BuildsFriendRequests()
    {
        Assert.Equal("0A06412331323334", Convert.ToHexString(ScrFriends.SendInvitationRequest("A#1234")));
        Assert.Equal("082A", Convert.ToHexString(ScrFriends.RemoveFriendRequest(42)));
        Assert.Equal("08AC02", Convert.ToHexString(ScrFriends.AnswerInvitationRequest(300)));
    }

    [Fact]
    public void DecodesAnInvitationAndItsRemoval()
    {
        // {1: {1: 300, 2: "A#1234"}, 2: removed}
        var invitation = Convert.FromHexString("0A0B08AC021206412331323334");

        Assert.Equal((new ScrInvitation(300, "A#1234"), false), ScrFriends.DecodeInvitation(invitation));
        Assert.Equal((new ScrInvitation(300, "A#1234"), true), ScrFriends.DecodeInvitation([.. invitation, 0x10, 0x01]));
    }
}

public class ScrStatsTests
{
    [Fact]
    public void RequestMatchesWhatTheGameSends()
    {
        // No program, gateway 10, the name, ".*", 0xFFFFFFFF.
        Assert.Equal("100A1A035365782202" + "2E2A" + "28FFFFFFFF0F", Convert.ToHexString(ScrStats.Request("Sex", null, 10)));
    }

    [Fact]
    public void DescribesALiveReply()
    {
        // A live reply for a classic player (2026-09-25).
        var (stats, status) = ScrStats.Decode(Convert.FromHexString(
            "0A180A126C65676163795F646973636F6E6E65637473100018000A120A0B6C65676163795F77696E7310EF0618000A130A0D6C65676163795F6C6F73736573100018000A270A196C65676163795F746F6F6E5F6372656174696F6E5F74696D65108E93EFE2B0C187E90118001014"));

        Assert.Equal(20u, status);
        Assert.Equal("classic: 879 wins, 0 losses · created 4 Oct 2016", ScrStats.Describe(stats));
    }
}
