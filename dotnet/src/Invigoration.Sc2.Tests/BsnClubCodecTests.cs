using Invigoration.Sc2.Bsn;
using Invigoration.Sc2.Native;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Tests;

/// <summary>
/// The schema-driven BSN codec against records the retail client sent and received, as captured
/// in ncarrillo/superiority's tests (MIT, core/src/games/sc2/native/protocol.rs).
/// </summary>
public class BsnClubCodecTests
{
    private const byte ClubSlot = 13;

    private static BsnCodec Codec => ClubLayouts.Codec;

    [Fact]
    public void Schema_HoldsTheClubCommandNumbers()
    {
        var commands = ClubSchema.Instance["Battlenet::Client::Club::CommandID_S2Map::Enum"];

        Assert.Equal(41, commands.IndexValues[commands.MemberNames.ToList().IndexOf("CREATE_CLUB")]);
        Assert.Equal(44, commands.IndexValues[commands.MemberNames.ToList().IndexOf("SET_MEMBER_RANK")]);
        Assert.Equal(57, commands.IndexValues[commands.MemberNames.ToList().IndexOf("CLUB_SETTINGS")]);
    }

    [Fact]
    public void GetToonClubsRequest_MatchesTheOneRetailSends()
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, 46, ClubSlot);
        Codec.Encode(writer, "Battlenet::Client::Club::GetToonClubsRequest", Codec.Struct(
            "Battlenet::Client::Club::GetToonClubsRequest",
            ("GetToonClubs", Codec.Struct("Battlenet::Client::Club::GetToonClubs", ("m_token", 0u))),
            ("m_toon", ToonHandle(region: 1, program: "S2", realm: 1, id: 0x00D5_9F26))));
        writer.Align();

        Assert.Equal(RetailGetToonClubs, Convert.ToHexString(writer.ToBytes()));
    }

    [Fact]
    public void InviteAction_ReadsTheClubAndCode()
    {
        var reader = new BitReader(Convert.FromHexString(Invitation));
        var header = RoutingHeader.Decode(reader);
        var action = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::InviteAction")!;
        var inner = (BsnStruct)action.Fields[0].Value!;

        Assert.Equal((byte)54, header.CommandId);
        Assert.Equal(0L, inner["m_code"]);
        Assert.Equal(535_241L, inner["m_clubId"]);
        Assert.Equal(0x5332u, ((BsnStruct)inner["m_member"]!)["m_programId"]);
    }

    [Theory]
    [InlineData(ThreeClubReply, 215, "cecw|Test Group A|Midigation", "50|30|50")]
    [InlineData(OneClubReply, 81, "Test Group A", "50")]
    public void GetToonClubsResponse_ReadsRetailReplies(string hex, int recordBytes, string names, string ranks)
    {
        var reader = new BitReader(Convert.FromHexString(hex));
        var header = RoutingHeader.Decode(reader);
        var reply = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::GetToonClubsResponse")!;
        reader.Align();
        var success = (BsnStruct)((BsnChoice)reply["m_result"]!).Value!;
        var clubs = ((List<object?>)success["m_clubInfo"]!).Cast<BsnStruct>().Select(c => (BsnStruct)c["m_summary"]!).ToList();

        Assert.Equal((byte)46, header.CommandId);
        Assert.Equal(names, string.Join("|", clubs.Select(c => c["m_name"])));
        Assert.Equal(ranks, string.Join("|", (List<object?>)success["m_rankInfo"]!));
        Assert.Equal(recordBytes * 8, reader.Position);
    }

    [Fact]
    public void ClubCommands_BuildTheRetailRequestAndReadTheReply()
    {
        Assert.Equal(RetailGetToonClubs, Convert.ToHexString(ClubCommands.GetToonClubs(0, new ToonHandleValue(0x5332, 1, 1, 0x00D5_9F26))));

        var reader = new BitReader(Convert.FromHexString(ThreeClubReply));
        RoutingHeader.Decode(reader);
        var clubs = ClubCommands.DecodeToonClubs(reader).Clubs;

        Assert.Equal(["cecw", "Test Group A", "Midigation"], clubs.Select(c => c.Summary.Name));
        Assert.Equal([ClubRanks.Owner, ClubRanks.Member, ClubRanks.Owner], clubs.Select(c => c.Rank));
        Assert.Equal([ClubSummary.GroupType, ClubSummary.GroupType, ClubSummary.ClanType], clubs.Select(c => c.Summary.Type));
        Assert.Equal(535_241u, clubs[1].Summary.Id);
        Assert.All(clubs, c => Assert.Equal("S2", c.Summary.Program));
    }

    [Fact]
    public void GetRoster_SendsClubThenOffsetThenToken()
    {
        var reader = new BitReader(ClubCommands.GetRoster(7, 20_799, 200));
        var header = RoutingHeader.Decode(reader);

        Assert.Equal((byte)45, header.CommandId);
        Assert.Equal(20_799UL, reader.Read(32));
        Assert.Equal(0UL, reader.Read(2));
        Assert.Equal(200UL, reader.Read(32));
        Assert.Equal(7UL, reader.Read(32));
    }

    [Fact]
    public void ResolveToonNames_ListsHandlesInWireOrder()
    {
        var reader = new BitReader(ClubCommands.ResolveToonNames([new ToonHandleValue(0x5332, 1, 1, 0x00D5_9F26)]));
        var header = RoutingHeader.Decode(reader);

        Assert.Equal(((byte)2, (byte?)14), (header.CommandId, header.ServiceSlot));
        Assert.Equal(1UL, reader.Read(6));
        Assert.Equal(0x5332UL, reader.Read(32));
        Assert.Equal(1UL, reader.Read(8));
        Assert.Equal(1UL, reader.Read(32));
        Assert.Equal(0x00D5_9F26UL, reader.Read(64));
    }

    [Fact]
    public void ClubChangeNotification_ReadsALiveSync()
    {
        // The first record Battle.net sent after a roster subscribe (2026-09-25): a group's summary and description.
        var reader = new BitReader(Convert.FromHexString(LiveClubChanges));
        RoutingHeader.Decode(reader);
        var changes = ClubCommands.DecodeClubChanges(reader).Changes;

        Assert.Equal(153 * 8, reader.Position);
        Assert.Equal(["summaryInfoFull", "Description"], changes.Select(c => c.Part));
        Assert.All(changes, c => Assert.Equal(20_801u, c.ClubId));
        Assert.StartsWith("Group chat for BNET.cc", changes[1].Text);
    }

    [Fact]
    public void Filler_FollowsTheRollingState()
    {
        // From an empty buffer the state is only inverted, then rotated.
        Assert.Equal(uint.RotateLeft(~4u, 8), BsnCodec.NextFiller(4, 0, []));
    }

    private static BsnStruct ToonHandle(byte region, string program, uint realm, ulong id) => Codec.Struct(
        "Battlenet::Toon::Handle",
        ("m_region", region),
        ("m_programId", program),
        ("m_realm", realm),
        ("m_id", unchecked((long)id)));

    private const string RetailGetToonClubs = "EE0500000000000A66020100000001000000001AB3E406";

    private const string ThreeClubReply =
        "ee056356e55503000000010082ab09010463656377f72bfaeabe6f45645400000000000000000100014cb200000000d32963020000000000000000000000602b72aa130000000200415609" +
        "010c546573742047726f75702041c12bfaeabe7b81645400000000000000010100014cb200000080a1f320020000000000000000000000602b72aa130000000100415514000a4d6964696761" +
        "74696f6eee2bfaea7ed34d645400000000000000010200014cb200000040054d4447544e42835f0400000000000000000000000003321e32c884d01e000000004d011a00010c000000000920";

    private const string OneClubReply =
        "ee056156e55503000000020082ac09010c546573742047726f75702041c12bfaeabe7b81645400000000000000010100014cb200000080a1f32002000000000000000000000000211201" +
        "08110100000000";

    private const string LiveClubChanges =
        "F10502000028A0632B72AA130000000200028A010106424E45546363E32BFAEA3ED7E9501400000000000000000100014CF244B2A5039F070962000000000005B08E1000028A" +
        "61D19654AF013B47726F7570206368617420666F7220424E45542E6363277320776562736974652E20436865636B206974206F75743A207777772E424E45542E63633F7FE200" +
        "00000000030000000000CFB903";

    private const string Invitation = "f6050002991201000000010000000006acf90600415669560000";
}

public class ChatMessageLimitTests
{
    [Fact]
    public void CutsAtTwoHundredFiftyFiveCharacters()
    {
        Assert.Equal("short", Invigoration.Sc2.Native.ChatCommands.CutToMessage("short"));
        Assert.Equal(255, Invigoration.Sc2.Native.ChatCommands.CutToMessage(new string('a', 400)).Length);
        Invigoration.Sc2.Native.ChatCommands.ChatMessage(0, Invigoration.Sc2.Native.ChatCommands.CutToMessage(new string('é', 400)));
    }
}
