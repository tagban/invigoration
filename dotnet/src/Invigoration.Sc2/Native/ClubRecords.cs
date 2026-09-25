using Invigoration.Sc2.Bsn;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>A club's rank values (Battlenet::Club::MemberRank).</summary>
public static class ClubRanks
{
    public const byte NotSet = 0;
    public const byte Banned = 5;
    public const byte Visitor = 10;
    public const byte Invited = 20;
    public const byte Member = 30;
    public const byte Officer = 40;
    public const byte Owner = 50;

    public static string Describe(byte rank) => rank switch
    {
        Owner => "Owner",
        Officer => "Officer",
        Member => "Member",
        25 => "Honorary",
        24 => "Benchwarmer",
        Invited => "Invited",
        18 => "Suggested",
        16 => "Requested",
        13 => "Rejected",
        Visitor => "Visitor",
        Banned => "Banned",
        _ => $"Rank {rank}",
    };
}

/// <summary>What Battle.net says about a club (Battlenet::Club::ClubSummaryInfo).</summary>
/// <param name="Type">1 group, 2 clan, 3 team.</param>
public sealed record ClubSummary(uint Id, string Name, string? Tag, byte Type, byte Category, uint Flags, uint MemberCount, string Program)
{
    public const byte GroupType = 1;
    public const byte ClanType = 2;

    public bool IsClan => Type == ClanType;
}

/// <summary>One of our character's clubs, with our rank in it and how many members are online.</summary>
public sealed record ToonClub(ClubSummary Summary, byte Rank, uint Online, uint InGame, uint InChat);

/// <summary>GetToonClubsResponse (club slot 13, command 46): our character's clubs, or an error code.</summary>
public sealed record ToonClubsRecord(IReadOnlyList<ToonClub> Clubs, bool IsLastPacket, ushort? Error);

/// <summary>Club InviteAction (club slot 13, command 54): an invitation to a club, or someone's answer to one.</summary>
/// <param name="Code">0 invited, 1 accepted, 2 declined.</param>
public sealed record ClubInviteRecord(uint ClubId, byte Code, ushort Result);

/// <summary>ClubSettings (club slot 13, command 57, at sign-in): the patterns a club's name and tag must match.</summary>
public sealed record ClubSettingsRecord(string NameRegex, string TagRegex);

/// <summary>One change to a club's members (MemberChangeInfo): joined, updated or removed, with a new rank or status.</summary>
/// <param name="ChangeType">0 insert, 1 update, 2 remove, 3 sync.</param>
/// <param name="OldRank">For a rank change; null for a status change.</param>
/// <param name="Status">For a status change: 0 offline, 1 online, 2 in a game, 3 in chat; null for a rank change.</param>
public sealed record ClubMemberChange(uint ClubId, ToonHandleValue Member, byte ChangeType, byte? OldRank, byte? NewRank, byte? Status);

/// <summary>MemberChangeNotification (club slot 13, command 50).</summary>
public sealed record ClubMemberChangesRecord(IReadOnlyList<ClubMemberChange> Changes);

/// <summary>One change to a club itself (ClubChangeInfo): which part changed, and its text if it's a description, announcement or message.</summary>
/// <param name="Part">The ClubChangeInfo::Info variant, e.g. summaryInfoFull, Description, Announcement.</param>
public sealed record ClubChange(uint ClubId, byte ChangeType, string Part, string? Text);

/// <summary>ClubChangeNotification (club slot 13, command 49): changes to clubs we follow, and a full sync right after subscribing.</summary>
public sealed record ClubChangesRecord(IReadOnlyList<ClubChange> Changes);

/// <summary>GetRosterResponse (club slot 13, command 45): one page of a club's members, by character handle, with ranks.</summary>
/// <param name="Token">The request's token, handed back: the club it was for, when the request used the club id.</param>
public sealed record ClubRosterRecord(uint Token, IReadOnlyList<(ToonHandleValue Member, byte Rank)> Members, bool IsLastPacket, ushort? Error);

/// <summary>ResolveToonHandleToNameResponse (profile slot 14, command 2): character names, and clan tags, for handles.</summary>
public sealed record ToonNamesRecord(IReadOnlyList<(ToonHandleValue Handle, string? Name, string? ClanTag, ushort Result)> Names);

/// <summary>
/// Club records, by SC2's own schema and wire layouts (<see cref="ClubLayouts"/>). The club
/// service is slot 13; its commands are Battlenet::Client::Club::CommandID_S2Map (41 to 60).
/// </summary>
public static class ClubCommands
{
    public const byte ClubSlot = 13;
    public const byte GetToonClubsCommand = 46;
    public const byte InviteActionCommand = 54;
    public const byte GetRosterCommand = 45;
    public const byte SubscribeCommand = 47;

    /// <summary>SubscriptionType ROSTER: member changes, including online status.</summary>
    public const byte RosterSubscription = 7;
    public const byte ClubChangeNotificationCommand = 49;
    public const byte MemberChangeNotificationCommand = 50;
    public const byte ProfileSlot = 14;
    public const byte ResolveToonNamesCommand = 2;

    /// <summary>ResolveToonHandleToNameRequest takes at most this many handles (a 6-bit count).</summary>
    public const int MaxNamesPerRequest = 32;
    public const byte ClubSettingsCommand = 57;

    private static BsnCodec Codec => ClubLayouts.Codec;

    /// <summary>GetToonClubsRequest: the clubs a character belongs to. Matches the request retail sends.</summary>
    public static byte[] GetToonClubs(uint token, ToonHandleValue toon)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, GetToonClubsCommand, ClubSlot);
        Codec.Encode(writer, "Battlenet::Client::Club::GetToonClubsRequest", Codec.Struct(
            "Battlenet::Client::Club::GetToonClubsRequest",
            ("GetToonClubs", Codec.Struct("Battlenet::Client::Club::GetToonClubs", ("m_token", token))),
            ("m_toon", Handle(toon))));
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>GetRosterRequest from <paramref name="offset"/>: the members, 200 at most per reply page.</summary>
    public static byte[] GetRoster(uint token, uint clubId, uint offset)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, GetRosterCommand, ClubSlot);
        Codec.Encode(writer, "Battlenet::Client::Club::GetRosterRequest", Codec.Struct(
            "Battlenet::Client::Club::GetRosterRequest",
            ("GetRoster", Codec.Struct("Battlenet::Client::Club::GetRoster", ("m_token", token))),
            ("m_clubId", clubId),
            ("m_param", new BsnChoice(0, "offset", offset))));
        writer.Align();
        return writer.ToBytes();
    }

    public static ClubRosterRecord DecodeRoster(BitReader reader)
    {
        var reply = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::GetRosterResponse")!;
        reader.Align();
        var token = (uint)(long)((BsnStruct)reply["GetRoster"]!)["m_token"]!;
        var result = (BsnChoice)reply["m_result"]!;
        if (result.Value is not BsnStruct success)
        {
            return new ClubRosterRecord(token, [], true, (ushort)Convert.ToInt64(result.Value));
        }

        var members = ((List<object?>)success["m_memberInfo"]!).Cast<BsnStruct>()
            .Select(m => (HandleOf((BsnStruct)m["m_handle"]!), (byte)(long)m["m_rank"]!))
            .ToList();
        return new ClubRosterRecord(token, members, (bool)success["m_isLastPacket"]!, null);
    }

    /// <summary>
    /// ClubSubscribeRequest (command 47): member changes for these clubs, as MemberChangeNotification
    /// (command 50), status (online, in a game, in chat) included. Stamp 0: from now.
    /// </summary>
    public static byte[] SubscribeToRosters(IEnumerable<uint> clubIds)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, SubscribeCommand, ClubSlot);
        Codec.Encode(writer, "Battlenet::Client::Club::ClubSubscribeRequest", Codec.Struct(
            "Battlenet::Client::Club::ClubSubscribeRequest",
            ("m_subscriptions", clubIds.Select(id => (object?)Codec.Struct(
                "Battlenet::Club::SubscriptionSyncInfo",
                ("m_clubId", id),
                ("m_type", RosterSubscription),
                ("m_stamp", 0))).ToList())));
        writer.Align();
        return writer.ToBytes();
    }

    /// <summary>ResolveToonHandleToNameRequest (profile slot 14, command 2) for up to <see cref="MaxNamesPerRequest"/> characters.</summary>
    public static byte[] ResolveToonNames(IReadOnlyList<ToonHandleValue> handles)
    {
        var writer = new BitWriter();
        RoutingHeader.Encode(writer, ResolveToonNamesCommand, ProfileSlot);
        Codec.Encode(writer, "Battlenet::Client::Profile::ResolveToonHandleToNameRequest", Codec.Struct(
            "Battlenet::Client::Profile::ResolveToonHandleToNameRequest",
            ("ResolveToonHandleToName", Codec.Struct("Battlenet::Client::Profile::ResolveToonHandleToName")),
            ("m_handles", handles.Select(h => (object?)Handle(h)).ToList())));
        writer.Align();
        return writer.ToBytes();
    }

    public static ToonNamesRecord DecodeToonNames(BitReader reader)
    {
        var reply = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Profile::ResolveToonHandleToNameResponse")!;
        reader.Align();
        var names = ((List<object?>)reply["m_responses"]!).Cast<BsnStruct>()
            .Select(r => (
                HandleOf((BsnStruct)r["m_handle"]!),
                (r["m_name"] as BsnStruct)?["m_name"] as string,
                r["m_tag"] as string,
                (ushort)(long)r["m_result"]!))
            .ToList();
        return new ToonNamesRecord(names);
    }

    public static ToonClubsRecord DecodeToonClubs(BitReader reader)
    {
        var reply = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::GetToonClubsResponse")!;
        reader.Align();
        var result = (BsnChoice)reply["m_result"]!;
        if (result.Value is not BsnStruct success)
        {
            return new ToonClubsRecord([], true, (ushort)Convert.ToInt64(result.Value));
        }

        var infos = (List<object?>)success["m_clubInfo"]!;
        var ranks = (List<object?>)success["m_rankInfo"]!;
        var clubs = new List<ToonClub>(infos.Count);
        for (var i = 0; i < infos.Count; i++)
        {
            var info = (BsnStruct)infos[i]!;
            var status = (BsnStruct)info["m_status"]!;
            clubs.Add(new ToonClub(
                Summary((BsnStruct)info["m_summary"]!),
                i < ranks.Count ? (byte)(long)ranks[i]! : ClubRanks.NotSet,
                (uint)(long)status["m_online"]!,
                (uint)(long)status["m_ingame"]!,
                (uint)(long)status["m_inchat"]!));
        }

        return new ToonClubsRecord(clubs, (bool)success["m_isLastPacket"]!, null);
    }

    public static ClubInviteRecord DecodeInviteAction(BitReader reader)
    {
        var action = (BsnStruct)((BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::InviteAction")!)["m_action"]!;
        reader.Align();
        return new ClubInviteRecord((uint)(long)action["m_clubId"]!, (byte)(long)action["m_code"]!, (ushort)(long)action["m_result"]!);
    }

    public static ClubChangesRecord DecodeClubChanges(BitReader reader)
    {
        var notification = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::ClubChangeNotification")!;
        reader.Align();
        var changes = new List<ClubChange>();
        foreach (BsnStruct delta in (List<object?>)notification["m_deltas"]!)
        {
            var info = (BsnChoice)delta["m_info"]!;
            changes.Add(new ClubChange((uint)(long)delta["m_clubId"]!, (byte)(long)delta["m_changeType"]!, info.Name, TextOf(info.Value)));
        }

        return new ClubChangesRecord(changes);
    }

    /// <summary>The text of a user-text or event change (description, announcement, message), if it has one.</summary>
    private static string? TextOf(object? value) => value switch
    {
        BsnStruct s when s.Has("m_text") => s["m_text"] as string,
        BsnStruct s when s.Has("m_title") => s["m_title"] as string,
        _ => null,
    };

    public static ClubSettingsRecord DecodeClubSettings(BitReader reader)
    {
        var settings = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::ClubSettings")!;
        reader.Align();
        return new ClubSettingsRecord((string)settings["m_clubNameRegEx"]!, (string)settings["m_clubTagRegEx"]!);
    }

    public static ClubMemberChangesRecord DecodeMemberChanges(BitReader reader)
    {
        var notification = (BsnStruct)Codec.Decode(reader, "Battlenet::Client::Club::MemberChangeNotification")!;
        reader.Align();
        var changes = new List<ClubMemberChange>();
        foreach (BsnStruct delta in (List<object?>)notification["m_deltas"]!)
        {
            var info = (BsnChoice)delta["m_info"]!;
            var rank = info.Value as BsnStruct;
            changes.Add(new ClubMemberChange(
                (uint)(long)delta["m_clubId"]!,
                HandleOf((BsnStruct)delta["m_member"]!),
                (byte)(long)delta["m_changeType"]!,
                rank is null ? null : (byte)(long)rank["m_oldRank"]!,
                rank is null ? null : (byte)(long)rank["m_newRank"]!,
                rank is null ? (byte)(long)info.Value! : null));
        }

        return new ClubMemberChangesRecord(changes);
    }

    private static ToonHandleValue HandleOf(BsnStruct handle) => new(
        (uint)handle["m_programId"]!,
        (byte)(long)handle["m_region"]!,
        (uint)(long)handle["m_realm"]!,
        unchecked((ulong)(long)handle["m_id"]!));

    private static ClubSummary Summary(BsnStruct summary) => new(
        (uint)(long)summary["m_id"]!,
        (string)summary["m_name"]!,
        summary["m_tag"] as string,
        (byte)(long)summary["m_type"]!,
        (byte)(long)summary["m_category"]!,
        (uint)(long)summary["m_flags"]!,
        (uint)(long)summary["m_memberCount"]!,
        FourCc.Decode((uint)summary["m_program"]!));

    private static BsnStruct Handle(ToonHandleValue toon) => Codec.Struct(
        "Battlenet::Toon::Handle",
        ("m_region", toon.Region),
        ("m_programId", toon.Program),
        ("m_realm", toon.Realm),
        ("m_id", unchecked((long)toon.Id)));
}
