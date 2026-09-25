namespace Invigoration.Sc2.Bsn;

/// <summary>
/// The wire layouts of obfuscated club structs: which member is sent when, and the filler bits
/// before it. Each was recovered from the SC2 5.0.16.97563 client's generated code or matched
/// against a retail capture; the source is on each line.
/// </summary>
public static class ClubLayouts
{
    public static IReadOnlyDictionary<string, BsnLayout> Known { get; } = new Dictionary<string, BsnLayout>
    {
        // ncarrillo/superiority (MIT), core/src/games/sc2/native/wire_layout.rs:
        ["Battlenet::Toon::Handle"] = new((1, 0), (0, 0), (2, 0), (3, 0)),
        ["Battlenet::Club::InviteAction"] = new((2, 0), (1, 0), (0, 0), (3, 11)),
        ["Battlenet::Client::Club::InviteAction"] = new((0, 0)),
        ["Battlenet::Club::SubscriptionSyncInfo"] = new((2, 11), (0, 0), (1, 0)),
        ["Battlenet::Club::ClubChangeInfo"] = new((0, 6), (3, 0), (1, 0), (2, 0)),
        ["Battlenet::Client::Club::ClubSubscribeRequest"] = new((0, 0)),

        // Superiority's generated ClubSummaryInfo reader (read_club_summary_info): locale, member
        // count, id, category, name, 6 filler, record address, flags, type, program, created, tag,
        // 25 filler, file handles.
        ["Battlenet::Club::ClubSummaryInfo"] = new((6, 0), (10, 0), (0, 0), (5, 0), (2, 0), (9, 6), (7, 0), (4, 0), (1, 0), (11, 0), (3, 0), (8, 25)),

        // The rest were read from the SC2 5.0.16.97563 client's generated writers, readers and size
        // functions (bit-write widths and filler calls in order), after the same method reproduced
        // every layout above. Base structs hold only the token.
        ["Battlenet::Client::Club::GetToonClubs"] = new((0, 0)),
        ["Battlenet::Client::Club::SetMemberRank"] = new((0, 0)),
        ["Battlenet::Client::Club::CreateClub"] = new((0, 0)),
        ["Battlenet::Client::Club::GetRoster"] = new((0, 0)),
        ["Battlenet::Client::Club::ModifyClub"] = new((0, 0)),
        ["Battlenet::Client::Club::ModifyMember"] = new((0, 0)),
        ["Battlenet::Client::Club::GetMemberClanTags"] = new((0, 0)),
        ["Battlenet::Client::Club::CreateClubV2"] = new((0, 26)),

        // Our character's clubs. The request also matches the one retail sends; the reply matches
        // two retail replies.
        ["Battlenet::Client::Club::GetToonClubsRequest"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::GetToonClubsResponse"] = new((1, 0), (0, 31)),
        ["Battlenet::Club::ClubInfo"] = new((0, 0), (1, 0)),
        ["Battlenet::Club::ClubOnlineStatus"] = new((2, 0), (1, 0), (0, 0)),
        ["Battlenet::Client::Club::GetClubInfoRequest"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::GetClubInfoResponse"] = new((1, 0), (0, 0)),

        // Ranks: old rank, new rank, token, club, the member's full name, their character id.
        ["Battlenet::Client::Club::SetMemberRankRequest"] = new((4, 0), (5, 0), (0, 0), (1, 0), (2, 0), (3, 0)),
        ["Battlenet::Client::Club::SetMemberRankResponse"] = new((1, 0), (0, 3)),

        // The member list, and changes to it.
        ["Battlenet::Client::Club::GetRosterRequest"] = new((1, 0), (2, 0), (0, 0)),
        ["Battlenet::Client::Club::GetRosterResponse"] = new((1, 0), (0, 0)),
        ["Battlenet::Club::MemberSummaryInfo"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::MemberChangeNotification"] = new((0, 0)),
        ["Battlenet::Club::MemberChangeInfo"] = new((4, 0), (0, 5), (2, 29), (1, 0), (3, 0)),
        ["Battlenet::Client::Club::ModifyMemberRequest"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::ModifyMemberResponse"] = new((1, 0), (0, 0)),

        // Creating a club: flags, 14 filler, token, type, tag, category, name, locale.
        ["Battlenet::Client::Club::CreateClubRequest"] = new((6, 0), (0, 14), (3, 0), (2, 0), (4, 0), (1, 0), (5, 0)),
        ["Battlenet::Client::Club::CreateClubResponse"] = new((0, 32), (1, 0)),
        ["Battlenet::Client::Club::CreateClubV2Request"] = new((0, 0), (1, 0)),
        ["Battlenet::Club::ClubCreationInfo"] = new((7, 13), (3, 0), (0, 0), (4, 0), (1, 0), (5, 0), (2, 0), (6, 0)),
        ["Battlenet::Client::Club::CreateClubV2Response"] = new((2, 0), (1, 0), (0, 0)),

        // Changing a club (name, tag, message...), settings, clan tags, change notices.
        ["Battlenet::Client::Club::ModifyClubRequest"] = new((1, 0), (0, 0)),
        ["Battlenet::Client::Club::ModifyClubResponse"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::ClubSettings"] = new((2, 0), (1, 0), (0, 0), (3, 0)),
        ["Battlenet::Client::Club::GetMemberClanTagsRequest"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Club::GetMemberClanTagsResponse"] = new((1, 24), (0, 0), (3, 0), (2, 0)),
        ["Battlenet::Client::Club::ClubChangeNotification"] = new((0, 0)),
        ["Battlenet::Club::ClubUserText"] = new((0, 0), (3, 0), (2, 0), (1, 0)),
        ["Battlenet::Club::ClubUserTextSimple"] = new((0, 0)),
        ["Battlenet::Club::ClubEvent"] = new((3, 0), (5, 0), (1, 0), (4, 0), (0, 0), (2, 0)),
        ["Battlenet::Club::ClubEventSimple"] = new((0, 0), (2, 0), (1, 0), (3, 0)),

        // Profile slot 14, command 2 both ways: character handles to names. 15 filler bits
        // lead the reply's list; each entry is sent tag, handle, name, result.
        ["Battlenet::Client::Profile::ResolveToonHandleToName"] = new(),
        ["Battlenet::Client::Profile::ResolveToonHandleToNameRequest"] = new((0, 0), (1, 0)),
        ["Battlenet::Client::Profile::ResolveToonHandleToNameResponse"] = new((0, 0), (1, 15)),
        ["Battlenet::Client::Profile::HandleToNameResponse"] = new((3, 0), (1, 0), (2, 0), (0, 0)),
    };

    public static BsnCodec Codec { get; } = new(ClubSchema.Instance, Known);
}
