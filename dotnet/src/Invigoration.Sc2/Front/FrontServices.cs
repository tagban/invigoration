namespace Invigoration.Sc2.Front;

/// <summary>Fully-qualified bgs.protocol service names, hashed via <see cref="Wire.ServiceHash"/> for the Front Header's service_hash field. Source: https://superioritybot.com/PROTOCOL's Front RPC section.</summary>
public static class FrontServices
{
    public const string Connection = "bnet.protocol.connection.ConnectionService";
    public const string AuthenticationServer = "bnet.protocol.authentication.AuthenticationServer";
    public const string AuthenticationClient = "bnet.protocol.authentication.AuthenticationClient";
    public const string ChallengeNotify = "bnet.protocol.challenge.ChallengeNotify";
    public const string GameUtilities = "bnet.protocol.game_utilities.GameUtilities";
    public const string Account = "bnet.protocol.account.AccountService";

    // Original (pre-v1) names, which the service_hash is computed from. Verified against the
    // OriginalHash constants in TrinityCore's generated friends_service.pb.h (0xA3DDB1BD,
    // 0x6F259A13), presence_service.pb.h (0xFA0796FF), presence_listener.pb.h (0x890AB85F) and the
    // pre-2018-11 channel_service.pb.h (0xBF8C8094).
    public const string Friends = "bnet.protocol.friends.FriendsService";
    public const string FriendsListener = "bnet.protocol.friends.FriendsNotify";
    public const string Presence = "bnet.protocol.presence.PresenceService";
    public const string PresenceListener = "bnet.protocol.presence.v1.PresenceListener";
    public const string ChannelListener = "bnet.protocol.channel.ChannelSubscriber";

    // The same listeners under their current package names (TrinityCore's NameHash). The server
    // addresses listeners by original name as far as anyone has seen, but both are accepted.
    public const string FriendsListenerV1 = "bgs.protocol.friends.v1.FriendsListener";
    public const string PresenceListenerV1 = "bgs.protocol.presence.v1.PresenceListener";
    public const string ChannelListenerV1 = "bgs.protocol.channel.v1.ChannelListener";
}
