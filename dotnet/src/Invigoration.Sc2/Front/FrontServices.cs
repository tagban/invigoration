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
}
