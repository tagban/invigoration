namespace Invigoration.App.ViewModels;

/// <summary>
/// Everything needed to connect one Hotline session — a plain options bag rather than a long
/// positional parameter list, now that AutoAcceptAgreement, the Discord relay settings, and the
/// ClientVersion/SendClientVersion pair all moved from tracker-level to per-server
/// (HotlineServerProfile) config: a connect from a saved profile carries all of it; an ad-hoc
/// connect straight from the tracker's server list only carries the tracker-level defaults
/// (Nickname/IconId) and leaves the rest at their safe defaults (never auto-accept, no relay
/// detection, the standard 6112 version identifying Invigoration itself — see
/// HotlineServerProfile.ClientVersion's remarks) — matching that those settings only exist once a
/// server is actually saved as a profile.
/// </summary>
public sealed record HotlineConnectOptions(
    string Host,
    int Port,
    string Login,
    string Password,
    string Nickname,
    ushort IconId,
    string? DisplayName = null,
    bool AutoAcceptAgreement = false,
    string DiscordRelayUsername = "",
    string DiscordRelayPrefix = "",
    ushort ClientVersion = 6112,
    bool SendClientVersion = true,
    string? ProfileId = null,
    bool TriviaEnabled = false,
    bool AdvertiseChatHistorySupport = false);
