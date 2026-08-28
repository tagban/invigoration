namespace Invigoration.Core.Hotline;

/// <summary>
/// A saved Hotline server connection — the "profile" the user asked for ("Each server connecting
/// should be a 'profile' and have the option to auto-connect"). See HotlineServerProfileStore.
/// </summary>
public sealed class HotlineServerProfile
{
    /// <summary>Stable identity, assigned once at creation — same pattern as BattlenetCredentialProfile.Id.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string Name { get; set; } = "New Server";

    public string Host { get; set; } = "";

    public ushort Port { get; set; } = HotlineConstants.DefaultServerPort;

    /// <summary>Empty means log in anonymously (a real Hotline client with no account still needs *some* login name — most servers accept a blank one).</summary>
    public string Login { get; set; } = "";

    public string Password { get; set; } = "";

    public string Nickname { get; set; } = "Guest";

    public ushort IconId { get; set; } = 414;

    /// <summary>Connect this server's tab automatically when the Hotline tab group is opened, instead of waiting for the user to pick it from the tracker or a manual "Connect" click.</summary>
    public bool AutoConnect { get; set; }

    /// <summary>Never auto-accept a server's agreement by default — per-server, not per-tracker (different servers under the same tracker can have very different rules to actually read). See HotlineSessionViewModel's Agreement handling.</summary>
    public bool AutoAcceptAgreement { get; set; }

    /// <summary>
    /// This server's Discord relay bot's own Hotline username (e.g. "Relay", "Discord" — genuinely
    /// different per server, confirmed live) — empty disables relay detection. Per-server, not
    /// per-tracker: two servers under the same tracker can run completely different relay bots.
    /// See HotlineSessionViewModel.TryAppendDiscordRelayMessage's remarks.
    /// </summary>
    public string DiscordRelayUsername { get; set; } = "";

    /// <summary>An optional literal prefix before "{DiscordUser}: {message}" in this server's relay messages (e.g. "Discord | ") — confirmed live this varies by server; empty means no prefix.</summary>
    public string DiscordRelayPrefix { get; set; } = "";

    /// <summary>
    /// The VersionNumber advertised at login — 6112 by default, per explicit request: the
    /// VersionNumber field reveals a lot about the connecting client (see the modern-server
    /// protocol docs at github.com/fogWraith/Hotline/tree/main/Docs/Protocol), so this is
    /// deliberately a distinctive, unused-by-any-real-client number chosen specifically to
    /// identify Invigoration's own connections as itself, not a claim about which real Hotline
    /// version's feature set this client speaks (contrast the old 150/1.5.x-honesty reasoning,
    /// superseded by this request). Per-server, not per-tracker: different real servers react
    /// differently to a claimed version, and per explicit request, a newer Hotline server variant
    /// expects the field not to be sent at all — see SendClientVersion.
    /// </summary>
    public ushort ClientVersion { get; set; } = 6112;

    /// <summary>Off skips the VersionNumber field entirely instead of sending ClientVersion — per explicit request, for a real newer Hotline server build that expects no VersionNumber field at all (not just a specific value). On by default since most servers, old and new, do expect one.</summary>
    public bool SendClientVersion { get; set; } = true;

    /// <summary>
    /// Off by default, per explicit request. Gates "/trivia" on this server the same hard way
    /// Config.TriviaFeatureEnabled already gates it for a Battle.net bot — see
    /// BotEngineTriviaToggleTests' "FeatureDisabled" case, which this mirrors. Per-server, not
    /// per-tracker: two servers under the same tracker can have very different chat cultures.
    /// A Hotline session never joins a cross-server/cross-protocol trivia group either way — a
    /// round only ever plays out on the one server it was started on (see
    /// HotlineTriviaHost.AnnounceStartedAsync/BroadcastAsync).
    /// </summary>
    public bool TriviaEnabled { get; set; }

    /// <summary>
    /// Off by default. Advertises CAPABILITY_CHAT_HISTORY (DATA_CAPABILITIES bit 4) at login — a
    /// server that confirms it lets this client pull the last 20 messages via the real "2.5"-era
    /// chat history extension (HotlineTransactionClient.GetChatHistoryAsync) instead of starting
    /// on a blank screen. Kept opt-in per-server rather than always-on: this is a brand-new field
    /// most real servers have never seen, and intermittent forced disconnects started appearing on
    /// at least one real server right around when this was first added — TLV framing means an
    /// unrecognized field should be safely skippable, but that's unproven against the wide range
    /// of real server implementations out there, so it's not worth risking connection stability
    /// for a nice-to-have nobody explicitly asked to always have on.
    /// </summary>
    public bool AdvertiseChatHistorySupport { get; set; }
}
