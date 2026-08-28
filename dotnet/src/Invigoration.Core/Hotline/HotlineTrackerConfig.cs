namespace Invigoration.Core.Hotline;

/// <summary>
/// One top-level "Add Bot"-style Hotline entity — the user's own framing: "Hotline should be
/// treated like add bot... each hotline connection is a 'tracker'." Each of these gets its own
/// top-level tab (HotlineTabViewModel) holding one tracker/server-browser view plus whatever
/// servers are currently connected under it. Deliberately separate from
/// <see cref="HotlineServerProfile"/> — a profile is one saved server to connect *to*; this is the
/// whole browsing session/tab it's connected *from*. Per-server settings that genuinely vary
/// server-to-server (AutoAcceptAgreement, Discord relay identity) live on the profile instead —
/// see HotlineServerProfile's remarks.
/// </summary>
public sealed class HotlineTrackerConfig
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    public string DisplayName { get; set; } = "Hotline";

    public string TrackerHost { get; set; } = "hltracker.com";

    /// <summary>The nickname/icon used connecting straight from the tracker's server list — "to use on all servers" per the user's own request, distinct from a saved HotlineServerProfile's own Nickname/IconId, which still take precedence for a profile-based connect.</summary>
    public string DefaultNickname { get; set; } = "Guest";

    public ushort DefaultIconId { get; set; } = 414;

    /// <summary>Off by default — logs every inbound transaction (type + fields) to a session's chat log, to diagnose exactly what a server sends right before an unexplained disconnect. Grouped under the tracker settings' collapsed "Advanced" section.</summary>
    public bool Debug { get; set; }

    /// <summary>Off by default — the "Copy Log" button on a session's top bar is a niche/debugging tool most people never need; shown only once explicitly opted into here, so it doesn't take up UI space by default.</summary>
    public bool ShowCopyLogButton { get; set; }

    /// <summary>
    /// Chat username colors, by rank — 2-tier only, matching what the protocol actually broadcasts
    /// for other users (just the single Admin bit in UserFlags; see HotlineAccessBits' remarks for
    /// why a genuine 3rd "Mod" tier isn't derivable client-side without a locally-maintained
    /// roster, which doesn't exist yet). Classic Hotline convention: purple for admins, and white
    /// (not the traditional black) since this app's chat log has a dark background.
    /// </summary>
    public string AdminColorHex { get; set; } = "#C77DFF";

    public string DefaultColorHex { get; set; } = "#FFFFFF";

    /// <summary>Per-username chat highlight color overrides (username → hex), set via the Users list's right-click menu — takes priority over the Admin/Default rank colors above, so a specific person's messages stand out regardless of their rank.</summary>
    public Dictionary<string, string> UserHighlightColors { get; set; } = [];

    /// <summary>
    /// Whether the tracker's own setup fields (name/host/identity/chat colors/advanced — everything
    /// except the two server lists) are shown expanded. True by default so a brand-new tracker
    /// starts open for configuring; HotlineTrackerViewModel.RefreshTracker auto-collapses this the
    /// first time it successfully pulls a non-empty server list, per explicit request ("settings
    /// for the tracker should collapse down once the tracker is setup for the first time"), leaving
    /// just the two server lists visible day-to-day. Persisted, and freely re-toggleable afterward.
    /// </summary>
    public bool SettingsExpanded { get; set; } = true;
}
