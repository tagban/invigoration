namespace Invigoration.Core.Tracking;

/// <summary>
/// One locally-cached chat line — kept purely so reconnecting to a server shows roughly "where the
/// conversation was last at," for a server/protocol with no server-side memory of its own. Already
/// fully-formatted display text, not raw wire bytes — this is a client-side convenience cache, not
/// a protocol replay log. See RecentMessageStore.
/// </summary>
public sealed class RecentMessage
{
    public string Text { get; set; } = "";

    /// <summary>Null for a line with no reliable timestamp — not every protocol/line carries one.</summary>
    public DateTimeOffset? TimestampUtc { get; set; }
}
