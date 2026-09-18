using Stimpak;

namespace Invigoration.Core;

/// <summary>
/// Whether this bot is doing anything at all — connected, part-way through connecting, or waiting
/// to reconnect — so a UI can offer Connect only when there's nothing it would start over the top
/// of, and Disconnect whenever there's something to stop.
/// </summary>
/// <remarks>
/// <para>Worked out from live state every time it's asked rather than counted up and down by
/// events: plenty of ways an attempt ends raise nothing (a BNLS server going quiet, a refused TCP
/// connect, an SC2 sign-in that never finishes), and a count that misses one of those stays wrong
/// for good. ActivityChanged only says "look again" — it carries no value, so a late or repeated
/// notification can't leave a UI showing something stale.</para>
/// <para>The one state no socket shows is a classic logon between BNLS answering and Battle.net
/// itself connecting — BNLS alone is up, and it stays up after logon too, so it can't count on its
/// own. _logonPending covers exactly that stretch.</para>
/// </remarks>
public sealed partial class BotEngine
{
    /// <summary>TCP connects currently being waited on — a connect to a dead host can take the OS's whole timeout with no socket up to show for it.</summary>
    private int _connectsInFlight;

    /// <summary>From a classic (BNLS) logon starting until it's logged on, dropped, abandoned or disconnected.</summary>
    private volatile bool _logonPending;

    /// <summary>Counts classic logons started, so a late report of an older connection closing can't end a newer logon (see the BNCS Disconnected handler).</summary>
    private int _logonAttempt;

    /// <summary>Which logon the current Battle.net connection was opened for.</summary>
    private int _bncsLogonAttempt;

    /// <summary>The Stimpak client whose session is still live — not the same as _sc2Client, which outlives a session that ended.</summary>
    private StimpakClient? _sc2LiveClient;

    /// <summary>Cancelled by DisconnectAsync (and replaced), so a connect still waiting on its socket stops instead of coming up after the user said stop.</summary>
    private CancellationTokenSource _connectCts = new();

    /// <summary>Something about IsIdle, IsReconnecting or IsWaitingToReconnect may have changed. Raised on whatever thread the change happened on.</summary>
    public event Action? ActivityChanged;

    /// <summary>An automatic reconnect is counting down or trying — see BotEngine.Reconnect.cs.</summary>
    public bool IsReconnecting => Volatile.Read(ref _reconnectRunning) == 1;

    /// <summary>Nothing connected, nothing connecting, no reconnect waiting: connecting now wouldn't collide with anything.</summary>
    public bool IsIdle => !IsReconnecting && NothingUnderWay;

    /// <summary>
    /// An automatic reconnect is only waiting out the delay between attempts — nothing up and
    /// nothing in flight. Connecting now is fine (ConnectAsync cancels the countdown), which is how
    /// someone who knows the server's back skips the wait.
    /// </summary>
    public bool IsWaitingToReconnect => IsReconnecting && NothingUnderWay;

    private bool NothingUnderWay =>
        Volatile.Read(ref _connectsInFlight) == 0 &&
        !_logonPending &&
        !_bncs.IsConnected &&
        !_chatTelnet.IsConnected &&
        Volatile.Read(ref _sc2LiveClient) is null;

    private void RaiseActivityChanged() => ActivityChanged?.Invoke();

    /// <summary>Subscribed after every other handler on these sockets, so a drop's auto-reconnect is already scheduled by the time anyone looks.</summary>
    private void WireActivity()
    {
        _bncs.Connected += RaiseActivityChanged;
        _bncs.Disconnected += _ => RaiseActivityChanged();
        _bnls.Disconnected += _ => RaiseActivityChanged();
        _chatTelnet.Connected += RaiseActivityChanged;
        _chatTelnet.Disconnected += _ => RaiseActivityChanged();
    }

    /// <summary>Runs one socket connect counted as in flight, and cancellable by DisconnectAsync as well as by the caller.</summary>
    private async Task TrackConnectAsync(Func<CancellationToken, Task> connect, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, Volatile.Read(ref _connectCts).Token);
        Interlocked.Increment(ref _connectsInFlight);
        RaiseActivityChanged();
        try
        {
            await connect(linked.Token).ConfigureAwait(false);
        }
        finally
        {
            Interlocked.Decrement(ref _connectsInFlight);
            RaiseActivityChanged();
        }
    }

    /// <summary>A classic logon that can't go any further — nothing else would clear _logonPending for it.</summary>
    private void EndPendingLogon()
    {
        _logonPending = false;
        RaiseActivityChanged();
    }

    /// <summary>Marks a Stimpak session over — only if it's still the current one, so an older client winding down can't make a newer connection look idle.</summary>
    private void EndSc2Activity(StimpakClient client)
    {
        if (Interlocked.CompareExchange(ref _sc2LiveClient, null, client) == client)
        {
            RaiseActivityChanged();
        }
    }
}
