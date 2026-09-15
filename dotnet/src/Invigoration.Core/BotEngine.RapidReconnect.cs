using System.Diagnostics;
using System.Net.Sockets;
using Invigoration.Core.Auth;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>
/// Rapid reconnect (BotConfig.RapidReconnect): for private servers where getting back on first
/// matters — holding a name, or ops in a channel. The moment the connection drops it tries to
/// connect, then tries again every second until the bot is logged back on.
/// </summary>
/// <remarks>
/// Each attempt skips BNLS entirely when it can: the CD key and password are hashed locally, and
/// the two answers only BNLS computes (the version byte, and the version check for the server's
/// challenge) are reused from an earlier logon this run (LogonCheckCache), so the whole logon is
/// just the Battle.net server itself — measured at roughly half the time of a logon that asks
/// BNLS. When it can't (Warcraft III's logon, a CD key only BNLS decodes, nothing cached yet, or a
/// server that changed its challenge), the attempt logs on the normal way, still every second.
/// Never used on official Battle.net, where a connection every second is how an address gets
/// banned; the ordinary auto-reconnect still applies there if it's on.
/// </remarks>
public sealed partial class BotEngine
{
    private static readonly TimeSpan RapidRetryInterval = TimeSpan.FromSeconds(1);

    /// <summary>An attempt that connected but still isn't logged on after this long is given up for a fresh one.</summary>
    private static readonly TimeSpan RapidAttemptTimeout = TimeSpan.FromSeconds(10);

    private int _rapidReconnectRunning;

    public static bool IsOfficialBattlenetServer(string server) =>
        BncsProduct.OfficialBattlenetServers.Contains(server.Trim(), StringComparer.OrdinalIgnoreCase);

    /// <summary>Whether a drop right now would be handled by rapid reconnect: on, a classic game, and not official Battle.net.</summary>
    public bool RapidReconnectApplies =>
        Config.RapidReconnect && !BncsProduct.IsStimpakBacked(Config.Product) && !IsOfficialBattlenetServer(Config.BattlenetServer);

    /// <summary>Whether this bot can log on without asking BNLS anything (see the remarks above).</summary>
    public bool CanLogOnWithoutBnls() =>
        !BncsProduct.UsesNewLoginSystem(Config.Product) &&
        !BncsProduct.IsStimpakBacked(Config.Product) &&
        (!BncsProduct.RequiresCdKey(Config.Product) ||
         (CdKeyDecoder.Decode(Config.CdKey) is not null &&
          (!BncsProduct.RequiresExpansionCdKey(Config.Product) || CdKeyDecoder.Decode(Config.ExpansionCdKey) is not null))) &&
        LogonCheckCache.TryGetVersionByte(Config.Product, out _);

    private async Task RunRapidReconnectAsync(CancellationToken cancellationToken)
    {
        if (Interlocked.Exchange(ref _rapidReconnectRunning, 1) == 1)
        {
            return;
        }

        var clock = Stopwatch.StartNew();
        var attempts = 0;
        TimeSpan? attemptStartedAt = null;
        LogWarning("Connection lost — rapid reconnect: trying every second until it's back.");
        try
        {
            while (true)
            {
                if (_auth.LoggedOnToBncs || _logonRejection is not null)
                {
                    break;
                }

                var attemptInFlight = attemptStartedAt is { } started && (_bncs.IsConnected || _bnls.IsConnected) && clock.Elapsed - started < RapidAttemptTimeout;
                if (!attemptInFlight)
                {
                    attempts++;
                    attemptStartedAt = clock.Elapsed;
                    try
                    {
                        await ConnectForRapidReconnectAsync(cancellationToken).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (ex is SocketException or IOException or TimeoutException)
                    {
                        LogDebug($"Rapid reconnect attempt {attempts}: {ex.Message}");
                    }
                }

                await Task.Delay(RapidRetryInterval, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Logging on cancels the countdown (OnLoggedOnAsync), and so does Disconnect.
        }
        finally
        {
            Interlocked.Exchange(ref _rapidReconnectRunning, 0);
        }

        if (_auth.LoggedOnToBncs)
        {
            LogInfo($"Rapid reconnect: back on after {attempts} attempt{(attempts == 1 ? "" : "s")} in {clock.Elapsed.TotalSeconds:0.0}s.");
        }
    }

    /// <summary>One attempt: straight to the Battle.net server with cached answers when possible, the normal way otherwise.</summary>
    private async Task ConnectForRapidReconnectAsync(CancellationToken cancellationToken)
    {
        AbandonConnectionAttempt();
        if (!CanLogOnWithoutBnls() || !LogonCheckCache.TryGetVersionByte(Config.Product, out var versionByte))
        {
            await ConnectAsync(cancellationToken).ConfigureAwait(false);
            return;
        }

        _isIntentionalDisconnect = false;
        _logonRejection = null;
        _auth.UsingCachedChecks = true;
        _auth.VersionByte = TryParseVersionByteOverride(Config.VersionByteOverride, out var overrideByte) ? overrideByte : versionByte;
        StartDiscordBridgeIfEnabled();
        LogDebug($"Rapid reconnect: connecting straight to {Config.BattlenetServer} (no BNLS).");
        await _bncs.ConnectAsync(Config.BattlenetServer, Config.BattlenetPort, cancellationToken, BuildProxyOptions()).ConfigureAwait(false);
    }

    /// <summary>Closes whatever a previous attempt left half-open, without it counting as a drop.</summary>
    private void AbandonConnectionAttempt()
    {
        if (_bncs.IsConnected)
        {
            Interlocked.Exchange(ref _replacingBncsConnection, 1);
        }

        _bncs.Close();
        _bnls.Close();
    }
}
