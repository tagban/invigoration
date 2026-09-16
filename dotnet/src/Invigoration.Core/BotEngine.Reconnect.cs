using System.Diagnostics;
using System.Net.Sockets;
using Invigoration.Core.Auth;
using Invigoration.Core.Protocol;

namespace Invigoration.Core;

/// <summary>
/// Reconnecting after an unexpected drop (BotConfig.AutoReconnect): an attempt every
/// AutoReconnectDelaySeconds until the bot is logged back on, giving up after
/// AutoReconnectMaxAttempts (0 keeps trying). The attempts count across drops — an attempt the
/// server hangs up on is just a failed attempt, not a new disconnect starting the count over.
/// </summary>
/// <remarks>
/// <para>Rapid reconnect (BotConfig.RapidReconnect) is for private servers where getting back on
/// first matters — holding a name, or ops in a channel. The first attempt goes out the instant the
/// connection drops instead of after the wait, and each attempt skips BNLS entirely when it can:
/// the CD key and password are hashed locally, and the two answers only BNLS computes (the version
/// byte, and the version check for the server's challenge) are reused from an earlier logon this
/// run (LogonCheckCache), so the whole logon is just the Battle.net server itself — measured at
/// roughly half the time of a logon that asks BNLS. When it can't (Warcraft III's logon, a CD key
/// only BNLS decodes, nothing cached yet, or a server that changed its challenge), the attempt logs
/// on the normal way. Never used on official Battle.net, where fast repeated connections get an
/// address banned; the ordinary reconnect applies there instead.</para>
/// </remarks>
public sealed partial class BotEngine
{
    /// <summary>An attempt that connected but still isn't logged on after this long is given up for a fresh one.</summary>
    private static readonly TimeSpan AttemptTimeout = TimeSpan.FromSeconds(10);

    private int _reconnectRunning;

    public static bool IsOfficialBattlenetServer(string server) =>
        BncsProduct.OfficialBattlenetServers.Contains(server.Trim(), StringComparer.OrdinalIgnoreCase);

    /// <summary>Whether a drop right now would be handled by rapid reconnect: on, a classic game, and not official Battle.net.</summary>
    public bool RapidReconnectApplies =>
        Config.RapidReconnect && !BncsProduct.IsStimpakBacked(Config.Product) && !IsOfficialBattlenetServer(Config.BattlenetServer);

    /// <summary>Whether this bot can log on without asking BNLS anything (see the remarks above).</summary>
    public bool CanLogOnWithoutBnls() =>
        // The Chat protocol never involves BNLS at all — there's nothing to cache or skip, so
        // every attempt is already the fast path (see BotEngine.Chat.cs).
        BncsProduct.IsChatTelnet(Config.Product) ||
        (!BncsProduct.UsesNewLoginSystem(Config.Product) &&
         !BncsProduct.IsStimpakBacked(Config.Product) &&
         (!BncsProduct.RequiresCdKey(Config.Product) ||
          (CdKeyDecoder.Decode(Config.CdKey) is not null &&
           (!BncsProduct.RequiresExpansionCdKey(Config.Product) || CdKeyDecoder.Decode(Config.ExpansionCdKey) is not null))) &&
         LogonCheckCache.TryGetVersionByte(Config.Product, out _));

    private async Task RunReconnectAsync(CancellationToken cancellationToken)
    {
        if (Interlocked.Exchange(ref _reconnectRunning, 1) == 1)
        {
            return;
        }

        var rapid = RapidReconnectApplies;
        var interval = TimeSpan.FromSeconds(Math.Max(1, Config.AutoReconnectDelaySeconds));
        var maxAttempts = Math.Max(0, Config.AutoReconnectMaxAttempts);
        var limit = maxAttempts > 0 ? $", up to {maxAttempts} attempt{(maxAttempts == 1 ? "" : "s")}" : "";
        var clock = Stopwatch.StartNew();
        var attempts = 0;
        TimeSpan? attemptStartedAt = null;
        var gaveUp = false;

        if (rapid)
        {
            LogWarning($"Connection lost — rapid reconnect: trying now, then every {interval.TotalSeconds:0}s{limit}.");
        }
        else
        {
            LogInfo($"Unexpected disconnect — reconnecting in {interval.TotalSeconds:0}s{limit} (auto-reconnect is on).");
        }

        try
        {
            if (!rapid)
            {
                await Task.Delay(interval, cancellationToken).ConfigureAwait(false);
            }

            while (!_auth.LoggedOnToBncs && _logonRejection is null)
            {
                var attemptInFlight = attemptStartedAt is { } started &&
                    (_bncs.IsConnected || _bnls.IsConnected || _chatTelnet.IsConnected) &&
                    clock.Elapsed - started < AttemptTimeout;
                if (!attemptInFlight)
                {
                    if (maxAttempts > 0 && attempts >= maxAttempts)
                    {
                        gaveUp = true;
                        break;
                    }

                    attempts++;
                    attemptStartedAt = clock.Elapsed;
                    if (attempts > 1)
                    {
                        LogDebug($"Reconnect attempt {attempts}{(maxAttempts > 0 ? $" of {maxAttempts}" : "")}.");
                    }

                    try
                    {
                        await (rapid ? ConnectForRapidReconnectAsync(cancellationToken) : ConnectCoreAsync(cancellationToken)).ConfigureAwait(false);
                    }
                    catch (Exception ex) when (ex is SocketException or IOException or TimeoutException)
                    {
                        LogDebug($"Reconnect attempt {attempts}: {ex.Message}");
                    }
                }

                await Task.Delay(interval, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
            // Logging on cancels the wait (OnLoggedOnAsync), and so do Connect and Disconnect.
        }
        finally
        {
            Interlocked.Exchange(ref _reconnectRunning, 0);
        }

        if (_auth.LoggedOnToBncs && attempts > 0)
        {
            LogInfo($"Reconnected after {attempts} attempt{(attempts == 1 ? "" : "s")} in {clock.Elapsed.TotalSeconds:0.0}s.");
        }
        else if (gaveUp)
        {
            AbandonConnectionAttempt();
            LogError($"Gave up reconnecting after {attempts} attempt{(attempts == 1 ? "" : "s")}. Connect again when the server's back.");
        }
    }

    /// <summary>One attempt: straight to the Battle.net server with cached answers when possible, the normal way otherwise.</summary>
    private async Task ConnectForRapidReconnectAsync(CancellationToken cancellationToken)
    {
        AbandonConnectionAttempt();

        // Nothing to skip on the Chat protocol — its ordinary connect is already a socket and two
        // credential lines, so the normal path *is* the fast one.
        if (BncsProduct.IsChatTelnet(Config.Product))
        {
            await ConnectCoreAsync(cancellationToken).ConfigureAwait(false);
            return;
        }

        if (!CanLogOnWithoutBnls() || !LogonCheckCache.TryGetVersionByte(Config.Product, out var versionByte))
        {
            await ConnectCoreAsync(cancellationToken).ConfigureAwait(false);
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
        if (_bncs.IsConnected || _chatTelnet.IsConnected)
        {
            Interlocked.Exchange(ref _replacingBncsConnection, 1);
        }

        _bncs.Close();
        _bnls.Close();
        _chatTelnet.Close();
    }
}
