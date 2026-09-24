using System.Reflection;
using Invigoration.Core.Chat;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using Invigoration.Core.Sc2;
using Stimpak;

namespace Invigoration.Core.Tests;

/// <summary>
/// StarCraft II reconnecting, driven with the exact event sequences Stimpak produces (checked
/// against its source): a live session lost to the network or a sign-in elsewhere ends with
/// SessionFailed then StageChanged(Disconnected) — never SessionEnded, which is only its worker
/// dying — and StageChanged(Connected) is the one sign the bot is back in chat. Before, a drop was
/// never noticed at all, and a reconnect that did start never saw itself succeed.
/// </summary>
public class BotEngineSc2ReconnectTests
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;

    /// <summary>A real client that's never connected — events are fed in by hand — standing where a connect would have put it.</summary>
    private static (BotEngine Engine, ISc2ChatClient Client) NewSc2Bot(bool autoReconnect = true)
    {
        StimpakNativeResolver.Register();
        var config = new BotConfig
        {
            Product = BncsProduct.Sc2,
            DisplayName = $"bot-{Guid.NewGuid():N}",
            AutoReconnect = autoReconnect,
            // Long enough that no real attempt (which would go to Battle.net) happens mid-test.
            AutoReconnectDelaySeconds = 600,
        };
        var engine = new BotEngine(config);
        var client = NewClient();
        typeof(BotEngine).GetField("_sc2Client", Private)!.SetValue(engine, client);
        typeof(BotEngine).GetField("_sc2LiveClient", Private)!.SetValue(engine, client);
        return (engine, client);
    }

    private static ISc2ChatClient NewClient() =>
        new StimpakSc2ChatClient(new StimpakClient(new StimpakClientOptions("cc.bnet.invigoration.tests") { CredentialPath = Path.Combine(Path.GetTempPath(), $"stimpak-test-{Guid.NewGuid():N}.bin") }));

    private static Task Feed(BotEngine engine, ISc2ChatClient client, SC2Event next, CancellationToken token = default) =>
        (Task)typeof(BotEngine).GetMethod("HandleSc2EventAsync", Private)!.Invoke(engine, [client, next, token])!;

    private static async Task GetIntoChat(BotEngine engine, ISc2ChatClient client)
    {
        await Feed(engine, client, new Joined(1, new PublicChannel(100, "General"), 1));
        await Feed(engine, client, new StageChanged(Stage.Connected));
    }

    private static async Task<bool> Waited(Func<bool> condition, int timeoutMs = 5000)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }

            await Task.Delay(25);
        }

        return condition();
    }

    [Fact]
    public async Task ALostSession_ClosesItsTabs_SaysSo_AndStartsReconnecting()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        var left = new List<byte>();
        var disconnected = 0;
        engine.Sc2ChannelLeft += left.Add;
        engine.BncsDisconnected += _ => disconnected++;
        await GetIntoChat(engine, client);

        await Feed(engine, client, new SessionFailed("Connection reset by peer (os error 54)"));
        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        Assert.Equal([1], left);
        Assert.Equal(1, disconnected);
        Assert.True(await Waited(() => engine.IsReconnecting), "a dropped SC2 session should start the auto-reconnect");
        Assert.True(engine.IsWaitingToReconnect);

        // The tabs close but the remembered channels don't go with them — they're what the reconnect restores.
        Assert.Contains(engine.Config.Sc2LastChannels, c => c is PublicChannelTarget { Id: 100 });
    }

    [Fact]
    public async Task AnAttemptThatNeverMadeItIntoChat_IsNotADrop()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        var disconnected = 0;
        engine.BncsDisconnected += _ => disconnected++;

        await Feed(engine, client, new StageChanged(Stage.WebAuthentication));
        await Feed(engine, client, new SessionFailed("Connection refused (os error 61)"));
        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        Assert.Equal(0, disconnected);
        Assert.False(engine.IsReconnecting);
        Assert.True(engine.IsIdle, "a failed attempt leaves nothing under way");
    }

    [Fact]
    public async Task BeingBackInChat_EndsTheReconnect()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        await GetIntoChat(engine, client);
        await Feed(engine, client, new StageChanged(Stage.Disconnected));
        Assert.True(await Waited(() => engine.IsReconnecting));

        // The next attempt reuses the same client (as ConnectSc2Async does) and gets back in.
        typeof(BotEngine).GetField("_sc2LiveClient", Private)!.SetValue(engine, client);
        await GetIntoChat(engine, client);

        Assert.True(await Waited(() => !engine.IsReconnecting), "Connected should end the reconnect instead of it carrying on forever");
    }

    [Fact]
    public async Task EventsFromAReplacedClient_ChangeNothing()
    {
        var (engine, _) = NewSc2Bot();
        await using var __ = engine;
        using var stale = NewClient();
        var joined = 0;
        var disconnected = 0;
        engine.Sc2ChannelJoined += (_, _, _) => joined++;
        engine.BncsDisconnected += _ => disconnected++;

        await Feed(engine, stale, new Joined(1, new PublicChannel(100, "General"), 1));
        await Feed(engine, stale, new StageChanged(Stage.Connected));
        await Feed(engine, stale, new StageChanged(Stage.Disconnected));

        Assert.Equal(0, joined);
        Assert.Equal(0, disconnected);
        Assert.False(engine.IsReconnecting);
    }

    /// <summary>Disconnect used to leave the Battle.net login window open with nothing waiting on it.</summary>
    [Fact]
    public async Task Disconnect_ClosesASignInStillOpen()
    {
        var (engine, client) = NewSc2Bot(autoReconnect: false);
        var loopCts = new CancellationTokenSource();
        typeof(BotEngine).GetField("_sc2ReceiveCts", Private)!.SetValue(engine, loopCts);
        var windowClosed = new TaskCompletionSource();
        engine.Sc2ChallengeHandler = async (_, token) =>
        {
            await using var registration = token.Register(() => windowClosed.TrySetResult());
            await Task.Delay(Timeout.Infinite, token);
            return [];
        };
        var errors = new List<string>();
        engine.Log += segments => errors.Add(string.Concat(segments.Select(s => s.Text)));

        var signIn = Feed(engine, client, new AuthenticationRequired(7, "https://example.invalid/login", false), loopCts.Token);
        Assert.False(signIn.IsCompleted);

        await engine.DisconnectAsync();

        await windowClosed.Task.WaitAsync(TimeSpan.FromSeconds(5));
        await signIn.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.DoesNotContain(errors, line => line.Contains("sign-in failed"));
        await engine.DisposeAsync();
    }

    /// <summary>Loses the session, then plays the reconnect getting back in: the same client, in chat again.</summary>
    private static async Task LoseItAndGetBackIn(BotEngine engine, ISc2ChatClient client)
    {
        await Feed(engine, client, new StageChanged(Stage.Disconnected));
        Assert.True(await Waited(() => engine.IsReconnecting));
        typeof(BotEngine).GetField("_sc2LiveClient", Private)!.SetValue(engine, client);
        await GetIntoChat(engine, client);
        Assert.True(await Waited(() => !engine.IsReconnecting));
    }

    /// <summary>
    /// Two sign-ins on one Battle.net account take its chat session from each other. Reconnecting
    /// every time would keep them knocking each other off forever, a Battle.net login each round.
    /// </summary>
    [Fact]
    public async Task TakenOverAgainAndAgain_StopsReconnecting()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        await GetIntoChat(engine, client);
        await LoseItAndGetBackIn(engine, client);
        await LoseItAndGetBackIn(engine, client);

        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        Assert.False(await Waited(() => engine.IsReconnecting, timeoutMs: 500), "a third quick loss should stop reconnecting");
        Assert.NotNull(typeof(BotEngine).GetField("_logonRejection", Private)!.GetValue(engine));
    }

    [Fact]
    public async Task ASessionThatLasted_StartsTheCountOver()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        await GetIntoChat(engine, client);
        await LoseItAndGetBackIn(engine, client);
        await LoseItAndGetBackIn(engine, client);

        // In chat for a good while this time: an ordinary drop, not a takeover.
        typeof(BotEngine).GetField("_connectedAt", Private)!.SetValue(engine, DateTimeOffset.UtcNow.AddHours(-1));
        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        Assert.True(await Waited(() => engine.IsReconnecting));
    }

    /// <summary>Closing the login window is an answer: an auto-reconnect would otherwise pop a new one every attempt.</summary>
    [Fact]
    public async Task ClosingTheSignIn_IsTakenAsNo()
    {
        var (engine, client) = NewSc2Bot();
        await using var _ = engine;
        engine.Sc2ChallengeHandler = (_, _) => throw new InvalidOperationException("The Battle.net login window was closed before finishing.");

        await Feed(engine, client, new AuthenticationRequired(7, "https://example.invalid/login", false));

        Assert.NotNull(typeof(BotEngine).GetField("_logonRejection", Private)!.GetValue(engine));
    }
}
