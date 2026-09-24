using System.Reflection;
using Invigoration.Core.Config;
using Invigoration.Core.Protocol;
using Invigoration.Core.Sc2;
using Stimpak;

namespace Invigoration.Core.Tests;

/// <summary>
/// One user at a time for a Battle.net login's saved sign-in. Two logins from one saved sign-in at
/// once end with the older one deleting the credential the newer one just saved, which is how an
/// SC2 bot lost its login.
/// </summary>
[Collection("BattlenetCredentialProfileStore")]
public class BattlenetSignInLeaseTests
{
    private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;

    private static string NewProfileId() => BattlenetCredentialProfileStore.CreateAndSave($"login-{Guid.NewGuid():N}").Id;

    private static (BotEngine Engine, List<string> Log) NewSc2Bot(string profileId, string name)
    {
        StimpakNativeResolver.Register();
        var engine = new BotEngine(new BotConfig
        {
            Product = BncsProduct.Sc2,
            DisplayName = name,
            BattlenetCredentialProfileId = profileId,
            AutoReconnect = false,
        });
        var log = new List<string>();
        engine.Log += segments => log.Add(string.Concat(segments.Select(s => s.Text)));
        return (engine, log);
    }

    private static ISc2ChatClient NewClient() =>
        new StimpakSc2ChatClient(new StimpakClient(new StimpakClientOptions("cc.bnet.invigoration.tests") { CredentialPath = Path.Combine(Path.GetTempPath(), $"stimpak-test-{Guid.NewGuid():N}.bin") }));

    private static Task Feed(BotEngine engine, ISc2ChatClient client, SC2Event next) =>
        (Task)typeof(BotEngine).GetMethod("HandleSc2EventAsync", Private)!.Invoke(engine, [client, next, CancellationToken.None])!;

    /// <summary>An SC2 bot part-way through an attempt on <paramref name="profileId"/>, holding its sign-in as ConnectSc2Async leaves it.</summary>
    private static (ISc2ChatClient Client, BattlenetSignInLease Lease) Attempting(BotEngine engine, string profileId)
    {
        Assert.True(BattlenetSignInLease.TryAcquire(profileId, engine, engine.Config.DisplayName, out var lease, out _));
        var client = NewClient();
        typeof(BotEngine).GetField("_sc2Client", Private)!.SetValue(engine, client);
        typeof(BotEngine).GetField("_sc2ClientProfileId", Private)!.SetValue(engine, profileId);
        typeof(BotEngine).GetField("_sc2LiveClient", Private)!.SetValue(engine, client);
        typeof(BotEngine).GetField("_sc2Lease", Private)!.SetValue(engine, lease);
        return (client, lease);
    }

    [Fact]
    public void OneHolderAtATime_AndFreeAgainOnceLetGo()
    {
        var profileId = NewProfileId();
        var first = new object();

        Assert.True(BattlenetSignInLease.TryAcquire(profileId, first, "bot A", out var lease, out _));
        Assert.False(BattlenetSignInLease.TryAcquire(profileId, new object(), "bot B", out _, out var holder));
        Assert.Same(first, holder!.Owner);
        Assert.Equal("bot A", holder.OwnerName);

        lease.Dispose();

        Assert.True(lease.Released.IsCompleted);
        Assert.True(BattlenetSignInLease.TryAcquire(profileId, new object(), "bot B", out var next, out _));
        next.Dispose();
    }

    [Fact]
    public void DifferentLogins_DontGetInEachOthersWay()
    {
        Assert.True(BattlenetSignInLease.TryAcquire(NewProfileId(), new object(), "bot A", out var a, out _));
        Assert.True(BattlenetSignInLease.TryAcquire(NewProfileId(), new object(), "bot B", out var b, out _));
        a.Dispose();
        b.Dispose();
    }

    [Fact]
    public async Task ConnectingOnALoginAnotherBotHas_IsRefused_AndStopsAnyReconnect()
    {
        var profileId = NewProfileId();
        Assert.True(BattlenetSignInLease.TryAcquire(profileId, new object(), "bot A", out var held, out _));
        var (engine, log) = NewSc2Bot(profileId, "bot B");
        await using var disposeEngine = engine;

        try
        {
            await engine.ConnectAsync();

            Assert.Null(typeof(BotEngine).GetField("_sc2Client", Private)!.GetValue(engine));
            Assert.NotNull(typeof(BotEngine).GetField("_logonRejection", Private)!.GetValue(engine));
            Assert.Contains(log, line => line.Contains("in use by bot A"));
            Assert.True(engine.IsIdle);
        }
        finally
        {
            held.Dispose();
        }
    }

    /// <summary>Stimpak's Disconnected: that attempt is over, so another bot can have the login.</summary>
    [Fact]
    public async Task AnAttemptEnding_LetsGoOfTheSignIn()
    {
        var profileId = NewProfileId();
        var (engine, _) = NewSc2Bot(profileId, "bot A");
        await using var __ = engine;
        var (client, lease) = Attempting(engine, profileId);

        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        Assert.True(lease.Released.IsCompleted);
    }

    /// <summary>
    /// Stimpak only takes a Disconnect once it's in chat, so a client stopped mid-attempt carries on
    /// using the saved sign-in. Its hold has to last until it says it has stopped, or the next connect
    /// on that login races it.
    /// </summary>
    [Fact]
    public async Task ADisconnectMidAttempt_KeepsTheSignIn_UntilStimpakSaysItHasStopped()
    {
        var profileId = NewProfileId();
        var (engine, _) = NewSc2Bot(profileId, "bot A");
        var (client, lease) = Attempting(engine, profileId);

        await engine.DisconnectAsync();

        Assert.False(lease.Released.IsCompleted, "a client still mid-attempt shouldn't let go yet");
        Assert.True(engine.IsIdle, "the bot itself has stopped, even while its old client finishes");

        await Feed(engine, client, new StageChanged(Stage.Disconnected));

        await lease.Released.WaitAsync(TimeSpan.FromSeconds(5));
        await engine.DisposeAsync();
    }

    [Theory]
    [InlineData(30, "30 seconds")]
    [InlineData(5 * 60, "5 minutes")]
    [InlineData(3 * 3600, "3 hours")]
    [InlineData(9 * 86400, "9 days")]
    public void DescribeAge_ReadsNaturally(int seconds, string expected)
    {
        Assert.Equal(expected, BotEngine.DescribeAge(TimeSpan.FromSeconds(seconds)));
    }
}
