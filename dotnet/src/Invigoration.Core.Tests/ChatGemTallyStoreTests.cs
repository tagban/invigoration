using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Serialized against the rest of the suite: ChatGemTallyStore is a static store with a
/// process-wide cache and an on-disk file, so parallel tests would race each other's state — the
/// same hazard that has repeatedly produced flaky failures around ClanRosterStore.
/// </summary>
[Collection("ChatGemTallyStore")]
public class ChatGemTallyStoreTests : IDisposable
{
    private readonly string _tempDir;

    public ChatGemTallyStoreTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "invig-gem-tally-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        ChatGemTallyStore.ConfigDirectoryOverride = _tempDir;
    }

    public void Dispose()
    {
        ChatGemTallyStore.ConfigDirectoryOverride = null;
        ChatGemTallyStore.ResetCacheForTests();
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup of a temp dir; not worth failing a test over.
        }
    }

    private static readonly DateTimeOffset September = new(2026, 9, 12, 1, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset October = new(2026, 10, 1, 0, 5, 0, TimeSpan.Zero);

    [Fact]
    public void ShareConsent_DefaultsToUnanswered_AndSharingIsOff()
    {
        Assert.Null(ChatGemTallyStore.ShareConsent);
        Assert.False(ChatGemTallyStore.SharingEnabled);
    }

    [Fact]
    public void SharingEnabled_OnlyAfterAnExplicitYes()
    {
        ChatGemTallyStore.ShareConsent = false;
        Assert.False(ChatGemTallyStore.SharingEnabled);

        ChatGemTallyStore.ShareConsent = true;
        Assert.True(ChatGemTallyStore.SharingEnabled);
    }

    [Fact]
    public void RecordActivation_CountsPerAccount()
    {
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.RecordActivation("someoneelse@bnet.cc", September);

        Assert.Equal(2, ChatGemTallyStore.CurrentMonthCount("tagban@bnet.cc", September));
        Assert.Equal(1, ChatGemTallyStore.CurrentMonthCount("someoneelse@bnet.cc", September));
    }

    // Counting is deliberately not gated on consent — only submission is — so that answering
    // "yes" partway through a month doesn't restart the user at zero.
    [Fact]
    public void RecordActivation_CountsEvenWithoutConsent()
    {
        Assert.Null(ChatGemTallyStore.ShareConsent);

        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);

        Assert.Equal(1, ChatGemTallyStore.CurrentMonthCount("tagban@bnet.cc", September));
    }

    [Fact]
    public void RecordActivation_RollsOverIntoANewMonth()
    {
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);

        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", October);

        Assert.Equal(1, ChatGemTallyStore.CurrentMonthCount("tagban@bnet.cc", October));
        // September's figure is gone from the live tally rather than added to — the leaderboard
        // resets monthly.
        Assert.Equal(0, ChatGemTallyStore.CurrentMonthCount("tagban@bnet.cc", September));
    }

    [Fact]
    public void PendingSubmissions_EmptyUntilThereAreUnacknowledgedClicks()
    {
        Assert.Empty(ChatGemTallyStore.PendingSubmissions());

        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);

        var pending = ChatGemTallyStore.PendingSubmissions();
        Assert.Single(pending);
        Assert.Equal("tagban@bnet.cc", pending[0].AccountKey);
        Assert.Equal(1, pending[0].Tally.Count);
    }

    [Fact]
    public void MarkSubmitted_ClearsPendingUntilTheCountMovesAgain()
    {
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.MarkSubmitted("tagban@bnet.cc", 1);

        Assert.Empty(ChatGemTallyStore.PendingSubmissions());

        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        Assert.Single(ChatGemTallyStore.PendingSubmissions());
    }

    // Regression: submissions carry a cumulative figure, so an out-of-order or replayed
    // acknowledgement must never drag Submitted backwards — that would resend forever.
    [Fact]
    public void MarkSubmitted_NeverMovesBackwards()
    {
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.MarkSubmitted("tagban@bnet.cc", 3);

        ChatGemTallyStore.MarkSubmitted("tagban@bnet.cc", 1);

        Assert.Empty(ChatGemTallyStore.PendingSubmissions());
    }

    [Fact]
    public void Tallies_SurviveAReload()
    {
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);
        ChatGemTallyStore.ShareConsent = true;

        ChatGemTallyStore.ResetCacheForTests();

        Assert.True(ChatGemTallyStore.SharingEnabled);
        Assert.Equal(1, ChatGemTallyStore.CurrentMonthCount("tagban@bnet.cc", September));
    }

    [Fact]
    public void RecordActivation_IgnoresABlankAccountKey()
    {
        ChatGemTallyStore.RecordActivation("", September);
        ChatGemTallyStore.RecordActivation("   ", September);

        Assert.Empty(ChatGemTallyStore.PendingSubmissions());
    }

    // The sender must be inert in a build with no endpoint configured, no matter what the user
    // consented to — this is the guard that keeps an unconfigured build from transmitting.
    [Fact]
    public void Sender_IsNotConfigured_AndCannotSubmitEvenWithConsent()
    {
        ChatGemTallyStore.ShareConsent = true;

        Assert.False(ChatGemTallySender.IsConfigured);
        Assert.False(ChatGemTallySender.CanSubmit);
    }

    [Fact]
    public async Task Sender_FlushIsANoOpWhenUnconfigured()
    {
        ChatGemTallyStore.ShareConsent = true;
        ChatGemTallyStore.RecordActivation("tagban@bnet.cc", September);

        await new ChatGemTallySender().FlushAsync();

        // Still pending: nothing was sent, so nothing was acknowledged.
        Assert.Single(ChatGemTallyStore.PendingSubmissions());
    }
}
