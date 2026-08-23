using System.Reflection;
using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Covers two real bugs found via live testing against atlas.bnetdocs.org:
/// (1) the login handshake originally matched literal "Username:"/
/// "Password:" text, but that server's actual banner reads "Enter your
/// login name and password." (not "account name" like the sample this was
/// first built from) — fixed to recognize a prompt by shape (line ends with
/// ':') instead of specific wording; (2) atlas.bnetdocs.org turned out not
/// to send field-specific prompts *at all* — just that one instructional
/// sentence, then silence, so the bot still hung even after fix (1). Fixed
/// with a delayed bare-telnet fallback (see SendCredentialsIfNoPromptArrivesAsync)
/// that blind-sends username then password if no real prompt shows up
/// shortly after a sentence that mentions both a name-ish word and
/// "password" — but backs off (no-ops) if a real prompt arrives first,
/// since a server that *does* send explicit prompts says basically the same
/// introductory sentence too.
/// </summary>
public class BotEngineChatTelnetLoginTests
{
    private static Task InvokeHandleLine(BotEngine engine, string line)
    {
        var method = typeof(BotEngine).GetMethod("HandleChatTelnetLineAsync", BindingFlags.NonPublic | BindingFlags.Instance)!;
        return (Task)method.Invoke(engine, [line])!;
    }

    private static int GetPromptsSeen(BotEngine engine) =>
        (int)typeof(BotEngine).GetField("_chatTelnetPromptsSeen", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;

    private static bool GetCredentialsSent(BotEngine engine) =>
        (bool)typeof(BotEngine).GetField("_chatTelnetCredentialsSent", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;

    private static bool GetLoggedIn(BotEngine engine) =>
        (bool)typeof(BotEngine).GetField("_chatTelnetLoggedIn", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;

    [Fact]
    public async Task HandleChatTelnetLineAsync_LiveAtlasBnetdocsWording_FallsBackToBareTelnetLoginAfterDelay()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleLine(engine, "Connection from [73.175.18.108:53279]");
        await InvokeHandleLine(engine, "");
        await InvokeHandleLine(engine, "Enter your login name and password.");
        await InvokeHandleLine(engine, "");

        // No real prompt ever arrives on this server — nothing should be sent yet.
        Assert.False(GetCredentialsSent(engine));

        // The fallback waits ~500ms for a real prompt before blind-sending; give it enough
        // margin to fire without making this test flaky under CI/machine load.
        await Task.Delay(700);

        Assert.True(GetCredentialsSent(engine));
        Assert.False(GetLoggedIn(engine));

        await InvokeHandleLine(engine, "2010 NAME SomeUser");
        Assert.True(GetLoggedIn(engine));
    }

    [Fact]
    public async Task HandleChatTelnetLineAsync_OriginalSampleWording_StillRecognizesBothPrompts()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleLine(engine, "Enter your account name and password.");
        Assert.Equal(0, GetPromptsSeen(engine));
        Assert.False(GetCredentialsSent(engine));

        await InvokeHandleLine(engine, "Username: ");
        Assert.Equal(1, GetPromptsSeen(engine));

        await InvokeHandleLine(engine, "Password: ");
        Assert.Equal(2, GetPromptsSeen(engine));
        Assert.True(GetCredentialsSent(engine));
    }

    [Fact]
    public async Task HandleChatTelnetLineAsync_RealPromptArrivesBeforeFallbackFires_FallbackBacksOff()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        // Same introductory sentence a bare-telnet server would send, but this one follows up
        // with a real prompt shortly after — the delayed fallback must not also blind-send.
        await InvokeHandleLine(engine, "Enter your account name and password.");
        await InvokeHandleLine(engine, "Username: ");
        await InvokeHandleLine(engine, "Password: ");
        Assert.Equal(2, GetPromptsSeen(engine));
        Assert.True(GetCredentialsSent(engine));

        await Task.Delay(700);

        // Still exactly 2 prompts handled — the fallback's delayed check found
        // _chatTelnetCredentialsSent already true and no-op'd.
        Assert.Equal(2, GetPromptsSeen(engine));
    }

    [Fact]
    public async Task HandleChatTelnetLineAsync_FirstEventLineWithoutNameConfirmation_StillMarksLoggedIn()
    {
        var config = new BotConfig();
        await using var engine = new BotEngine(config);

        await InvokeHandleLine(engine, "1007 CHANNEL \"Public Chat 1\"");

        Assert.True(GetLoggedIn(engine));
    }
}
