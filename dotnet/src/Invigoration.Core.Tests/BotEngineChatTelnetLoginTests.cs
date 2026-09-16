using System.Net;
using System.Net.Sockets;
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

    /// <summary>
    /// The Chat path marks the session on through the same BotSessionAuth flag every other
    /// login path uses, rather than a flag of its own — that's what lets chat be sent, stops a
    /// reconnect that's still counting down, and ends the reconnect loop (see BotEngine.Chat.cs's
    /// OnChatTelnetLoggedOnAsync).
    /// </summary>
    private static bool GetLoggedIn(BotEngine engine)
    {
        var auth = typeof(BotEngine).GetField("_auth", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(engine)!;
        return (bool)auth.GetType().GetProperty("LoggedOnToBncs")!.GetValue(auth)!;
    }

    /// <summary>
    /// Against a real socket, since the fallback deliberately declines to send on a connection
    /// that's already gone — driving the line handler with no socket at all would exercise that
    /// bail-out instead of the login it's meant to cover.
    /// </summary>
    [Fact]
    public async Task ABannerWithNoPrompts_FallsBackToBareTelnetLoginAfterDelay()
    {
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        var clientLines = new List<string>();
        // Held open until the assertions are done: closing the socket is a disconnect like any
        // other, and the engine (correctly) clears the logged-on flag when one happens.
        var finished = new TaskCompletionSource();
        var loggedOn = new TaskCompletionSource();
        var served = Task.Run(async () =>
        {
            using var client = await listener.AcceptTcpClientAsync();
            using var stream = client.GetStream();

            var handshake = new byte[2];
            await stream.ReadExactlyAsync(handshake);
            Assert.Equal([0x03, 0x04], handshake);

            // atlas.bnetdocs.org's actual banner: an instructional sentence and nothing else — no
            // "Username:"/"Password:" prompt ever follows.
            await stream.WriteAsync("Connection from [73.175.18.108:53279]\r\n\r\nEnter your login name and password.\r\n\r\n"u8.ToArray());

            var buffer = new byte[256];
            var deadline = DateTime.UtcNow.AddSeconds(5);
            var text = "";
            while (DateTime.UtcNow < deadline && text.Count(c => c == '\n') < 2)
            {
                var read = await stream.ReadAsync(buffer);
                if (read == 0)
                {
                    break;
                }

                text += System.Text.Encoding.UTF8.GetString(buffer, 0, read);
            }

            clientLines.AddRange(text.Split("\r\n", StringSplitOptions.RemoveEmptyEntries));
            await stream.WriteAsync("2010 NAME SomeUser\r\n"u8.ToArray());
            loggedOn.TrySetResult();
            await finished.Task.WaitAsync(TimeSpan.FromSeconds(10));
        });

        await using var engine = new BotEngine(new BotConfig
        {
            Product = Invigoration.Core.Protocol.BncsProduct.Chat,
            Username = "SomeUser",
            Password = "somepassword",
            BattlenetServer = "127.0.0.1",
            BattlenetPort = port,
        });
        await engine.ConnectAsync();
        await loggedOn.Task.WaitAsync(TimeSpan.FromSeconds(10));

        try
        {
            Assert.Equal(["SomeUser", "somepassword"], clientLines);
            Assert.True(GetCredentialsSent(engine));

            // The confirmation line is read and handled on the engine's own receive loop, so wait
            // for it rather than assuming it landed the instant the server finished writing.
            Assert.True(await Waited(() => GetLoggedIn(engine)), "the bot never registered the logon");
        }
        finally
        {
            finished.TrySetResult();
            await served;
        }
    }

    /// <summary>Polls until the condition holds or the timeout passes — for state a background receive loop sets.</summary>
    private static async Task<bool> Waited(Func<bool> condition, int timeoutMs = 3000)
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
