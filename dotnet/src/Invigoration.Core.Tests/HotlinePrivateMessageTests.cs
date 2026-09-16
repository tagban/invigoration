using System.Reflection;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// Transaction 104 carries two different things: a broadcast from the server, and a private message
/// relayed from another user. Telling them apart by whether a sender is present is what makes a PM
/// routable — before this, every private message arrived as an anonymous line in the chat log with
/// no way to see who sent it or reply.
/// </summary>
public class HotlinePrivateMessageTests
{
    private static (List<HotlinePrivateMessage> Pms, List<string> ServerMessages) Dispatch(params HotlineField[] fields)
    {
        var client = new HotlineTransactionClient();
        var pms = new List<HotlinePrivateMessage>();
        var serverMessages = new List<string>();
        client.PrivateMessageReceived += pm => pms.Add(pm);
        client.ServerMessageReceived += m => serverMessages.Add(m);

        // The receive path treats the first frame as the handshake reply, so mark that done or
        // nothing past it is ever dispatched.
        typeof(HotlineTransactionClient)
            .GetField("_handshakeComplete", BindingFlags.NonPublic | BindingFlags.Instance)!
            .SetValue(client, true);

        // Fed in as encoded bytes through the same path the socket uses, so the dispatch itself is
        // what's under test rather than a hand-called handler.
        var frame = HotlineTransactionFrame.Create(HotlineTransactionType.ServerMessage, 1, fields).Encode();
        typeof(HotlineTransactionClient)
            .GetMethod("OnPacketReceivedCore", BindingFlags.NonPublic | BindingFlags.Instance)!
            .Invoke(client, [frame]);

        return (pms, serverMessages);
    }

    [Fact]
    public void AMessageCarryingASender_IsAPrivateMessage()
    {
        var (pms, serverMessages) = Dispatch(
            new HotlineField(HotlineFieldType.UserId, (ushort)42),
            new HotlineField(HotlineFieldType.UserName, "Knezzen"),
            new HotlineField(HotlineFieldType.Data, "are you there?"));

        var pm = Assert.Single(pms);
        Assert.Equal(42, pm.SenderId);
        Assert.Equal("Knezzen", pm.SenderName);
        Assert.Equal("are you there?", pm.Text);
        Assert.Empty(serverMessages);
    }

    /// <summary>A broadcast has no sender — it must stay a server message, not become a PM from nobody.</summary>
    [Fact]
    public void AMessageWithNoSender_StaysAServerMessage()
    {
        var (pms, serverMessages) = Dispatch(new HotlineField(HotlineFieldType.Data, "Server going down in 5 minutes."));

        Assert.Empty(pms);
        Assert.Equal("Server going down in 5 minutes.", Assert.Single(serverMessages));
    }

    /// <summary>Some servers send the name without an id; it's still a private message, just one that can't be replied to by id alone.</summary>
    [Fact]
    public void AMessageWithANameButNoId_IsStillAPrivateMessage()
    {
        var (pms, _) = Dispatch(
            new HotlineField(HotlineFieldType.UserName, "Ari"),
            new HotlineField(HotlineFieldType.Data, "hi"));

        Assert.Equal("Ari", Assert.Single(pms).SenderName);
    }
}
