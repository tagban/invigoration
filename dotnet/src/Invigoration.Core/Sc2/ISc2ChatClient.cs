using Stimpak;

namespace Invigoration.Core.Sc2;

/// <summary>
/// What BotEngine needs from a StarCraft II chat connection. Stimpak's client is one
/// (<see cref="StimpakSc2ChatClient"/>); Invigoration's own native client is the other
/// (<see cref="NativeSc2ChatClient"/>), which speaks in Stimpak's event types so the
/// same handler, roster and channel tabs serve both. Failures are thrown as
/// <see cref="StimpakException"/> by either, so BotEngine's error handling is shared too.
/// </summary>
public interface ISc2ChatClient : IDisposable
{
    PeopleRegistry People { get; }

    /// <summary>Raised on the client's own thread as each event is produced, before it's readable from <see cref="ReadEventsAsync"/>.</summary>
    event Action<SC2Event>? EventReceived;

    IAsyncEnumerable<SC2Event> ReadEventsAsync(CancellationToken cancellation = default);

    void Connect(StimpakConnectOptions options);

    void Disconnect();

    void JoinPublic(ushort channelId);

    void JoinPrivate(string name);

    void Leave(byte channelIndex);

    void SendMessage(byte channelIndex, string body);

    void SendWhisper(string name, string body);

    void SubmitAuth(ulong authId, string token);

    void CancelAuth(ulong authId);
}

/// <summary>Stimpak's native-library client, unchanged, behind <see cref="ISc2ChatClient"/>.</summary>
public sealed class StimpakSc2ChatClient(StimpakClient client) : ISc2ChatClient
{
    public StimpakClient Inner { get; } = client;

    public PeopleRegistry People => Inner.People;

    public event Action<SC2Event>? EventReceived
    {
        add => Inner.EventReceived += value;
        remove => Inner.EventReceived -= value;
    }

    public IAsyncEnumerable<SC2Event> ReadEventsAsync(CancellationToken cancellation = default) => Inner.ReadEventsAsync(cancellation);

    public void Connect(StimpakConnectOptions options) => Inner.Connect(options);

    public void Disconnect() => Inner.Disconnect();

    public void JoinPublic(ushort channelId) => Inner.JoinPublic(channelId);

    public void JoinPrivate(string name) => Inner.JoinPrivate(name);

    public void Leave(byte channelIndex) => Inner.Leave(channelIndex);

    public void SendMessage(byte channelIndex, string body) => Inner.SendMessage(channelIndex, body);

    public void SendWhisper(string name, string body) => Inner.SendWhisper(name, body);

    public void SubmitAuth(ulong authId, string token) => Inner.SubmitAuth(authId, token);

    public void CancelAuth(ulong authId) => Inner.CancelAuth(authId);

    public void Dispose() => Inner.Dispose();
}
