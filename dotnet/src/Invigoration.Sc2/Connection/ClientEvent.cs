using Invigoration.Sc2.Chat;

namespace Invigoration.Sc2.Connection;

/// <summary>
/// Events raised by a running client actor. Mirrors core/src/connection.rs's
/// ClientEvent. <see cref="Authentication"/> is the real-Battle.net-login
/// step: the actor hands the caller a URL to show in a browser (Avalonia's
/// WebAuthenticationBroker) and awaits the resulting credential via
/// <see cref="AuthenticationRequest.Completion"/> — the actor never sees a
/// password directly.
/// </summary>
public abstract record ClientEvent
{
    private ClientEvent()
    {
    }

    public sealed record Stage(ConnectionStage Value) : ClientEvent;

    public sealed record Authentication(AuthenticationRequest Request) : ClientEvent;

    public sealed record Chat(ChatEvent Value) : ClientEvent;

    public sealed record CommandError(string Message) : ClientEvent;

    public sealed record Error(string Message) : ClientEvent;
}

/// <summary>A pending web-authentication step: show <see cref="Url"/> to the user, then complete <see cref="Completion"/> with the resulting credential (or fault it on cancellation/failure).</summary>
public sealed record AuthenticationRequest(Uri Url, TaskCompletionSource<SecretBytes> Completion);
