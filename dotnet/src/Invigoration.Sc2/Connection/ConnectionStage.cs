namespace Invigoration.Sc2.Connection;

/// <summary>
/// The connect sequence's phases, in order. Mirrors core/src/connection.rs's
/// ConnectionStage: Front web auth -> GameUtilities handoff -> native
/// ("Sunken") resume auth -> chat bootstrap -> connected.
/// </summary>
public enum ConnectionStage
{
    Disconnected,
    WebAuthentication,
    GameUtilities,
    NativeAuthentication,
    ChatBootstrap,
    Connected,
}
