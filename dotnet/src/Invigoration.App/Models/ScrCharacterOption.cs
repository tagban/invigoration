using Invigoration.Scr;

namespace Invigoration.App.Models;

/// <summary>One StarCraft: Remastered character in the bot settings picker. An empty name means the first character on the gateway.</summary>
public sealed record ScrCharacterOption(uint Gateway, string Name, string? GatewayName = null)
{
    public string Label => $"{(Name.Length > 0 ? Name : "First character")} ({GatewayName ?? ScrGateways.NameOf(Gateway)})";

    public override string ToString() => Label;
}
