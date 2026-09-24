using Invigoration.Sc2.Protobuf;

namespace Invigoration.Scr;

/// <summary>A StarCraft: Remastered character. Characters belong to one gateway; the same name can exist on several.</summary>
public sealed record ScrToon(uint Id, string Name, uint Gateway)
{
    public override string ToString() => $"{Name} ({ScrGateways.NameOf(Gateway)})";
}

/// <summary>A gateway as Battle.net announces it (Gateway.GatewayUpdate).</summary>
public sealed record ScrGateway(uint Id, string Name);

/// <summary>What the settings window shows: who signed in, their characters, and the gateways Battle.net offers.</summary>
public sealed record ScrAccount(string? BattleTag, IReadOnlyList<ScrToon> Toons, IReadOnlyList<ScrGateway> Gateways);

/// <summary>
/// SC:R's gateways. The retail client announces these five with GatewayUpdate at sign-in; the names
/// here are only used when that list isn't at hand.
/// </summary>
public static class ScrGateways
{
    public const uint UsWest = 10;
    public const uint UsEast = 11;
    public const uint Europe = 20;
    public const uint Korea = 30;
    public const uint Asia = 45;

    public static IReadOnlyList<ScrGateway> Known { get; } =
    [
        new(UsWest, "U.S. West"),
        new(UsEast, "U.S. East"),
        new(Europe, "Europe"),
        new(Korea, "Korea"),
        new(Asia, "Asia"),
    ];

    public static string NameOf(uint gateway) => Known.FirstOrDefault(g => g.Id == gateway)?.Name ?? $"gateway {gateway}";

    /// <summary>GatewayUpdate's body: {1: {1: id, 2: name, 3: short code, 5: online, 7: news URL}, 2: 0}.</summary>
    public static ScrGateway? DecodeUpdate(byte[] body)
    {
        var outer = new ProtoReader(body);
        while (outer.HasMore)
        {
            var (field, type) = outer.ReadTag();
            if (field != 1 || type != WireType.LengthDelimited)
            {
                outer.Skip(type);
                continue;
            }

            uint id = 0;
            var name = "";
            var r = new ProtoReader(outer.ReadLengthDelimited());
            while (r.HasMore)
            {
                var (inner, innerType) = r.ReadTag();
                switch (inner)
                {
                    case 1 when innerType == WireType.Varint:
                        id = (uint)r.ReadVarint();
                        break;
                    case 2 when innerType == WireType.LengthDelimited:
                        name = r.ReadString();
                        break;
                    default:
                        r.Skip(innerType);
                        break;
                }
            }

            return id == 0 ? null : new ScrGateway(id, name.Length > 0 ? name : NameOf(id));
        }

        return null;
    }
}

/// <summary>The account's characters (GameAccount.GetToons) and choosing the one to play as.</summary>
public static class ScrToons
{
    /// <summary>GetToons' reply: a repeated field 1 of {1: character ID, 2: name, 3: gateway}.</summary>
    public static IReadOnlyList<ScrToon> Decode(byte[] body)
    {
        var toons = new List<ScrToon>();
        var outer = new ProtoReader(body);
        while (outer.HasMore)
        {
            var (field, type) = outer.ReadTag();
            if (field != 1 || type != WireType.LengthDelimited)
            {
                outer.Skip(type);
                continue;
            }

            uint id = 0, gateway = 0;
            var name = "";
            var r = new ProtoReader(outer.ReadLengthDelimited());
            while (r.HasMore)
            {
                var (inner, innerType) = r.ReadTag();
                switch (inner)
                {
                    case 1 when innerType == WireType.Varint:
                        id = (uint)r.ReadVarint();
                        break;
                    case 2 when innerType == WireType.LengthDelimited:
                        name = r.ReadString();
                        break;
                    case 3 when innerType == WireType.Varint:
                        gateway = (uint)r.ReadVarint();
                        break;
                    default:
                        r.Skip(innerType);
                        break;
                }
            }

            toons.Add(new ScrToon(id, name, gateway));
        }

        return toons;
    }

    /// <summary>
    /// The character on <paramref name="gateway"/> called <paramref name="characterName"/>, or the
    /// first one on that gateway when no name is given. Throws a message fit to show the user when
    /// there isn't one: Invigoration can't create characters yet.
    /// </summary>
    public static ScrToon Choose(IReadOnlyList<ScrToon> toons, uint gateway, string? characterName)
    {
        var onGateway = toons.Where(t => t.Gateway == gateway).ToList();
        var named = string.IsNullOrWhiteSpace(characterName)
            ? onGateway.FirstOrDefault()
            : onGateway.FirstOrDefault(t => t.Name.Equals(characterName.Trim(), StringComparison.OrdinalIgnoreCase));
        if (named is not null)
        {
            return named;
        }

        var elsewhere = toons.Count == 0 ? "This account has no characters at all." : $"Its characters: {string.Join(", ", toons)}.";
        var what = string.IsNullOrWhiteSpace(characterName) ? "no character" : $"no character called {characterName.Trim()}";
        throw new ScrCharacterException(
            $"This Battle.net account has {what} on {ScrGateways.NameOf(gateway)}. {elsewhere} " +
            "Pick one in the bot's settings, or create one on that gateway once in StarCraft: Remastered itself.");
    }
}

/// <summary>The account has no usable character for the chosen gateway. The message is meant for the user.</summary>
public sealed class ScrCharacterException(string message) : InvalidOperationException(message);
