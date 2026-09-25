namespace Invigoration.Sc2.Bsn;

/// <summary>A BSN type's kind, as SC2's embedded metadata numbers them (tag 1 = Array ... 16 = Alias).</summary>
public enum BsnKind
{
    Array,
    ByteString,
    BitArray,
    Blob,
    Bool,
    Choice,
    Enum,
    FourCc,
    Integer,
    Void,
    Optional,
    Float32,
    Float64,
    Struct,
    String,
    Alias,
}

/// <summary>
/// One BSN type from SC2's metadata. Integers, enums, lengths and choice indices are sent as
/// <see cref="BitWidth"/> bits holding the value minus <see cref="Minimum"/> (no bits when -1).
/// A struct's or choice's members are <see cref="Members"/>, named by <see cref="MemberNames"/> and
/// numbered by <see cref="IndexValues"/>; arrays, optionals and aliases wrap <see cref="Element"/>.
/// </summary>
/// <param name="Obfuscated">
/// SC2 may send this struct's fields in another order, with filler bits between them. The order
/// then comes from a <see cref="BsnLayout"/>, recovered from the client's generated code.
/// </param>
public sealed record BsnType(
    uint Id,
    string? Name,
    BsnKind Kind,
    bool Obfuscated,
    int BitWidth,
    long Minimum,
    uint Element,
    long[] IndexValues,
    uint[] Members,
    string?[] MemberNames)
{
    /// <summary>A member's name, or "#index" for an unnamed one.</summary>
    public string MemberName(int position) => MemberNames[position] ?? $"#{IndexValues[position]}";
}

public sealed class BsnSchema
{
    private readonly Dictionary<uint, BsnType> _byId;
    private readonly Dictionary<string, BsnType> _byName;

    public BsnSchema(IEnumerable<BsnType> types)
    {
        _byId = types.ToDictionary(t => t.Id);
        _byName = _byId.Values.Where(t => t.Name is not null).GroupBy(t => t.Name!).ToDictionary(g => g.Key, g => g.First());
    }

    public BsnType this[uint id] => _byId.TryGetValue(id, out var type) ? type : throw new BsnException($"BSN type #{id} isn't in the schema.");

    public BsnType this[string name] => _byName.TryGetValue(name, out var type) ? type : throw new BsnException($"BSN type {name} isn't in the schema.");

    public bool Contains(string name) => _byName.ContainsKey(name);
}

public sealed class BsnException(string message) : Exception(message);
