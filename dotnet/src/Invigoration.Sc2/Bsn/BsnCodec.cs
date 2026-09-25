using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Bsn;

/// <summary>A decoded BSN struct: its members by name, in metadata order.</summary>
public sealed class BsnStruct(uint typeId, IReadOnlyList<KeyValuePair<string, object?>> fields)
{
    public uint TypeId { get; } = typeId;

    public IReadOnlyList<KeyValuePair<string, object?>> Fields { get; } = fields;

    public object? this[string name] =>
        Fields.FirstOrDefault(f => f.Key == name) is { Key: not null } field ? field.Value : throw new BsnException($"The struct has no member {name}.");

    public bool Has(string name) => Fields.Any(f => f.Key == name);
}

/// <summary>A decoded BSN choice: which member was sent, and its value.</summary>
public sealed record BsnChoice(long Index, string Name, object? Value);

/// <summary>A value of a Void type.</summary>
public sealed class BsnVoid
{
    public static BsnVoid Value { get; } = new();

    private BsnVoid()
    {
    }
}

/// <summary>
/// How an obfuscated struct goes on the wire: its members' positions (in metadata order) in the
/// order they're sent, each after that many filler bits. Recovered from SC2's generated code.
/// </summary>
public sealed record BsnLayout(params (int Position, int FillerBefore)[] Fields)
{
    /// <summary>Metadata order with no filler: what an unobfuscated struct uses.</summary>
    public static BsnLayout MetadataOrder(int count) => new([.. Enumerable.Range(0, count).Select(p => (p, 0))]);
}

/// <summary>
/// Reads and writes BSN values by their schema, the way SC2's generic codec does (ported from
/// ncarrillo/superiority's core/src/games/sc2/bsn/codec.rs, MIT). Values: long for integers and
/// enums (a u64 above long.MaxValue keeps its bits), uint for FourCCs, string, byte[] for byte
/// strings and blobs, bool, float/double, List&lt;object?&gt; for arrays, null or the value for
/// optionals, <see cref="BsnChoice"/>, <see cref="BsnStruct"/> and <see cref="BsnVoid"/>.
/// An obfuscated struct without a <see cref="BsnLayout"/> is refused rather than guessed.
/// </summary>
public sealed class BsnCodec(BsnSchema schema, IReadOnlyDictionary<string, BsnLayout> layouts)
{
    public BsnSchema Schema { get; } = schema;

    public object? Decode(BitReader reader, string typeName) => Decode(reader, Schema[typeName].Id);

    public void Encode(BitWriter writer, string typeName, object? value) => Encode(writer, Schema[typeName].Id, value);

    public object? Decode(BitReader reader, uint typeId)
    {
        var type = Schema[typeId];
        switch (type.Kind)
        {
            case BsnKind.Alias:
                return Decode(reader, type.Element);
            case BsnKind.Void:
                return BsnVoid.Value;
            case BsnKind.Bool:
                return reader.Read(1) != 0;
            case BsnKind.Integer:
            case BsnKind.Enum:
                return ReadRange(reader, type);
            case BsnKind.FourCc:
                return (uint)reader.Read(32);
            case BsnKind.Float32:
                return BitConverter.Int32BitsToSingle((int)reader.Read(32));
            case BsnKind.Float64:
                return BitConverter.Int64BitsToDouble((long)reader.Read(64));
            case BsnKind.ByteString:
            case BsnKind.Blob:
                return reader.ReadBytes(checked((int)ReadRange(reader, type)), aligned: true);
            case BsnKind.String:
                // Lengths are declared in characters but sent as UTF-8 byte counts: two more bits.
                var length = checked((int)ReadRange(reader, type, extraBits: 2));
                return Encoding.UTF8.GetString(reader.ReadBytes(length, aligned: true));
            case BsnKind.BitArray:
                return reader.ReadRaw(checked((int)ReadRange(reader, type)));
            case BsnKind.Array:
                var count = checked((int)ReadRange(reader, type));
                var items = new List<object?>(count);
                for (var i = 0; i < count; i++)
                {
                    items.Add(Decode(reader, type.Element));
                }

                return items;
            case BsnKind.Optional:
                return reader.Read(1) == 0 ? null : Decode(reader, type.Element);
            case BsnKind.Choice:
                var index = ReadRange(reader, type);
                var position = Array.IndexOf(type.IndexValues, index);
                if (position < 0)
                {
                    throw new BsnException($"{type.Name}: choice {index} isn't declared.");
                }

                return new BsnChoice(index, type.MemberName(position), Decode(reader, type.Members[position]));
            case BsnKind.Struct:
                return DecodeStruct(reader, type);
            default:
                throw new BsnException($"{type.Name}: unknown kind {type.Kind}.");
        }
    }

    public void Encode(BitWriter writer, uint typeId, object? value)
    {
        var type = Schema[typeId];
        switch (type.Kind)
        {
            case BsnKind.Alias:
                Encode(writer, type.Element, value);
                break;
            case BsnKind.Void:
                break;
            case BsnKind.Bool:
                writer.Write(Expect<bool>(type, value) ? 1UL : 0UL, 1);
                break;
            case BsnKind.Integer:
            case BsnKind.Enum:
                WriteRange(writer, type, ToLong(type, value));
                break;
            case BsnKind.FourCc:
                writer.Write(value is string text ? FourCc.Encode(text) : Convert.ToUInt32(value), 32);
                break;
            case BsnKind.Float32:
                writer.Write((uint)BitConverter.SingleToInt32Bits(Expect<float>(type, value)), 32);
                break;
            case BsnKind.Float64:
                writer.Write((ulong)BitConverter.DoubleToInt64Bits(Expect<double>(type, value)), 64);
                break;
            case BsnKind.ByteString:
            case BsnKind.Blob:
                var bytes = Expect<byte[]>(type, value);
                WriteRange(writer, type, bytes.Length);
                writer.WriteBytes(bytes, aligned: true);
                break;
            case BsnKind.String:
                var utf8 = Encoding.UTF8.GetBytes(Expect<string>(type, value));
                WriteRange(writer, type, utf8.Length, extraBits: 2);
                writer.WriteBytes(utf8, aligned: true);
                break;
            case BsnKind.Array:
                var items = Expect<IReadOnlyList<object?>>(type, value);
                WriteRange(writer, type, items.Count);
                foreach (var item in items)
                {
                    Encode(writer, type.Element, item);
                }

                break;
            case BsnKind.Optional:
                writer.Write(value is null ? 0UL : 1UL, 1);
                if (value is not null)
                {
                    Encode(writer, type.Element, value);
                }

                break;
            case BsnKind.Choice:
                var choice = Expect<BsnChoice>(type, value);
                var position = Array.IndexOf(type.IndexValues, choice.Index);
                if (position < 0)
                {
                    throw new BsnException($"{type.Name}: choice {choice.Index} isn't declared.");
                }

                WriteRange(writer, type, choice.Index);
                Encode(writer, type.Members[position], choice.Value);
                break;
            case BsnKind.Struct:
                EncodeStruct(writer, type, Expect<BsnStruct>(type, value));
                break;
            default:
                throw new BsnException($"{type.Name}: {type.Kind} can't be written.");
        }
    }

    /// <summary>Builds a struct of <paramref name="typeName"/> from members given by name; missing ones are filled only when optional or empty.</summary>
    public BsnStruct Struct(string typeName, params (string Name, object? Value)[] members)
    {
        var type = Schema[typeName];
        return new BsnStruct(type.Id, [.. members.Select(m => new KeyValuePair<string, object?>(m.Name, m.Value))]);
    }

    public BsnLayout? LayoutOf(BsnType type) =>
        type.Name is { } name && layouts.TryGetValue(name, out var layout) ? layout
        : type.Obfuscated ? null
        : BsnLayout.MetadataOrder(type.Members.Length);

    private BsnStruct DecodeStruct(BitReader reader, BsnType type)
    {
        var layout = LayoutOf(type) ?? throw new BsnException($"{type.Name} is obfuscated and its wire layout isn't known.");
        var values = new object?[type.Members.Length];
        foreach (var (position, filler) in layout.Fields)
        {
            if (filler > 0)
            {
                reader.Read(filler);
            }

            values[position] = Decode(reader, type.Members[position]);
        }

        return new BsnStruct(type.Id, [.. values.Select((v, i) => new KeyValuePair<string, object?>(type.MemberName(i), v))]);
    }

    private void EncodeStruct(BitWriter writer, BsnType type, BsnStruct value)
    {
        var layout = LayoutOf(type) ?? throw new BsnException($"{type.Name} is obfuscated and its wire layout isn't known.");
        var fillerState = (uint)type.Members.Length;
        foreach (var (position, filler) in layout.Fields)
        {
            if (filler > 0)
            {
                fillerState = NextFiller(fillerState, writer.Position, writer.ToBytes());
                writer.Write(filler >= 32 ? fillerState : fillerState & ((1U << filler) - 1), filler);
            }

            var name = type.MemberName(position);
            var member = type.Members[position];
            if (value.Has(name))
            {
                Encode(writer, member, value[name]);
            }
            else
            {
                Encode(writer, member, DefaultOf(member, type, name));
            }
        }
    }

    /// <summary>What goes in a member the caller left out: nothing for an optional, an empty struct for a memberless one.</summary>
    private object? DefaultOf(uint typeId, BsnType owner, string member)
    {
        var type = Schema[typeId];
        return type.Kind switch
        {
            BsnKind.Alias => DefaultOf(type.Element, owner, member),
            BsnKind.Optional => null,
            BsnKind.Void => BsnVoid.Value,
            BsnKind.Struct when type.Members.Length == 0 => new BsnStruct(type.Id, []),
            _ => throw new BsnException($"{owner.Name}: {member} is required."),
        };
    }

    /// <summary>
    /// SC2's filler bits: a rolling state seeded with the struct's member count, mixed with the
    /// bytes written so far and rotated left 8. The client skips them without checking.
    /// </summary>
    internal static uint NextFiller(uint state, int usedBits, byte[] output)
    {
        var at = usedBits / 8;
        state = usedBits switch
        {
            < 8 => ~state,
            < 16 => state + output[at - 1],
            < 32 => state + BitConverter.ToUInt16(output, at - 2),
            _ => state + BitConverter.ToUInt32(output, at - 4) + BitConverter.ToUInt16(output, at - 2),
        };
        return uint.RotateLeft(state, 8);
    }

    private static long ReadRange(BitReader reader, BsnType type, int extraBits = 0) =>
        type.BitWidth < 0 ? type.Minimum : unchecked((long)reader.Read(type.BitWidth + extraBits) + type.Minimum);

    private static void WriteRange(BitWriter writer, BsnType type, long value, int extraBits = 0)
    {
        if (type.BitWidth < 0)
        {
            if (value != type.Minimum)
            {
                throw new BsnException($"{type.Name} only holds {type.Minimum}.");
            }

            return;
        }

        var width = type.BitWidth + extraBits;
        var raw = unchecked((ulong)(value - type.Minimum));
        if (width < 64 && raw >> width != 0)
        {
            throw new BsnException($"{value} doesn't fit {type.Name} ({width} bits from {type.Minimum}).");
        }

        writer.Write(raw, width);
    }

    private static long ToLong(BsnType type, object? value) => value switch
    {
        long l => l,
        ulong u => unchecked((long)u),
        int i => i,
        uint u => u,
        short s => s,
        ushort u => u,
        byte b => b,
        Enum e => Convert.ToInt64(e),
        _ => throw new BsnException($"{type.Name} needs a number, not {value?.GetType().Name ?? "null"}."),
    };

    private static T Expect<T>(BsnType type, object? value) =>
        value is T typed ? typed : throw new BsnException($"{type.Name} needs a {typeof(T).Name}, not {value?.GetType().Name ?? "null"}.");
}
