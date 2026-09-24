using System.Text;
using Invigoration.Sc2.Protobuf;

namespace Invigoration.Diablo;

/// <summary>One protobuf field as read off the wire: a number for varint and fixed fields, bytes for length-delimited ones.</summary>
public readonly record struct ProtoField(int Number, WireType Type, ulong Value, byte[]? Bytes);

/// <summary>
/// A protobuf message read into its fields without a schema. The chat messages
/// here are only partly mapped and differ a little between D2R, D4 and older
/// replies, so decoders look fields up by number and tolerate either encoding
/// where the games disagree.
/// </summary>
public sealed class ProtoFields
{
    private readonly List<ProtoField> _fields;

    private ProtoFields(List<ProtoField> fields) => _fields = fields;

    public static ProtoFields Parse(byte[] message)
    {
        var fields = new List<ProtoField>();
        var r = new ProtoReader(message);
        while (r.HasMore)
        {
            var (number, type) = r.ReadTag();
            fields.Add(type switch
            {
                WireType.Varint => new ProtoField(number, type, r.ReadVarint(), null),
                WireType.Fixed32 => new ProtoField(number, type, r.ReadFixed32(), null),
                WireType.Fixed64 => new ProtoField(number, type, r.ReadFixed64(), null),
                WireType.LengthDelimited => new ProtoField(number, type, 0, r.ReadLengthDelimited()),
                _ => throw new InvalidOperationException($"Unsupported protobuf wire type {type}."),
            });
        }

        return new ProtoFields(fields);
    }

    public IEnumerable<ProtoField> All(int number) => _fields.Where(f => f.Number == number);

    public ProtoField? First(int number)
    {
        foreach (var field in _fields)
        {
            if (field.Number == number)
            {
                return field;
            }
        }

        return null;
    }

    /// <summary>A varint or fixed-width number, whichever encoding the field uses.</summary>
    public ulong? Number(int number) => First(number) is { Type: not WireType.LengthDelimited } f ? f.Value : null;

    public byte[]? Bytes(int number) => First(number) is { Type: WireType.LengthDelimited } f ? f.Bytes : null;

    public string? String(int number) => Bytes(number) is { } bytes ? Encoding.UTF8.GetString(bytes) : null;

    public ProtoFields? Message(int number) => Bytes(number) is { } bytes ? Parse(bytes) : null;

    public IEnumerable<ProtoFields> Messages(int number) =>
        All(number).Where(f => f.Type == WireType.LengthDelimited).Select(f => Parse(f.Bytes!));
}
