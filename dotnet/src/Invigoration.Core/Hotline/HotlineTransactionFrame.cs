using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

/// <summary>A single field/parameter within a transaction's body — Hotline's own TLV shape (2-byte type, 2-byte length, then raw data).</summary>
public sealed record HotlineField(ushort Type, byte[] Data)
{
    public HotlineField(HotlineFieldType type, byte[] data) : this((ushort)type, data)
    {
    }

    public HotlineField(HotlineFieldType type, string value) : this((ushort)type, Encoding.UTF8.GetBytes(value))
    {
    }

    public HotlineField(HotlineFieldType type, ushort value) : this((ushort)type, ToBigEndian(value))
    {
    }

    public HotlineField(HotlineFieldType type, uint value) : this((ushort)type, ToBigEndian(value))
    {
    }

    public HotlineField(HotlineFieldType type, ulong value) : this((ushort)type, ToBigEndian(value))
    {
    }

    /// <summary>
    /// Decodes as text — UTF-8, not the classic Mac OS Roman encoding real Hotline clients
    /// historically used. A known, accepted simplification (same shape as HotlineTrackerClient's
    /// server-name decoding): plain-ASCII names/chat round-trip correctly either way, and adding a
    /// full Mac Roman code page table is real extra complexity for no benefit to the actual users
    /// of this app.
    /// </summary>
    public string AsString() => Encoding.UTF8.GetString(Data);

    public ushort AsUInt16() => Data.Length >= 2 ? BinaryPrimitives.ReadUInt16BigEndian(Data) : (ushort)0;

    public uint AsUInt32() => Data.Length >= 4 ? BinaryPrimitives.ReadUInt32BigEndian(Data) : 0u;

    public ulong AsUInt64() => Data.Length >= 8 ? BinaryPrimitives.ReadUInt64BigEndian(Data) : 0ul;

    /// <summary>A single-byte boolean field (e.g. DATA_HISTORY_HAS_MORE) — nonzero is true, matching the "1 if ... 0 otherwise" convention documented for that field.</summary>
    public bool AsBool() => Data.Length > 0 && Data[0] != 0;

    private static byte[] ToBigEndian(ushort value)
    {
        var bytes = new byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, value);
        return bytes;
    }

    private static byte[] ToBigEndian(uint value)
    {
        var bytes = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(bytes, value);
        return bytes;
    }

    private static byte[] ToBigEndian(ulong value)
    {
        var bytes = new byte[8];
        BinaryPrimitives.WriteUInt64BigEndian(bytes, value);
        return bytes;
    }
}

/// <summary>
/// One Hotline transaction — a 20-byte header (flags, is_reply, type, id, error_code, total_size,
/// data_size, all big-endian except the two single-byte flags) followed by a field count and that
/// many <see cref="HotlineField"/> entries. Ported byte-for-byte from Hotline-Navigator's
/// transaction.rs (fetched directly, not guessed). Split/multi-packet transactions (total_size !=
/// data_size, used for large file transfers) aren't supported — every transaction this client
/// sends or expects to receive (login, chat, user list) fits in a single packet.
/// </summary>
public sealed class HotlineTransactionFrame
{
    public byte Flags { get; init; }
    public bool IsReply { get; init; }
    public ushort Type { get; init; }
    public uint Id { get; init; }
    public uint ErrorCode { get; init; }
    public List<HotlineField> Fields { get; init; } = [];

    public static HotlineTransactionFrame Create(HotlineTransactionType type, uint id, params HotlineField[] fields) =>
        new() { Type = (ushort)type, Id = id, Fields = [.. fields] };

    public static HotlineTransactionFrame CreateReply(uint replyToId, uint errorCode = 0, params HotlineField[] fields) =>
        new() { Type = (ushort)HotlineTransactionType.Reply, Id = replyToId, IsReply = true, ErrorCode = errorCode, Fields = [.. fields] };

    public HotlineField? Field(HotlineFieldType type) => Fields.FirstOrDefault(f => f.Type == (ushort)type);

    public byte[] Encode()
    {
        var body = new List<byte>();
        AppendUInt16(body, (ushort)Fields.Count);
        foreach (var field in Fields)
        {
            AppendUInt16(body, field.Type);
            AppendUInt16(body, (ushort)field.Data.Length);
            body.AddRange(field.Data);
        }

        var frame = new byte[HotlineConstants.TransactionHeaderSize + body.Count];
        frame[0] = Flags;
        frame[1] = (byte)(IsReply ? 1 : 0);
        BinaryPrimitives.WriteUInt16BigEndian(frame.AsSpan(2), Type);
        BinaryPrimitives.WriteUInt32BigEndian(frame.AsSpan(4), Id);
        BinaryPrimitives.WriteUInt32BigEndian(frame.AsSpan(8), ErrorCode);
        BinaryPrimitives.WriteUInt32BigEndian(frame.AsSpan(12), (uint)body.Count);
        BinaryPrimitives.WriteUInt32BigEndian(frame.AsSpan(16), (uint)body.Count);
        body.CopyTo(frame, HotlineConstants.TransactionHeaderSize);
        return frame;
    }

    /// <summary>Given the bytes buffered so far (header may not have arrived yet), returns the full frame length (header + body) once known — the exact shape FramedTcpClient.TryGetFrameLength needs.</summary>
    public static int? TryGetFrameLength(IReadOnlyList<byte> buffer)
    {
        if (buffer.Count < HotlineConstants.TransactionHeaderSize)
        {
            return null;
        }

        var header = new byte[HotlineConstants.TransactionHeaderSize];
        for (var i = 0; i < header.Length; i++)
        {
            header[i] = buffer[i];
        }

        var totalSize = BinaryPrimitives.ReadUInt32BigEndian(header.AsSpan(12));
        return HotlineConstants.TransactionHeaderSize + (int)totalSize;
    }

    /// <summary>Decodes one complete, exactly-sized frame — the shape FramedTcpClient.PacketReceived hands over once TryGetFrameLength says a full frame is buffered.</summary>
    public static HotlineTransactionFrame Decode(byte[] frame)
    {
        var result = new HotlineTransactionFrame
        {
            Flags = frame[0],
            IsReply = frame[1] != 0,
            Type = BinaryPrimitives.ReadUInt16BigEndian(frame.AsSpan(2)),
            Id = BinaryPrimitives.ReadUInt32BigEndian(frame.AsSpan(4)),
            ErrorCode = BinaryPrimitives.ReadUInt32BigEndian(frame.AsSpan(8)),
        };

        var offset = HotlineConstants.TransactionHeaderSize;
        var fieldCount = BinaryPrimitives.ReadUInt16BigEndian(frame.AsSpan(offset));
        offset += 2;

        for (var i = 0; i < fieldCount; i++)
        {
            var fieldType = BinaryPrimitives.ReadUInt16BigEndian(frame.AsSpan(offset));
            var fieldSize = BinaryPrimitives.ReadUInt16BigEndian(frame.AsSpan(offset + 2));
            offset += 4;
            var data = frame[offset..(offset + fieldSize)];
            offset += fieldSize;
            result.Fields.Add(new HotlineField(fieldType, data));
        }

        return result;
    }

    private static void AppendUInt16(List<byte> buffer, ushort value)
    {
        var bytes = new byte[2];
        BinaryPrimitives.WriteUInt16BigEndian(bytes, value);
        buffer.AddRange(bytes);
    }
}
