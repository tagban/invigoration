using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>
/// The Front RPC envelope (bgs.protocol.Header). Every Front frame is
/// `u16 BE header_length ++ Header ++ body`; see <see cref="FrontFrame"/>.
/// </summary>
public sealed class Header
{
    public required uint ServiceId { get; init; }
    public uint? MethodId { get; init; }
    public required uint Token { get; init; }
    public ulong? ObjectId { get; init; }
    public uint? Size { get; init; }
    public uint? Status { get; init; }
    public List<ErrorInfo> Errors { get; init; } = [];
    public ulong? Timeout { get; init; }
    public bool? IsResponse { get; init; }
    public List<ProcessId> ForwardTargets { get; init; } = [];
    public uint? ServiceHash { get; init; }
    public string? ClientId { get; init; }
    public List<FanoutTarget> FanoutTargets { get; init; } = [];
    public List<string> ClientIdFanoutTargets { get; init; } = [];
    public byte[]? ClientRecord { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteUInt32(1, ServiceId);
        w.WriteUInt32(2, MethodId);
        w.WriteUInt32(3, Token);
        w.WriteUInt64(4, ObjectId);
        w.WriteUInt32(5, Size);
        w.WriteUInt32(6, Status);
        foreach (var e in Errors) w.WriteBytesField(7, e.Encode());
        w.WriteUInt64(8, Timeout);
        w.WriteBool(9, IsResponse);
        foreach (var t in ForwardTargets) w.WriteBytesField(10, t.Encode());
        w.WriteFixed32(11, ServiceHash);
        w.WriteString(13, ClientId);
        foreach (var f in FanoutTargets) w.WriteBytesField(14, f.Encode());
        foreach (var s in ClientIdFanoutTargets) w.WriteString(15, s);
        w.WriteBytesField(16, ClientRecord);
        return w.ToArray();
    }

    public static Header Decode(byte[] data)
    {
        uint serviceId = 0, token = 0;
        uint? methodId = null, size = null, status = null, serviceHash = null;
        ulong? objectId = null, timeout = null;
        bool? isResponse = null;
        string? clientId = null;
        byte[]? clientRecord = null;
        List<ErrorInfo> errors = [];
        List<ProcessId> forwardTargets = [];
        List<FanoutTarget> fanoutTargets = [];
        List<string> clientIdFanoutTargets = [];

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: serviceId = (uint)r.ReadVarint(); break;
                case 2: methodId = (uint)r.ReadVarint(); break;
                case 3: token = (uint)r.ReadVarint(); break;
                case 4: objectId = r.ReadVarint(); break;
                case 5: size = (uint)r.ReadVarint(); break;
                case 6: status = (uint)r.ReadVarint(); break;
                case 7: errors.Add(ErrorInfo.Decode(r.ReadLengthDelimited())); break;
                case 8: timeout = r.ReadVarint(); break;
                case 9: isResponse = r.ReadVarint() != 0; break;
                case 10: forwardTargets.Add(ProcessId.Decode(r.ReadLengthDelimited())); break;
                case 11: serviceHash = r.ReadFixed32(); break;
                case 13: clientId = r.ReadString(); break;
                case 14: fanoutTargets.Add(FanoutTarget.Decode(r.ReadLengthDelimited())); break;
                case 15: clientIdFanoutTargets.Add(r.ReadString()); break;
                case 16: clientRecord = r.ReadLengthDelimited(); break;
                default: r.Skip(type); break;
            }
        }

        return new Header
        {
            ServiceId = serviceId,
            MethodId = methodId,
            Token = token,
            ObjectId = objectId,
            Size = size,
            Status = status,
            Errors = errors,
            Timeout = timeout,
            IsResponse = isResponse,
            ForwardTargets = forwardTargets,
            ServiceHash = serviceHash,
            ClientId = clientId,
            FanoutTargets = fanoutTargets,
            ClientIdFanoutTargets = clientIdFanoutTargets,
            ClientRecord = clientRecord,
        };
    }
}
