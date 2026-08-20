using Invigoration.Sc2.Protobuf;

namespace Invigoration.Sc2.Front;

/// <summary>bgs.protocol.authentication.v1 messages needed to log on with a web (SecretBytes) credential.</summary>
public sealed class LogonRequest
{
    public string? Program { get; init; }
    public string? Platform { get; init; }
    public string? Locale { get; init; }
    public string? Version { get; init; }
    public int? ApplicationVersion { get; init; }
    public bool? AllowLogonQueueNotifications { get; init; }
    public byte[]? CachedWebCredentials { get; init; }
    public string? UserAgent { get; init; }
    public string? DeviceId { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteString(1, Program);
        w.WriteString(2, Platform);
        w.WriteString(3, Locale);
        w.WriteString(5, Version);
        w.WriteInt32(6, ApplicationVersion);
        w.WriteBool(10, AllowLogonQueueNotifications);
        w.WriteBytesField(12, CachedWebCredentials);
        w.WriteString(14, UserAgent);
        w.WriteString(15, DeviceId);
        return w.ToArray();
    }

    public static LogonRequest Decode(byte[] data)
    {
        string? program = null, platform = null, locale = null, version = null, userAgent = null, deviceId = null;
        int? applicationVersion = null;
        bool? allowLogonQueueNotifications = null;
        byte[]? cachedWebCredentials = null;

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: program = r.ReadString(); break;
                case 2: platform = r.ReadString(); break;
                case 3: locale = r.ReadString(); break;
                case 5: version = r.ReadString(); break;
                case 6: applicationVersion = unchecked((int)r.ReadVarint()); break;
                case 10: allowLogonQueueNotifications = r.ReadVarint() != 0; break;
                case 12: cachedWebCredentials = r.ReadLengthDelimited(); break;
                case 14: userAgent = r.ReadString(); break;
                case 15: deviceId = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return new LogonRequest
        {
            Program = program,
            Platform = platform,
            Locale = locale,
            Version = version,
            ApplicationVersion = applicationVersion,
            AllowLogonQueueNotifications = allowLogonQueueNotifications,
            CachedWebCredentials = cachedWebCredentials,
            UserAgent = userAgent,
            DeviceId = deviceId,
        };
    }
}

public sealed class LogonResult
{
    public required uint ErrorCode { get; init; }
    public EntityId? AccountId { get; init; }
    public List<EntityId> GameAccountIds { get; init; } = [];
    public string? Email { get; init; }
    public List<uint> AvailableRegions { get; init; } = [];
    public uint? ConnectedRegion { get; init; }
    public string? BattleTag { get; init; }
    public byte[]? SessionKey { get; init; }
    public bool? RestrictedMode { get; init; }
    public string? ClientId { get; init; }

    public static LogonResult Decode(byte[] data)
    {
        uint errorCode = 0;
        EntityId? accountId = null;
        List<EntityId> gameAccountIds = [];
        string? email = null, battleTag = null, clientId = null;
        List<uint> availableRegions = [];
        uint? connectedRegion = null;
        byte[]? sessionKey = null;
        bool? restrictedMode = null;

        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: errorCode = (uint)r.ReadVarint(); break;
                case 2: accountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                case 3: gameAccountIds.Add(EntityId.Decode(r.ReadLengthDelimited())); break;
                case 4: email = r.ReadString(); break;
                case 5: availableRegions.Add((uint)r.ReadVarint()); break;
                case 6: connectedRegion = (uint)r.ReadVarint(); break;
                case 7: battleTag = r.ReadString(); break;
                case 9: sessionKey = r.ReadLengthDelimited(); break;
                case 10: restrictedMode = r.ReadVarint() != 0; break;
                case 11: clientId = r.ReadString(); break;
                default: r.Skip(type); break;
            }
        }

        return new LogonResult
        {
            ErrorCode = errorCode,
            AccountId = accountId,
            GameAccountIds = gameAccountIds,
            Email = email,
            AvailableRegions = availableRegions,
            ConnectedRegion = connectedRegion,
            BattleTag = battleTag,
            SessionKey = sessionKey,
            RestrictedMode = restrictedMode,
            ClientId = clientId,
        };
    }
}

public sealed class LogonUpdateRequest
{
    public required uint ErrorCode { get; init; }

    public static LogonUpdateRequest Decode(byte[] data)
    {
        uint errorCode = 0;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1)
            {
                errorCode = (uint)r.ReadVarint();
            }
            else
            {
                r.Skip(type);
            }
        }

        return new LogonUpdateRequest { ErrorCode = errorCode };
    }
}

public sealed class VerifyWebCredentialsRequest
{
    public byte[]? WebCredentials { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteBytesField(1, WebCredentials);
        return w.ToArray();
    }
}

public sealed class GenerateWebCredentialsRequest
{
    public uint? Program { get; init; }

    public byte[] Encode()
    {
        var w = new ProtoWriter();
        w.WriteFixed32(1, Program);
        return w.ToArray();
    }
}

public sealed class GenerateWebCredentialsResponse
{
    public byte[]? WebCredentials { get; init; }

    public static GenerateWebCredentialsResponse Decode(byte[] data)
    {
        byte[]? webCredentials = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            if (field == 1)
            {
                webCredentials = r.ReadLengthDelimited();
            }
            else
            {
                r.Skip(type);
            }
        }

        return new GenerateWebCredentialsResponse { WebCredentials = webCredentials };
    }
}

public sealed class GameAccountSelectedRequest
{
    public required uint Result { get; init; }
    public EntityId? GameAccountId { get; init; }

    public static GameAccountSelectedRequest Decode(byte[] data)
    {
        uint result = 0;
        EntityId? gameAccountId = null;
        var r = new ProtoReader(data);
        while (r.HasMore)
        {
            var (field, type) = r.ReadTag();
            switch (field)
            {
                case 1: result = (uint)r.ReadVarint(); break;
                case 2: gameAccountId = EntityId.Decode(r.ReadLengthDelimited()); break;
                default: r.Skip(type); break;
            }
        }

        return new GameAccountSelectedRequest { Result = result, GameAccountId = gameAccountId };
    }
}
