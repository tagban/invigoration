namespace Invigoration.Sc2.Native;

/// <summary>Decoded payload of a MembershipChangeNotify record (command 1). Mirrors Battlenet::Client::Chat::MembershipChangeNotify.</summary>
public sealed record MembershipChangeNotifyRecord(bool EndOfInitial, byte ChannelIndex, IReadOnlyList<MembershipChange> Changes);

/// <summary>Battlenet::Chat::MembershipChange — one entry in a MembershipChangeNotify's change list.</summary>
public abstract record MembershipChange
{
    private MembershipChange()
    {
    }

    public sealed record Leave(uint MemberHandle, ushort Reason) : MembershipChange;

    public sealed record Join(uint MemberHandle, uint PresenceId, IReadOnlyList<MemberStatus> Statuses) : MembershipChange;

    public sealed record Update(uint MemberHandle, MemberStatus Status) : MembershipChange;
}

/// <summary>Battlenet::Chat::MemberStatusSingle — the 8-way status choice attached to a chat member.</summary>
public abstract record MemberStatus
{
    private MemberStatus()
    {
    }

    public sealed record ClubData(byte Rank) : MemberStatus;

    public sealed record Licenses(IReadOnlyList<uint> LicenseIds) : MemberStatus;

    public sealed record Party(byte PartyStatus, byte? ExpansionLevel, bool Captain) : MemberStatus;

    public sealed record TalkerNetworkId(byte Id) : MemberStatus;

    public sealed record TalkerInfo(bool Enabled, TalkerId Id) : MemberStatus;

    public sealed record VoiceEnabled(bool Enabled) : MemberStatus;

    public sealed record Display(ToonFullName ToonName) : MemberStatus;

    public sealed record Active(bool Value) : MemberStatus;

    public sealed record Sentinel : MemberStatus;
}

/// <summary>Battlenet::Toon::FullName — region/program/realm-qualified toon name. NOT the same wire layout as Toon::Handle; see <see cref="PlayerTarget.ToonHandle"/>.</summary>
public sealed record ToonFullName(byte Region, uint ProgramId, uint Realm, string Name);

/// <summary>Battlenet::Client::ComSat::TalkerId::Id.</summary>
public abstract record TalkerId
{
    private TalkerId()
    {
    }

    public sealed record Invalid : TalkerId;

    public sealed record DatagramConnectionEndPoint(PlayerTarget Target) : TalkerId;

    public sealed record Stream(byte StreamId) : TalkerId;
}

/// <summary>Battlenet::Client::Defines::PlayerTarget, reached only through the rare voice/network-diagnostics TalkerInfo status path.</summary>
public abstract record PlayerTarget
{
    private PlayerTarget()
    {
    }

    public sealed record PresenceId(uint Id) : PlayerTarget;

    public sealed record ToonName(ToonFullName Name) : PlayerTarget;

    public sealed record AccountMail(byte[] Mail) : PlayerTarget;

    public sealed record AccountId(uint Id) : PlayerTarget;

    public sealed record ProfileRecordAddress(uint Label, ulong RecordId) : PlayerTarget;

    /// <summary>
    /// Battlenet::Toon::Handle. Its metadata declares field order (region,
    /// programId, realm, id), but this type uses a "generated layout" — the
    /// actual wire order differs. Confirmed via the already-golden-vector-
    /// verified chat_whisper native builder (core/src/native/protocol.rs),
    /// which writes program_id, then region, then realm, then id for this
    /// exact type. See <see cref="MembershipChangeDecoder"/>.
    /// </summary>
    public sealed record ToonHandle(byte Region, uint ProgramId, uint Realm, ulong Id) : PlayerTarget;
}
