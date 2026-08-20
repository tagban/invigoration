using System.Text;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Decodes MembershipChangeNotify (command 1) — previously undecodable
/// without upstream's ~500KB schema table (core/src/native/schema/wire.rs).
/// This closes that gap using the field-level schema published at
/// https://superioritybot.com/PROTOCOL's type registry (built by the author
/// of ncarrillo/superiority), rather than the schema table itself.
///
/// Two of this record's referenced types — <c>Battlenet::Client::ComSat::TalkerInfo</c>
/// and <c>Battlenet::Toon::Handle</c> — are marked "generated layout" on that
/// site, meaning their declared field order is NOT their wire order. Where
/// upstream's own decode.rs was directly quoted (TalkerInfo: TalkerId decodes
/// before the trailing enabled bit) that quote is authoritative. Where it
/// wasn't (Toon::Handle, reached here via the rare voice/network-diagnostics
/// PlayerTarget::ToonHandle path), this decoder instead matches the already
/// golden-vector-verified <see cref="ChatCommands.ChatWhisper"/> encoder's
/// order for that same type (program_id, region, realm, id) — internally
/// consistent, but unlike the rest of this file, not independently verified
/// against a captured packet, since that specific path is rarely exercised
/// by a non-voice client.
/// </summary>
public static class MembershipChangeDecoder
{
    public static MembershipChangeNotifyRecord Decode(BitReader reader)
    {
        var endOfInitial = reader.Read(1) != 0;
        var channelIndex = (byte)reader.Read(3);
        if (channelIndex >= ChatCommands.ChannelIndexCount)
        {
            throw new InvalidOperationException("Chat membership channel index is invalid.");
        }

        var changeCount = (int)reader.Read(6) + 1;
        var changes = new List<MembershipChange>(changeCount);
        for (var i = 0; i < changeCount; i++)
        {
            changes.Add(DecodeMembershipChange(reader));
        }

        return new MembershipChangeNotifyRecord(endOfInitial, channelIndex, changes);
    }

    private static MembershipChange DecodeMembershipChange(BitReader reader)
    {
        var selector = reader.Read(2);
        return selector switch
        {
            0 => DecodeLeave(reader),
            1 => DecodeJoin(reader),
            2 => DecodeUpdate(reader),
            _ => throw new InvalidOperationException("Chat membership change has an unknown choice."),
        };
    }

    private static MembershipChange.Leave DecodeLeave(BitReader reader)
    {
        var memberHandle = (uint)reader.Read(32);
        var reason = (ushort)reader.Read(16);
        return new MembershipChange.Leave(memberHandle, reason);
    }

    private static MembershipChange.Join DecodeJoin(BitReader reader)
    {
        var memberHandle = (uint)reader.Read(32);
        var presenceId = (uint)reader.Read(32);
        var statusCount = (int)reader.Read(3);
        var statuses = new List<MemberStatus>(statusCount);
        for (var i = 0; i < statusCount; i++)
        {
            statuses.Add(DecodeMemberStatus(reader));
        }

        return new MembershipChange.Join(memberHandle, presenceId, statuses);
    }

    private static MembershipChange.Update DecodeUpdate(BitReader reader)
    {
        var memberHandle = (uint)reader.Read(32);
        var status = DecodeMemberStatus(reader);
        return new MembershipChange.Update(memberHandle, status);
    }

    private static MemberStatus DecodeMemberStatus(BitReader reader)
    {
        var selector = reader.Read(3);
        return selector switch
        {
            0 => DecodeOther(reader),
            1 => DecodeParty(reader),
            2 => new MemberStatus.TalkerNetworkId((byte)reader.Read(8)),
            3 => DecodeTalkerInfo(reader),
            4 => new MemberStatus.VoiceEnabled(reader.Read(1) != 0),
            5 => new MemberStatus.Display(DecodeToonFullName(reader)),
            6 => new MemberStatus.Active(reader.Read(1) != 0),
            7 => new MemberStatus.Sentinel(),
            _ => throw new InvalidOperationException("Chat member status has an unknown choice."),
        };
    }

    private static MemberStatus DecodeOther(BitReader reader)
    {
        var selector = reader.Read(8);
        return selector switch
        {
            0 => new MemberStatus.ClubData((byte)reader.Read(8)),
            1 => new MemberStatus.Licenses(DecodeLicenseIdList(reader)),
            _ => throw new InvalidOperationException("Chat member status 'Other' has an unknown choice."),
        };
    }

    private static IReadOnlyList<uint> DecodeLicenseIdList(BitReader reader)
    {
        var count = (int)reader.Read(16);
        var ids = new List<uint>(count);
        for (var i = 0; i < count; i++)
        {
            ids.Add((uint)reader.Read(32));
        }

        return ids;
    }

    private static MemberStatus.Party DecodeParty(BitReader reader)
    {
        var partyStatus = (byte)reader.Read(2);
        byte? expansionLevel = reader.Read(1) != 0 ? (byte)reader.Read(2) : null;
        var captain = reader.Read(1) != 0;
        return new MemberStatus.Party(partyStatus, expansionLevel, captain);
    }

    /// <summary>TalkerInfo's wire order is (m_id, then a trailing 1-bit m_enabled) — the reverse of its declared field order. Confirmed directly from decode.rs's trace for MemberStatusSingle variant 3.</summary>
    private static MemberStatus.TalkerInfo DecodeTalkerInfo(BitReader reader)
    {
        var id = DecodeTalkerId(reader);
        var enabled = reader.Read(1) != 0;
        return new MemberStatus.TalkerInfo(enabled, id);
    }

    private static TalkerId DecodeTalkerId(BitReader reader)
    {
        var selector = reader.Read(2);
        return selector switch
        {
            0 => new TalkerId.Invalid(),
            1 => new TalkerId.DatagramConnectionEndPoint(DecodePlayerTarget(reader)),
            2 => new TalkerId.Stream((byte)reader.Read(8)),
            _ => throw new InvalidOperationException("Talker id has an unknown choice."),
        };
    }

    /// <summary>EndPoint::Id is a degenerate 0-bit choice with exactly one variant (PlayerTarget), so no selector bits are consumed before it.</summary>
    private static PlayerTarget DecodePlayerTarget(BitReader reader)
    {
        var selector = reader.Read(3);
        return selector switch
        {
            0 => new PlayerTarget.PresenceId((uint)reader.Read(32)),
            1 => new PlayerTarget.ToonName(DecodeToonFullName(reader)),
            2 => new PlayerTarget.AccountMail(DecodeMail(reader)),
            3 => new PlayerTarget.AccountId((uint)reader.Read(32)),
            4 => new PlayerTarget.ProfileRecordAddress((uint)reader.Read(32), reader.Read(64)),
            5 => DecodeToonHandle(reader),
            _ => throw new InvalidOperationException("Player target has an unknown choice."),
        };
    }

    private static PlayerTarget.ToonHandle DecodeToonHandle(BitReader reader)
    {
        var programId = (uint)reader.Read(32);
        var region = (byte)reader.Read(8);
        var realm = (uint)reader.Read(32);
        var id = reader.Read(64);
        return new PlayerTarget.ToonHandle(region, programId, realm, id);
    }

    private static byte[] DecodeMail(BitReader reader)
    {
        var byteCount = (int)reader.Read(9);
        return reader.ReadBytes(byteCount, aligned: true);
    }

    /// <summary>Battlenet::Toon::FullName — "metadata order", so declared order (region, programId, realm, name) IS the wire order. A 5-bit length field biased by +2 (raw 0..23 maps to 2..25 bytes).</summary>
    private static ToonFullName DecodeToonFullName(BitReader reader)
    {
        var region = (byte)reader.Read(8);
        var programId = (uint)reader.Read(32);
        var realm = (uint)reader.Read(32);
        var byteCount = (int)reader.Read(5) + 2;
        var nameBytes = reader.ReadBytes(byteCount, aligned: true);
        var name = Encoding.UTF8.GetString(nameBytes);
        return new ToonFullName(region, programId, realm, name);
    }
}
