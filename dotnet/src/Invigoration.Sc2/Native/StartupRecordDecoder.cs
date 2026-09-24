using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>
/// Records Battle.net sends around sign-in and channel entry that a chat client
/// has no use for, but must read to the exact bit to find the record after
/// them. Only the widths matter here, so fields are named in comments only.
/// Layouts are the ones written up on docs.bnet.cc's "StarCraft II chat on
/// Sunken"; none has a captured test vector yet.
/// </summary>
public static class StartupRecordDecoder
{
    private const int RecordAddressBits = 96;
    private const int ToonHandleBits = 136;

    /// <summary>Toon/Welcome (Toon 10). Returns the number of unlock entries it listed.</summary>
    public static int SkipToonWelcome(BitReader reader)
    {
        reader.Read(32);
        var handles = (int)reader.Read(4);
        for (var i = 0; i < handles; i++)
        {
            SkipCacheHandle(reader);
            reader.Read(32);
        }

        reader.Read(1);
        reader.Read(32);
        reader.Skip(31);
        reader.Read(32);
        reader.ReadBlob(6);
        reader.Skip(128);

        var realmMaps = (int)reader.Read(3);
        for (var i = 0; i < realmMaps; i++)
        {
            reader.Skip(32 + 32 + 1 + 8);
            reader.Skip((int)reader.Read(5) * 32);
        }

        var unlocks = (int)reader.Read(8);
        for (var i = 0; i < unlocks; i++)
        {
            reader.Skip(4);
            SkipCacheHandle(reader);
        }

        reader.Skip(3);
        reader.Read(16);
        reader.ReadBlob(13);
        reader.ReadBlob(13);
        reader.Read(32);
        return unlocks;
    }

    /// <summary>BillingUpdateNotify (Toon 13).</summary>
    public static void SkipBillingUpdate(BitReader reader)
    {
        reader.Read(32);
        reader.Skip(19);
        reader.SkipOptional(r => r.Read(32));
        reader.Skip(28);
        reader.SkipOptional(r => r.Read(32));
        reader.Read(8);
    }

    /// <summary>PresenceUpdateNotify (Presence 0).</summary>
    public static void SkipPresenceUpdate(BitReader reader)
    {
        reader.Skip(19);
        reader.Read(1);
        reader.Skip(64);
        reader.ReadBlob(11);
        reader.Skip(11);
        foreach (var width in (int[])[32, 32, 16])
        {
            reader.Skip((int)reader.Read(4) * width);
        }

        reader.SkipOptional(r => r.Skip(1 + 32));
        reader.Read(8);
    }

    /// <summary>FieldSpecAnnounce (Presence 1).</summary>
    public static void SkipFieldSpecAnnounce(BitReader reader)
    {
        var count = (int)reader.Read(7);
        for (var i = 0; i < count; i++)
        {
            reader.Skip(3);
            if (reader.Read(1) == 0)
            {
                reader.Read(16);
            }

            reader.Skip(1 + 8 + 32);
        }
    }

    /// <summary>PresenceStatisticsUpdate (Presence 3): up to 63 pairs of a 32-bit key and 64-bit value.</summary>
    public static void SkipPresenceStatistics(BitReader reader) =>
        reader.Skip((int)reader.Read(6) * (32 + 64));

    /// <summary>TemporaryPresenceRequest (Presence 4): one to sixteen toon handles.</summary>
    public static void SkipTemporaryPresenceRequest(BitReader reader) =>
        reader.Skip(((int)reader.Read(4) + 1) * ToonHandleBits);

    /// <summary>
    /// Public channel list response (Chat 22): a 37-bit header, then up to 63 entries of three
    /// numbers (8, 16 and 24 bits), then 16 bits. No names. Which number is the channel ID isn't
    /// confirmed; the 16-bit one is the likely candidate, since public channel IDs are 16-bit.
    /// </summary>
    public static IReadOnlyList<(byte A, ushort B, uint C)> DecodeChannelListResponse(BitReader reader)
    {
        reader.Skip(37);
        var count = (int)reader.Read(6);
        var entries = new List<(byte, ushort, uint)>(count);
        for (var i = 0; i < count; i++)
        {
            entries.Add(((byte)reader.Read(8), (ushort)reader.Read(16), (uint)reader.Read(24)));
        }

        reader.Read(16);
        return entries;
    }

    /// <summary>Channel category descriptions (Chat 24).</summary>
    public static void SkipChannelCategories(BitReader reader)
    {
        reader.Skip((int)reader.Read(6) * (8 + 16 + 16));
        reader.Read(1);
    }

    /// <summary>Channel member counts (Chat 26).</summary>
    public static void SkipChannelMemberCounts(BitReader reader)
    {
        reader.Read(1);
        reader.Skip(27);
        reader.SkipOptional(r => r.Read(32));
        reader.Skip((int)reader.Read(6) * (23 + 32 + 16 + 1));
    }

    /// <summary>AccountBlockNotify (Friends 31).</summary>
    public static void SkipAccountBlocks(BitReader reader)
    {
        reader.SkipOptional(r => r.Read(1));
        var count = (int)reader.Read(7);
        for (var i = 0; i < count; i++)
        {
            reader.Skip(9);
            reader.Read(32);
            reader.SkipOptional(r => r.ReadBlob(7));
            reader.Skip(20);
            reader.SkipOptional(r =>
            {
                r.ReadBlob(8);
                r.ReadBlob(8);
            });
            reader.Read(32);
        }
    }

    /// <summary>CurrentSeason (S2 master 27).</summary>
    public static void SkipCurrentSeason(BitReader reader)
    {
        if (reader.Read(1) != 0)
        {
            reader.Skip(16);
            return;
        }

        reader.Skip(1);
        reader.Skip(reader.ReadCount(7, 100, "Season matchmakers") * (64 + 16 + 19 + 64 + 16 + 20 + 16 + 96 + 82 + 32));
        reader.Skip(reader.ReadCount(9, 448, "Season leagues") * (96 + 8 + 82 + 3));
        reader.Skip(32 + 10);
        reader.SkipOptional(r => r.Skip(32));
        reader.Skip(16);
        reader.SkipOptional(r => r.Skip(16));
        reader.Skip(32 + 16 + 16 + 25);
        reader.SkipOptional(r => r.Skip(16));
        reader.SkipOptional(r => r.Skip(32));
        reader.SkipOptional(r => r.Skip(16));
        reader.Skip(64);
        var configs = reader.ReadCount(9, 448, "Season configurations");
        for (var i = 0; i < configs; i++)
        {
            reader.Skip(82 + 3 + 64);
            reader.SkipOptional(r => r.Skip(32));
        }
    }

    /// <summary>ProfileSettingsAvailable (Profile 4).</summary>
    public static void SkipProfileSettings(BitReader reader)
    {
        reader.Read(2);
        reader.ReadBlob(6);
        reader.Skip(RecordAddressBits);
    }

    /// <summary>ClubSettings (S2 maps 57).</summary>
    public static void SkipClubSettings(BitReader reader)
    {
        reader.ReadBlob(13);
        reader.ReadBlob(13);
        reader.Skip(32 * 5);
    }

    private static void SkipCacheHandle(BitReader reader) => reader.ReadBytes(40, aligned: true);
}
