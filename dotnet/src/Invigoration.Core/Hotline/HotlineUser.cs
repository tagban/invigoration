using System.Buffers.Binary;
using System.Text;

namespace Invigoration.Core.Hotline;

public sealed record HotlineUser(ushort UserId, ushort IconId, ushort Flags, string Name)
{
    public bool IsAdmin => (Flags & (1 << HotlineUserFlagBits.Admin)) != 0;

    /// <summary>
    /// Decodes a single UserNameWithInfo(300) field's packed binary payload —
    /// userId(2)/iconId(2)/flags(2)/nameLength(2)/name(N), all big-endian, per
    /// Hotline-Navigator's users.rs. Confirmed live this isn't the whole story, though: an older
    /// real server (a "1.2.3-modern server structure" gap flagged directly by the user) sends no
    /// explicit name-length field at all — just userId/iconId/flags followed by the name filling
    /// the rest of the field, which made the newer-format read treat part of the name's own bytes
    /// as a bogus (too-large) length and crash. Falls back to that older 3-field shape whenever
    /// the newer read's declared length doesn't actually fit the data, rather than assuming every
    /// server speaks the newer variant.
    /// </summary>
    public static HotlineUser Parse(byte[] data)
    {
        var userId = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(0));
        var iconId = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(2));
        var flags = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(4));

        if (data.Length >= 8)
        {
            var nameLength = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(6));
            if (8 + nameLength <= data.Length)
            {
                return new HotlineUser(userId, iconId, flags, Encoding.UTF8.GetString(data.AsSpan(8, nameLength)));
            }
        }

        var name = data.Length > 6 ? Encoding.UTF8.GetString(data.AsSpan(6)) : "";
        return new HotlineUser(userId, iconId, flags, name);
    }
}
