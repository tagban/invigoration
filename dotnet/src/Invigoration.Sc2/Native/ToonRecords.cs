using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>Decoded payload of Toon/6 ToonSelected — confirms the toon chosen by Toon/5 ToonSelect (<see cref="ChatCommands.ToonSelect"/>).</summary>
public sealed record ToonSelectedRecord(uint RecordLabel, ulong RecordId, ToonHandleValue Handle, uint Realm, int LastLogon, string ToonName);

/// <summary>Battlenet::Toon::Handle — note the wire order (program, region, realm, id) does NOT match its type-registry field declaration order; confirmed independently from both the SC2Docs service-endpoint listing and the golden-vector-verified chat_whisper native builder.</summary>
public sealed record ToonHandleValue(uint Program, byte Region, uint Realm, ulong Id);

public static class ToonRecordDecoder
{
    /// <summary>Decodes Toon/6 ToonSelected. Field order here is service-endpoint "generated order" (true wire order), not the type-registry declaration order.</summary>
    public static ToonSelectedRecord DecodeToonSelected(BitReader reader)
    {
        var recordLabel = (uint)reader.Read(32);
        var recordId = reader.Read(64);
        var handle = DecodeToonHandle(reader);
        var realm = (uint)reader.Read(32);
        var lastLogon = unchecked((int)reader.Read(32));
        var byteCount = (int)reader.Read(7) + 2;
        var nameBytes = reader.ReadBytes(byteCount, aligned: true);
        var toonName = System.Text.Encoding.UTF8.GetString(nameBytes);
        return new ToonSelectedRecord(recordLabel, recordId, handle, realm, lastLogon, toonName);
    }

    private static ToonHandleValue DecodeToonHandle(BitReader reader)
    {
        var program = (uint)reader.Read(32);
        var region = (byte)reader.Read(8);
        var realm = (uint)reader.Read(32);
        var id = reader.Read(64);
        return new ToonHandleValue(program, region, realm, id);
    }
}
