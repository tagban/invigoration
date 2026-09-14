using System.Net;
using System.Net.Sockets;
using System.Text;

namespace Invigoration.Core.Tests;

/// <summary>A small hand-written d2-equipment.json in Command Center's shape, plus a loopback BNFTP server — shared by the equipment map, store and client tests.</summary>
internal static class D2EquipmentTestData
{
    public const string Json = """
    {
      "format": "bnetcc-d2-equipment",
      "version": 1,
      "game": "Diablo II 1.14d",
      "portrait": { "length": 33, "class_offset": 13, "level_offset": 25, "none": 255 },
      "slots": [
        { "index": 0, "name": "head", "component": "HD", "offset": 2, "tint_offset": 14,
          "values": [ { "value": 57, "code": "cap", "items": [ { "code": "cap", "name": "Cap" }, { "code": "xap", "name": "War Hat" }, { "code": "uap", "name": "Shako" } ] } ] },
        { "index": 1, "name": "torso", "component": "TR", "offset": 3, "tint_offset": 15,
          "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 2, "name": "legs", "component": "LG", "offset": 4, "tint_offset": 16, "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 3, "name": "right_arm", "component": "RA", "offset": 5, "tint_offset": 17, "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 4, "name": "left_arm", "component": "LA", "offset": 6, "tint_offset": 18, "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 5, "name": "right_hand", "component": "RH", "offset": 7, "tint_offset": 19,
          "values": [ { "value": 51, "code": "ob1", "items": [ { "code": "ob1", "name": "Eagle Orb" }, { "code": "ob6", "name": "Glowing Orb" }, { "code": "oba", "name": "Eldritch Orb" }, { "code": "obf", "name": "Dimensional Shard" } ] },
                      { "value": 70, "code": "lxb", "items": [ { "code": "lxb", "name": "Light Crossbow" } ] } ] },
        { "index": 6, "name": "left_hand", "component": "LH", "offset": 8, "tint_offset": 20,
          "values": [ { "value": 41, "code": "sbw", "items": [ { "code": "sbw", "name": "Short Bow" } ] },
                      { "value": 70, "code": "lxb", "items": [ { "code": "lxb", "name": "Light Crossbow" } ] } ] },
        { "index": 7, "name": "shield", "component": "SH", "offset": 9, "tint_offset": 21,
          "values": [ { "value": 81, "code": "kit", "items": [ { "code": "kit", "name": "Kite Shield" }, { "code": "uit", "name": "Monarch" } ] } ] },
        { "index": 8, "name": "right_shoulder", "component": "S1", "offset": 10, "tint_offset": 22, "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 2, "code": "med", "weight": "medium" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 9, "name": "left_shoulder", "component": "S2", "offset": 11, "tint_offset": 23, "values": [ { "value": 1, "code": "lit", "weight": "light" }, { "value": 2, "code": "med", "weight": "medium" }, { "value": 3, "code": "hvy", "weight": "heavy" } ] },
        { "index": 10, "name": "special", "component": "S3", "offset": 12, "tint_offset": 24, "values": [] }
      ],
      "body_armor": { "sets": [ { "parts": [1, 1, 1, 1, 2, 2], "items": [ { "code": "qui", "name": "Quilted Armor" }, { "code": "uui", "name": "Dusk Shroud" } ] } ] },
      "not_drawn": { "items": [ { "code": "ci0", "name": "Circlet" } ] },
      "tints": { "colors": [ "White", "Light Grey", "Dark Grey", "Black", "Light Blue", "Dark Blue", "Crystal Blue" ] },
      "graphics": []
    }
    """;

    public static byte[] Bytes => Encoding.UTF8.GetBytes(Json);

    public static byte[] Gear(params (int Slot, int Value)[] worn)
    {
        var gear = Enumerable.Repeat((byte)0xFF, 11).ToArray();
        foreach (var (slot, value) in worn)
        {
            gear[slot] = (byte)value;
        }

        return gear;
    }

    /// <summary>A D2 realm statstring wearing <paramref name="gear"/> (11 bytes) with <paramref name="tints"/> (11 bytes, default untinted).</summary>
    public static string Statstring(byte[] gear, byte[]? tints = null)
    {
        var p = Enumerable.Repeat((char)0xFF, 33).ToArray();
        p[0] = (char)0x84;
        p[1] = (char)0x80;
        p[13] = (char)2;
        p[25] = (char)80;
        p[26] = (char)0xA0;
        p[27] = (char)0x9E;
        for (var i = 0; i < 11; i++)
        {
            p[2 + i] = (char)gear[i];
            p[14 + i] = (char)(tints?[i] ?? 0xFF);
        }

        return "PX2DUSEast,Kilua," + new string(p);
    }

    /// <summary>
    /// A one-shot BNFTP server on loopback. <paramref name="reply"/> gets the requested file name and
    /// returns the file's bytes, or null to close without replying (how a server says "no such file").
    /// <paramref name="announcedSize"/> overrides the size in the header, for lying-server tests.
    /// </summary>
    public static async Task<(int Port, Task<string> RequestedName)> ServeOnceAsync(Func<string, byte[]?> reply, uint? announcedSize = null, int? sendOnly = null)
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var port = ((IPEndPoint)listener.LocalEndpoint).Port;

        async Task<string> Run()
        {
            try
            {
                using var client = await listener.AcceptTcpClientAsync();
                await using var stream = client.GetStream();
                var head = new byte[3];
                await stream.ReadExactlyAsync(head);
                var body = new byte[BitConverter.ToUInt16(head, 1) - 2];
                await stream.ReadExactlyAsync(body);
                // After the length: version (2), platform (4), product (4), banner id (4), extension (4), start (4), file time (8).
                var nameEnd = Array.IndexOf(body, (byte)0, 30);
                var name = Encoding.ASCII.GetString(body, 30, nameEnd - 30);

                if (reply(name) is { } data)
                {
                    var nameBytes = Encoding.ASCII.GetBytes(name);
                    var header = new List<byte>();
                    header.AddRange(BitConverter.GetBytes((ushort)(2 + 2 + 4 + 4 + 4 + 8 + nameBytes.Length + 1)));
                    header.AddRange(BitConverter.GetBytes((ushort)0));
                    header.AddRange(BitConverter.GetBytes(announcedSize ?? (uint)data.Length));
                    header.AddRange(BitConverter.GetBytes(0u));
                    header.AddRange(BitConverter.GetBytes(0u));
                    header.AddRange(BitConverter.GetBytes(134335450650000000L));
                    header.AddRange(nameBytes);
                    header.Add(0);
                    await stream.WriteAsync(header.ToArray());
                    await stream.WriteAsync(data.AsMemory(0, sendOnly ?? data.Length));
                }

                return name;
            }
            finally
            {
                listener.Stop();
            }
        }

        return (port, Run());
    }
}
