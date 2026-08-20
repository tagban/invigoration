using System.Text;

namespace Invigoration.Sc2.Wire;

/// <summary>
/// FNV-1a 32-bit hash of a fully-qualified bgs.protocol service name, used
/// as the Front RPC header's service_hash field. Ported from
/// core/src/wire/protobuf.rs's service_hash().
/// </summary>
public static class ServiceHash
{
    public static uint Compute(string serviceName)
    {
        var hash = 0x811c9dc5u;
        foreach (var b in Encoding.UTF8.GetBytes(serviceName))
        {
            hash ^= b;
            hash *= 0x01000193u;
        }

        return hash;
    }
}
