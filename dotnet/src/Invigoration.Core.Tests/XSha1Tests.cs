using System.Buffers.Binary;
using System.Text;
using Invigoration.Core.Auth;

namespace Invigoration.Core.Tests;

public class XSha1Tests
{
    /// <summary>
    /// From wjlafrance/broken-sha1's main.c (a C port of Rob Paveza's
    /// MBNCSUtil), printed there via printf("%08x%08x%08x%08x%08x", ...) on
    /// the five raw state words — i.e. each word's big-endian *value*
    /// representation, not wire-format bytes. XSha1.Hash's actual byte
    /// output is little-endian per word (matching Battle.net's wire format),
    /// so this reads each 4-byte group back out as little-endian before
    /// formatting, to compare values rather than raw bytes.
    /// </summary>
    [Fact]
    public void Hash_MatchesBrokenSha1CReferenceVector()
    {
        var result = XSha1.Hash(Encoding.ASCII.GetBytes("1234567890"));

        Assert.Equal("99f0fab8b5b4523e0d58e5efe126fa5f12633b4b", ToBigEndianWordHex(result));
    }

    /// <summary>
    /// From Davnit/bncs.py's bsha.py docstring, which documents the output
    /// of that library's own .hexdigest() — i.e. the actual little-endian
    /// wire bytes hex-encoded directly, the same convention XSha1.Hash uses.
    /// </summary>
    [Fact]
    public void Hash_MatchesBncsPyReferenceVector()
    {
        var result = XSha1.Hash(Encoding.ASCII.GetBytes("The quick brown fox jumps over the lazy dog"));

        Assert.Equal("a0db6e70616033a7b5fdda37cee2d43f2da10288", Convert.ToHexString(result).ToLowerInvariant());
    }

    [Fact]
    public void Hash_MultipleSegments_MatchesEquivalentSingleSegment()
    {
        var whole = XSha1.Hash(Encoding.ASCII.GetBytes("1234567890"));
        var split = XSha1.Hash(Encoding.ASCII.GetBytes("12345"), Encoding.ASCII.GetBytes("67890"));

        Assert.Equal(whole, split);
    }

    private static string ToBigEndianWordHex(byte[] digest)
    {
        var hex = new StringBuilder();
        for (var i = 0; i < digest.Length; i += 4)
        {
            var value = BinaryPrimitives.ReadUInt32LittleEndian(digest.AsSpan(i, 4));
            hex.Append(value.ToString("x8"));
        }

        return hex.ToString();
    }
}
