using System.Text;
using Invigoration.Core.Auth;

namespace Invigoration.Core.Tests;

public class Crc32Tests
{
    [Fact]
    public void Compute_MatchesStandardCheckValue()
    {
        // "123456789" is the standard CRC-32/ISO-HDLC check value: 0xCBF43926.
        var result = Crc32.Compute(Encoding.ASCII.GetBytes("123456789"));

        Assert.Equal(0xCBF43926u, result);
    }

    [Fact]
    public void Compute_EmptyInput_ReturnsZero()
    {
        var result = Crc32.Compute([]);

        Assert.Equal(0u, result);
    }
}

public class BnlsChecksumTests
{
    [Fact]
    public void Compute_MatchesManualCrc32OfConcatenatedSecretAndHexServerCode()
    {
        var expected = Crc32.Compute(Encoding.Latin1.GetBytes("Invigoration000000FF"));

        var result = BnlsChecksum.Compute("Invigoration", 0xFF);

        Assert.Equal(expected, result);
    }
}
