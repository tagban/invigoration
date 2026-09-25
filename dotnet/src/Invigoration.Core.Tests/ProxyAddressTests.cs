using Invigoration.Core.Networking;

namespace Invigoration.Core.Tests;

public class ProxyAddressTests
{
    [Theory]
    [InlineData("socks5://alice:s3cret@203.0.113.5:1080", ProxyProtocol.Socks5, "203.0.113.5", 1080, "alice", "s3cret")]
    [InlineData("socks5h://proxy.example.com:9050", ProxyProtocol.Socks5, "proxy.example.com", 9050, null, null)]
    [InlineData("socks4://bob@10.0.0.2:1080", ProxyProtocol.Socks4, "10.0.0.2", 1080, "bob", null)]
    [InlineData("socks4a://proxy.example.com:1080/", ProxyProtocol.Socks4, "proxy.example.com", 1080, null, null)]
    [InlineData("http://user%40mail:p%3Ass@proxy.example.com:8080", ProxyProtocol.Http, "proxy.example.com", 8080, "user@mail", "p:ss")]
    [InlineData("HTTP://[2001:db8::1]:3128", ProxyProtocol.Http, "2001:db8::1", 3128, null, null)]
    public void Urls(string text, ProxyProtocol protocol, string host, int port, string? user, string? pass) =>
        Assert.Equal(new ProxyAddress(protocol, host, port, user, pass), Parse(text));

    [Theory]
    [InlineData("alice:s3cret@203.0.113.5:1080", "203.0.113.5", 1080, "alice", "s3cret")]
    [InlineData("alice:s3cret:203.0.113.5:1080", "203.0.113.5", 1080, "alice", "s3cret")]
    [InlineData("203.0.113.5:1080:alice:s3cret", "203.0.113.5", 1080, "alice", "s3cret")]
    [InlineData("gate.provider.io:7000:user-zone-us:pw", "gate.provider.io", 7000, "user-zone-us", "pw")]
    [InlineData("user-zone-us:pw:gate.provider.io:7000", "gate.provider.io", 7000, "user-zone-us", "pw")]
    [InlineData("203.0.113.5:1080:alice:12345", "203.0.113.5", 1080, "alice", "12345")]
    [InlineData("alice:12345:203.0.113.5:1080", "203.0.113.5", 1080, "alice", "12345")]
    [InlineData("203.0.113.5:1080:alice:pa:ss", "203.0.113.5", 1080, "alice", "pa:ss")]
    [InlineData("  203.0.113.5:1080  ", "203.0.113.5", 1080, null, null)]
    public void ProviderFormats(string text, string host, int port, string? user, string? pass) =>
        Assert.Equal(new ProxyAddress(null, host, port, user, pass), Parse(text));

    [Theory]
    [InlineData("")]
    [InlineData("https://proxy.example.com:443")]
    [InlineData("ftp://proxy.example.com:21")]
    [InlineData("proxy.example.com")]
    [InlineData("proxy.example.com:99999")]
    [InlineData("2001:db8::1:1080")]
    public void Refuses(string text) => Assert.False(ProxyAddressParser.TryParse(text, out _, out _));

    [Fact]
    public void Describe_LeavesThePasswordOut() =>
        Assert.Equal("SOCKS5 203.0.113.5:1080, signing in as alice with a password",
            ProxyAddressParser.Describe(Parse("alice:s3cret@203.0.113.5:1080"), ProxyProtocol.Socks5));

    [Fact]
    public void Socks4_ConnectsToAnIpv4Address() =>
        Assert.Equal("04011A0BC6336407626F6200", Convert.ToHexString(Socks4Connector.BuildConnectRequest("198.51.100.7", 6667, "bob")));

    [Fact]
    public void Socks4a_SendsAHostNameForTheProxyToResolve() =>
        Assert.Equal("0401177000000001" + "00" + Convert.ToHexString("useast.battle.net"u8) + "00",
            Convert.ToHexString(Socks4Connector.BuildConnectRequest("useast.battle.net", 6000, null)));

    private static ProxyAddress Parse(string text)
    {
        Assert.True(ProxyAddressParser.TryParse(text, out var address, out var error), error);
        return address;
    }
}
