using System.Text;
using Invigoration.Sc2.Front;

namespace Invigoration.Sc2.Tests;

public class ChallengeMessagesTests
{
    [Fact]
    public void GetValidatedWebAuthUrl_AcceptsAccountBattleNetHttps()
    {
        var challenge = new ChallengeExternalRequest
        {
            PayloadType = "web_auth_url",
            Payload = Encoding.UTF8.GetBytes("https://us.account.battle.net/login?abc=1"),
        };

        var url = challenge.GetValidatedWebAuthUrl();

        Assert.Equal("https", url.Scheme);
        Assert.EndsWith(".account.battle.net", url.Host);
    }

    [Fact]
    public void GetValidatedWebAuthUrl_RejectsNonHttps()
    {
        var challenge = new ChallengeExternalRequest
        {
            PayloadType = "web_auth_url",
            Payload = Encoding.UTF8.GetBytes("http://us.account.battle.net/login"),
        };

        Assert.Throws<InvalidOperationException>(() => challenge.GetValidatedWebAuthUrl());
    }

    [Fact]
    public void GetValidatedWebAuthUrl_RejectsWrongHost()
    {
        var challenge = new ChallengeExternalRequest
        {
            PayloadType = "web_auth_url",
            Payload = Encoding.UTF8.GetBytes("https://evil.example.com/login"),
        };

        Assert.Throws<InvalidOperationException>(() => challenge.GetValidatedWebAuthUrl());
    }

    [Fact]
    public void GetValidatedWebAuthUrl_RejectsHostSuffixSpoof()
    {
        // "notaccount.battle.net" and "account.battle.net.evil.com"-style tricks must not pass.
        var challenge = new ChallengeExternalRequest
        {
            PayloadType = "web_auth_url",
            Payload = Encoding.UTF8.GetBytes("https://us.account.battle.net.evil.com/login"),
        };

        Assert.Throws<InvalidOperationException>(() => challenge.GetValidatedWebAuthUrl());
    }

    [Fact]
    public void GetValidatedWebAuthUrl_RejectsWrongPayloadType()
    {
        var challenge = new ChallengeExternalRequest
        {
            PayloadType = "something_else",
            Payload = Encoding.UTF8.GetBytes("https://us.account.battle.net/login"),
        };

        Assert.Throws<InvalidOperationException>(() => challenge.GetValidatedWebAuthUrl());
    }
}
