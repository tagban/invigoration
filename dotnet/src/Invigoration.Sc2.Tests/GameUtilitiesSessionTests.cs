using Invigoration.Sc2.Front;
using Attribute = Invigoration.Sc2.Front.Attribute;

namespace Invigoration.Sc2.Tests;

public class GameUtilitiesSessionTests
{
    [Fact]
    public void BuildProcessClientRequest_EncodesExpectedAttributesAndGameAccountId()
    {
        var gameAccountId = new EntityId { High = 1, Low = 2 };
        var sessionKey = new byte[64];
        for (var i = 0; i < 64; i++)
        {
            sessionKey[i] = (byte)i;
        }

        var request = GameUtilitiesSession.BuildProcessClientRequest(gameAccountId, sessionKey);
        var decoded = ClientRequest.Decode(request.Encode());

        Assert.Equal(1ul, decoded.GameAccountId!.High);
        Assert.Null(decoded.AccountId);
        Assert.Null(decoded.Program);
        Assert.Equal(4, decoded.Attributes.Count);
        Assert.Equal("0.0.1", decoded.Attributes[0].Value.StringValue);
        Assert.Equal("US", decoded.Attributes[1].Value.StringValue);
        Assert.Equal(sessionKey, decoded.Attributes[2].Value.BlobValue);
        Assert.Equal("enUS", decoded.Attributes[3].Value.StringValue);
    }

    [Fact]
    public void BuildProcessClientRequest_RejectsWrongLengthSessionKey()
    {
        var gameAccountId = new EntityId { High = 1, Low = 2 };

        Assert.Throws<ArgumentException>(() => GameUtilitiesSession.BuildProcessClientRequest(gameAccountId, new byte[32]));
    }

    [Fact]
    public void ParseHandoff_ExtractsAllRequiredAttributes()
    {
        var sessionKey = new byte[64];
        var response = new ClientResponse
        {
            Attributes =
            [
                new Attribute { Name = "address", Value = new Variant { StringValue = "us.actual.battle.net:1119" } },
                new Attribute { Name = "session_key", Value = new Variant { BlobValue = sessionKey } },
                new Attribute { Name = "account_region", Value = new Variant { UintValue = 1 } },
                new Attribute { Name = "game_account_name", Value = new Variant { StringValue = "Tagban" } },
                new Attribute { Name = "account_mail", Value = new Variant { StringValue = "player@example.com" } },
                new Attribute { Name = "logon_response", Value = new Variant { BlobValue = [1, 2, 3] } },
            ],
        };

        var handoff = GameUtilitiesSession.ParseHandoff(response);

        Assert.Equal("us.actual.battle.net:1119", handoff.Address);
        Assert.Equal(sessionKey, handoff.SessionKey);
        Assert.Equal(1, handoff.AccountRegion);
        Assert.Equal("Tagban", handoff.GameAccountName);
        Assert.Equal("player@example.com", handoff.AccountMail);
        Assert.Equal(new byte[] { 1, 2, 3 }, handoff.LogonResponse);
    }

    [Fact]
    public void ParseHandoff_MissingRequiredAttribute_Throws()
    {
        var response = new ClientResponse
        {
            Attributes = [new Attribute { Name = "address", Value = new Variant { StringValue = "x" } }],
        };

        Assert.Throws<InvalidOperationException>(() => GameUtilitiesSession.ParseHandoff(response));
    }

    [Theory]
    [InlineData("us.sunken.battle.net", 1119, "us.sunken.battle.net", 1119)]
    [InlineData("us.sunken.battle.net:9999", 1119, "us.sunken.battle.net", 9999)]
    [InlineData("[::1]", 1119, "::1", 1119)]
    [InlineData("[::1]:9999", 1119, "::1", 9999)]
    public void Endpoint_ParsesHostAndPort(string address, int defaultPort, string expectedHost, int expectedPort)
    {
        var handoff = new SunkenHandoff(address, new byte[64], 1, "Tagban", "player@example.com", null);

        var (host, port) = handoff.Endpoint(defaultPort);

        Assert.Equal(expectedHost, host);
        Assert.Equal(expectedPort, port);
    }

    [Fact]
    public void ParseHandoff_WrongLengthSessionKey_Throws()
    {
        var response = new ClientResponse
        {
            Attributes =
            [
                new Attribute { Name = "address", Value = new Variant { StringValue = "x" } },
                new Attribute { Name = "session_key", Value = new Variant { BlobValue = new byte[10] } },
                new Attribute { Name = "account_region", Value = new Variant { UintValue = 1 } },
                new Attribute { Name = "game_account_name", Value = new Variant { StringValue = "x" } },
                new Attribute { Name = "account_mail", Value = new Variant { StringValue = "x" } },
            ],
        };

        Assert.Throws<InvalidOperationException>(() => GameUtilitiesSession.ParseHandoff(response));
    }
}

public class LogonBuilderTests
{
    [Fact]
    public void BuildLogonRequest_SetsFixedFieldsPerDocumentedClientBuild()
    {
        var request = LogonBuilder.BuildLogonRequest();
        var decoded = LogonRequest.Decode(request.Encode());

        Assert.Equal("S2", decoded.Program);
        Assert.Equal("Mc64", decoded.Platform);
        Assert.Equal("enUS", decoded.Locale);
        Assert.Equal(LogonBuilder.EmbeddedSdkVersion, decoded.Version);
        Assert.Equal(0, decoded.ApplicationVersion);
        Assert.True(decoded.AllowLogonQueueNotifications);
        Assert.Null(decoded.CachedWebCredentials);
    }

    [Fact]
    public void BuildLogonRequest_WithCachedCredentials_RoundTrips()
    {
        byte[] cached = [1, 2, 3, 4];

        var request = LogonBuilder.BuildLogonRequest(cached);
        var decoded = LogonRequest.Decode(request.Encode());

        Assert.Equal(cached, decoded.CachedWebCredentials);
    }
}
