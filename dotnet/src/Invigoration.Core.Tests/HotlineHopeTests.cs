using System.Buffers.Binary;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using Invigoration.Core.Hotline;

namespace Invigoration.Core.Tests;

/// <summary>
/// HOPE's secure login. The whole point is that the password never crosses the wire, so the MACs
/// are checked against independently-computed values rather than against themselves — a MAC that
/// hashes the wrong way round still looks like a MAC, and only a real server would notice.
/// </summary>
public class HotlineHopeTests
{
    private static byte[] SessionKey(string address = "203.0.113.7", ushort port = 5500)
    {
        var key = new byte[64];
        IPAddress.Parse(address).GetAddressBytes().CopyTo(key, 0);
        BinaryPrimitives.WriteUInt16BigEndian(key.AsSpan(4), port);
        for (var i = 6; i < key.Length; i++)
        {
            key[i] = (byte)i;
        }

        return key;
    }

    [Fact]
    public void Mac_KeyedAlgorithms_MacTheSessionKeyUnderThePassword()
    {
        var key = SessionKey();
        var password = "hunter2";
        var passwordBytes = Encoding.UTF8.GetBytes(password);

        Assert.Equal(HMACSHA256.HashData(passwordBytes, key), HotlineHope.Mac("HMAC-SHA256", password, key));
        Assert.Equal(HMACSHA1.HashData(passwordBytes, key), HotlineHope.Mac("HMAC-SHA1", password, key));
        Assert.Equal(HMACMD5.HashData(passwordBytes, key), HotlineHope.Mac("HMAC-MD5", password, key));
    }

    /// <summary>The bare digests hash password-then-key concatenated, which is a different value from the keyed forms.</summary>
    [Fact]
    public void Mac_BareDigests_HashTheConcatenation()
    {
        var key = SessionKey();
        var combined = Encoding.UTF8.GetBytes("hunter2").Concat(key).ToArray();

        Assert.Equal(SHA1.HashData(combined), HotlineHope.Mac("SHA1", "hunter2", key));
        Assert.Equal(MD5.HashData(combined), HotlineHope.Mac("MD5", "hunter2", key));
        Assert.NotEqual(HotlineHope.Mac("SHA1", "hunter2", key), HotlineHope.Mac("HMAC-SHA1", "hunter2", key));
    }

    /// <summary>INVERSE is the classic scheme, kept as the fallback both sides must understand.</summary>
    [Fact]
    public void Mac_Inverse_IsTheClassicObfuscation() =>
        Assert.Equal(HotlineTransactionClient.XorObfuscate("hunter2"), HotlineHope.Mac("INVERSE", "hunter2", SessionKey()));

    [Fact]
    public void Mac_UnknownAlgorithm_IsNull() => Assert.Null(HotlineHope.Mac("HMAC-SHA3-512", "hunter2", SessionKey()));

    /// <summary>The password must never appear in what goes on the wire — the entire reason HOPE exists.</summary>
    [Fact]
    public void AnAuthenticatedLogin_NeverCarriesThePasswordItself()
    {
        var identification = new HotlineHope.ServerIdentification(SessionKey(), "HMAC-SHA256", LoginIsMacd: false, "HotSocket 3.0");

        var fields = HotlineHope.AuthenticatedLoginFields(identification, "tagban", "hunter2", "Tagban", 414, 6112)!;

        var everything = fields.SelectMany(f => f.Data).ToArray();
        Assert.True(everything.AsSpan().IndexOf("hunter2"u8) < 0, "the plaintext password appeared in the login fields");

        var passwordField = fields.Single(f => f.Type == (ushort)HotlineFieldType.UserPassword);
        Assert.Equal(HMACSHA256.HashData("hunter2"u8.ToArray(), SessionKey()), passwordField.Data);
    }

    /// <summary>When the server's login field is blank, the login name goes the classic way while the password is still MAC'd.</summary>
    [Fact]
    public void WhenTheServerDoesntAskForAMacdLogin_TheLoginIsSentClassically()
    {
        var identification = new HotlineHope.ServerIdentification(SessionKey(), "HMAC-SHA256", LoginIsMacd: false, null);

        var fields = HotlineHope.AuthenticatedLoginFields(identification, "tagban", "hunter2", "Tagban", 414, null)!;
        var loginField = fields.Single(f => f.Type == (ushort)HotlineFieldType.UserLogin);

        Assert.Equal(HotlineTransactionClient.XorObfuscate("tagban"), loginField.Data);
    }

    [Fact]
    public void WhenTheServerAsksForAMacdLogin_TheLoginIsMacdToo()
    {
        var identification = new HotlineHope.ServerIdentification(SessionKey(), "HMAC-SHA256", LoginIsMacd: true, null);

        var fields = HotlineHope.AuthenticatedLoginFields(identification, "tagban", "hunter2", "Tagban", 414, null)!;
        var loginField = fields.Single(f => f.Type == (ushort)HotlineFieldType.UserLogin);

        Assert.Equal(HMACSHA256.HashData("tagban"u8.ToArray(), SessionKey()), loginField.Data);
    }

    /// <summary>An anonymous login sends an empty password field, not a MAC of the empty string.</summary>
    [Fact]
    public void AnEmptyPassword_IsSentEmpty()
    {
        var identification = new HotlineHope.ServerIdentification(SessionKey(), "HMAC-SHA256", LoginIsMacd: false, null);

        var fields = HotlineHope.AuthenticatedLoginFields(identification, "guest", "", "Guest", 414, null)!;

        Assert.Empty(fields.Single(f => f.Type == (ushort)HotlineFieldType.UserPassword).Data);
    }

    /// <summary>An algorithm this client can't compute must produce nothing at all, so the caller falls back rather than sending a meaningless value.</summary>
    [Fact]
    public void AnUncomputableAlgorithm_ProducesNoFields()
    {
        var identification = new HotlineHope.ServerIdentification(SessionKey(), "HMAC-WHIRLPOOL", LoginIsMacd: false, null);

        Assert.Null(HotlineHope.AuthenticatedLoginFields(identification, "tagban", "hunter2", "Tagban", 414, null));
    }

    // --- The identification exchange ---

    /// <summary>The null login is the signal that asks a server whether it speaks HOPE.</summary>
    [Fact]
    public void IdentificationFields_AskWithANullLoginAndSayWhatThisClientIs()
    {
        var fields = HotlineHope.IdentificationFields();

        Assert.Equal(new byte[] { 0x00 }, fields.Single(f => f.Type == (ushort)HotlineFieldType.UserLogin).Data);
        Assert.Equal("INVG", Encoding.ASCII.GetString(fields.Single(f => f.Type == (ushort)HotlineFieldType.HopeAppId).Data));
        Assert.Contains("Invigoration", fields.Single(f => f.Type == (ushort)HotlineFieldType.HopeAppString).AsString());

        var offered = HotlineHope.DecodeNameList(fields.Single(f => f.Type == (ushort)HotlineFieldType.HopeMacAlgorithm).Data);
        Assert.Equal("HMAC-SHA256", offered[0]);
        Assert.Equal("INVERSE", offered[^1]);
    }

    [Fact]
    public void NameList_RoundTrips()
    {
        string[] names = ["HMAC-SHA256", "HMAC-SHA1", "INVERSE"];

        Assert.Equal(names, HotlineHope.DecodeNameList(HotlineHope.EncodeNameList(names)));
    }

    /// <summary>A server replying with one bare algorithm name — no count prefix — still has to be understood.</summary>
    [Fact]
    public void NameList_ABareNameWithNoCount_IsReadAsOneName()
    {
        var decoded = HotlineHope.DecodeNameList("INVERSE"u8.ToArray());

        Assert.Equal(["INVERSE"], decoded);
    }

    private static HotlineTransactionFrame Reply(params HotlineField[] fields) =>
        HotlineTransactionFrame.CreateReply(1, 0, fields);

    [Fact]
    public void ReadServerIdentification_ReadsTheChallengeAndChoice()
    {
        var reply = Reply(
            new HotlineField(HotlineFieldType.HopeSessionKey, SessionKey()),
            new HotlineField(HotlineFieldType.HopeMacAlgorithm, HotlineHope.EncodeNameList(["HMAC-SHA256"])),
            new HotlineField(HotlineFieldType.HopeAppString, "HotSocket 3.0"));

        var identification = HotlineHope.ReadServerIdentification(reply);

        Assert.NotNull(identification);
        Assert.Equal("HMAC-SHA256", identification.MacAlgorithm);
        Assert.Equal("HotSocket 3.0", identification.ServerApp);
        Assert.False(identification.LoginIsMacd);
    }

    /// <summary>A server that doesn't speak HOPE answers the null login like any other failed one — no session key, so nothing to negotiate.</summary>
    [Fact]
    public void ReadServerIdentification_AClassicServersReply_IsNotHope()
    {
        Assert.Null(HotlineHope.ReadServerIdentification(Reply(new HotlineField(HotlineFieldType.ErrorText, "Bad login"))));
        Assert.Null(HotlineHope.ReadServerIdentification(null));
    }

    /// <summary>The session key carries the address the server believes it has, so a client can notice something in the middle.</summary>
    [Fact]
    public void TheSessionKey_CarriesTheServersOwnAddress()
    {
        var identification = HotlineHope.ReadServerIdentification(Reply(
            new HotlineField(HotlineFieldType.HopeSessionKey, SessionKey("198.51.100.4", 5501))))!;

        var embedded = identification.EmbeddedEndpoint;

        Assert.Equal(IPAddress.Parse("198.51.100.4"), embedded?.Address);
        Assert.Equal(5501, embedded?.Port);
    }
}
