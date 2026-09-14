using Invigoration.Core.Auth;
using Invigoration.Core.Protocol;

namespace Invigoration.Core.Tests;

public class BattlenetPasswordTests
{
    [Theory]
    [InlineData(BncsProduct.Starcraft)]
    [InlineData(BncsProduct.StarcraftBroodWar)]
    [InlineData(BncsProduct.StarcraftJapanese)]
    [InlineData(BncsProduct.Warcraft2BNE)]
    [InlineData(BncsProduct.Diablo)]
    [InlineData(BncsProduct.DiabloII)]
    [InlineData(BncsProduct.DiabloIILoD)]
    public void Normalize_LowercasesForClassicProducts(string product)
    {
        Assert.Equal("huntertwo!", BattlenetPassword.Normalize("HunterTWO!", product));
    }

    [Theory]
    [InlineData(BncsProduct.Warcraft3)]
    [InlineData(BncsProduct.Warcraft3TFT)]
    public void Normalize_UppercasesForWarcraftIII(string product)
    {
        Assert.Equal("HUNTERTWO!", BattlenetPassword.Normalize("HunterTWO!", product));
    }

    [Theory]
    [InlineData("HunterTWO", BncsProduct.Starcraft, "huntertwo")]
    [InlineData("HunterTWO", BncsProduct.Warcraft3TFT, "HUNTERTWO")]
    public void RetryCasing_OffersTheGameClientForm(string typed, string product, string expected)
    {
        Assert.Equal(expected, BattlenetPassword.RetryCasing(typed, product));
    }

    // No retry when it would send the identical string again — it could only fail the same way.
    [Theory]
    [InlineData("huntertwo", BncsProduct.DiabloII)]
    [InlineData("HUNTERTWO", BncsProduct.Warcraft3)]
    [InlineData("12345!", BncsProduct.Starcraft)]
    public void RetryCasing_IsNullWhenAlreadyInThatForm(string typed, string product)
    {
        Assert.Null(BattlenetPassword.RetryCasing(typed, product));
    }

    // The NLS logon's BNLS_LOGONCHALLENGE carries the normalized string, so changing case must
    // never change an ASCII password's length.
    [Theory]
    [InlineData(BncsProduct.Starcraft)]
    [InlineData(BncsProduct.Warcraft3TFT)]
    public void Normalize_KeepsLengthForAsciiPasswords(string product)
    {
        const string typed = "MiXeD-CaSe_Pass99";
        Assert.Equal(typed.Length, BattlenetPassword.Normalize(typed, product).Length);
    }

    // Reference values from BNLS itself (bnls.bnetdocs.org, BNLS_HASHDATA, made-up passwords): the
    // local hashes must be exactly what the logon used to get from BNLS.
    [Theory]
    [InlineData("password", "ecc80d1d76e758c0b9da8c25ff106aff8e242916")]
    [InlineData("HunterTWO", "23fa529f188801de7a0458283abc8a1dbd143a93")]
    [InlineData("huntertwo", "8d761837043274eea68ed4a26856fc99b8a0c599")]
    [InlineData("aaaaaaaaaaaaaaaaaaaa", "2aad7effe592e9144e79817a53fcde2ca3aabf71")]
    public void Hash_MatchesBnls(string password, string expected)
    {
        Assert.Equal(expected, Convert.ToHexStringLower(BattlenetPassword.Hash(password)));
    }

    [Fact]
    public void Proof_MatchesBnlsDoubleHash()
    {
        Assert.Equal("3df0cefa10382b8ab02e9ad6f28fe4b93f2b2f78",
            Convert.ToHexStringLower(BattlenetPassword.Proof(0x12345678, 0x9ABCDEF0, "huntertwo")));
    }

    [Fact]
    public void Hash_KeepsTheCaseItsGiven()
    {
        Assert.NotEqual(BattlenetPassword.Hash("HunterTWO"), BattlenetPassword.Hash("huntertwo"));
    }

    [Theory]
    [InlineData("HunterTWO", BncsProduct.Starcraft, false, "huntertwo", "HunterTWO")]
    [InlineData("HunterTWO", BncsProduct.Starcraft, true, "HunterTWO", "huntertwo")]
    [InlineData("HunterTWO", BncsProduct.Warcraft3TFT, false, "HUNTERTWO", "HunterTWO")]
    [InlineData("huntertwo", BncsProduct.Warcraft2BNE, false, "huntertwo", null)]
    [InlineData("huntertwo", BncsProduct.Warcraft2BNE, true, "huntertwo", null)]
    public void FirstAttemptAndOtherCasing(string typed, string product, bool sentAsTyped, string first, string? other)
    {
        Assert.Equal(first, BattlenetPassword.FirstAttempt(typed, product, sentAsTyped));
        Assert.Equal(other, BattlenetPassword.OtherCasing(typed, product, sentAsTyped));
    }
}
