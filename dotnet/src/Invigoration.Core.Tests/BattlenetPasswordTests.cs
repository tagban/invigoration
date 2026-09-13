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

    // SendPasswordHashRequestAsync writes the length field from the normalized string, so
    // changing case must never change an ASCII password's length.
    [Theory]
    [InlineData(BncsProduct.Starcraft)]
    [InlineData(BncsProduct.Warcraft3TFT)]
    public void Normalize_KeepsLengthForAsciiPasswords(string product)
    {
        const string typed = "MiXeD-CaSe_Pass99";
        Assert.Equal(typed.Length, BattlenetPassword.Normalize(typed, product).Length);
    }
}
