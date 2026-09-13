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
    public void Normalize_LeavesWarcraftIIIPasswordsAsTyped(string product)
    {
        Assert.Equal("HunterTWO!", BattlenetPassword.Normalize("HunterTWO!", product));
    }

    // SendPasswordHashRequestAsync writes the length field from the normalized string, so
    // lowercasing must never change an ASCII password's length.
    [Fact]
    public void Normalize_KeepsLengthForAsciiPasswords()
    {
        const string typed = "MiXeD-CaSe_Pass99";
        Assert.Equal(typed.Length, BattlenetPassword.Normalize(typed, BncsProduct.Starcraft).Length);
    }
}
