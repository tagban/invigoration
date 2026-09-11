using Invigoration.Core.Chat;

namespace Invigoration.Core.Tests;

public class ChatIconTests
{
    [Theory]
    [InlineData("3RAW", "war3")]
    [InlineData("PX3W", "war3")]
    [InlineData("PX2D", "d2exp")]
    [InlineData("VD2D", "diablo2")]
    [InlineData("LTRD", "diablo")]
    [InlineData("RHSD", "dshr")]
    [InlineData("RATS", "sc")]
    [InlineData("PXES", "scbw")]
    [InlineData("RHSS", "sware")]
    [InlineData("RTSJ", "jsc")]
    [InlineData("NB2W", "war2")]
    // Regression: StatStringParser already recognizes "TAHC" as a Chat bot's product code, but
    // GetProductIconKey had no matching case, so a user connecting with that product silently
    // got no icon at all (the "chat" icon asset/catalog entry existed but was unreachable).
    [InlineData("TAHC", "chat")]
    [InlineData("sc2", "sc2")]
    public void GetProductIconKey_KnownProduct_ReturnsExpectedKey(string statString, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetProductIconKey(statString));
    }

    [Fact]
    public void GetProductIconKey_UnknownProduct_ReturnsEmpty()
    {
        Assert.Equal("", ChatIcon.GetProductIconKey("ZZZZ"));
    }

    // "LTRD" + 1 icon byte, then "<level> <dots> <class> <str> <magic> <dex> <vit> <gold>
    // <unused>" glued on with no separating space (statString[5..] is the first value's own first
    // digit, not a delimiter) — same wire layout ParseDiabloClassicStats reads. dots (index 1) is
    // classic.battle.net/info/icons.shtml's "red dots" progress marker (0-3, no/Normal/Nightmare/
    // Hell killed).
    [Theory]
    [InlineData("LTRDX30 0 0 55 20 15 40 100 0", "diablo-dot0")]
    [InlineData("LTRDX30 1 0 55 20 15 40 100 0", "diablo-dot1")]
    [InlineData("LTRDX30 2 0 55 20 15 40 100 0", "diablo-dot2")]
    [InlineData("LTRDX30 3 0 55 20 15 40 100 0", "diablo-dot3")]
    public void GetProductIconKey_DiabloWithDots_ReturnsDotBadge(string statString, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetProductIconKey(statString));
    }

    // Regression: an Open Character (no icon-class stats parsed yet, statString.Length <= 5) has
    // no dots to read — GetProductIconKey must fall back to the flat "diablo" icon rather than
    // throwing or defaulting to dot0 by accident.
    [Fact]
    public void GetProductIconKey_DiabloOpenCharacter_FallsBackToPlainIcon()
    {
        Assert.Equal("diablo", ChatIcon.GetProductIconKey("LTRD"));
    }

    // Regression: a status icon (e.g. a Blizzard rep playing Diablo) must keep showing its status
    // badge as the more important thing — the dot badge should only ever replace the plain
    // product icon, matching the explicit request to show it "when they don't have a special
    // icon."
    [Fact]
    public void GetProductIconKey_DiabloWithStatusIcon_StatusTakesPriorityOverDots()
    {
        var statString = "LTRDX30 3 0 55 20 15 40 100 0";
        Assert.Equal("diablo", ChatIcon.GetProductIconKey(statString, (uint)UserFlags.Blizzard));
    }

    [Theory]
    [InlineData(UserFlags.Blizzard, "blizz")]
    [InlineData(UserFlags.Admin, "sysop")]
    [InlineData(UserFlags.Operator, "mod-gavel")]
    [InlineData(UserFlags.Speaker, "mega")]
    // Regression: UserFlags.Special (BLIZZARD_GUEST, "glasses" icon in the original) was defined
    // but GetStatusIconKey had no case for it, so a special-guest user never got a status badge at
    // all — ChatPalette already treated this same flag as "Guest" for text color, but the icon
    // half was missing.
    [InlineData(UserFlags.Special, "guest")]
    [InlineData(UserFlags.Squelched, "ignore")]
    [InlineData(UserFlags.None, "")]
    public void GetStatusIconKey_KnownFlag_ReturnsExpectedKey(UserFlags flag, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetStatusIconKey((uint)flag));
    }

    [Fact]
    public void GetStatusIconKey_PrecedenceOrder_BlizzardBeatsEverythingElse()
    {
        var flags = UserFlags.Blizzard | UserFlags.Admin | UserFlags.Operator | UserFlags.Speaker |
                    UserFlags.Special | UserFlags.Squelched;
        Assert.Equal("blizz", ChatIcon.GetStatusIconKey((uint)flags));
    }

    [Fact]
    public void GetStatusIconKey_SpeakerAndSpecial_SpeakerTakesPrecedence()
    {
        var flags = UserFlags.Speaker | UserFlags.Special;
        Assert.Equal("mega", ChatIcon.GetStatusIconKey((uint)flags));
    }

    [Fact]
    public void IsPrivileged_SpecialGuestAlone_IsNotPrivileged()
    {
        Assert.False(ChatIcon.IsPrivileged((uint)UserFlags.Special));
    }

    // "RATSX" + "<rating> <?> <wins> <spawn> ..." (9 values total) — same wire layout
    // ParseIconClassStats reads (rating = values[0], wins = values[2]) for StarCraft/Brood War/
    // SC-Japanese. Real icons_STAR.bni sprites (bnetdocs.org/document/25/icons-bni) give half a
    // star per win up to 5 stars at 10+ wins — confirmed against the user's own firsthand
    // knowledge of the real thresholds, not guessed.
    [Theory]
    [InlineData("RATSX0 0 0 0 0 0 0 0 0", "sc-stars0")]
    [InlineData("RATSX0 0 1 0 0 0 0 0 0", "sc-stars1")]
    [InlineData("RATSX0 0 5 0 0 0 0 0 0", "sc-stars5")]
    [InlineData("RATSX0 0 10 0 0 0 0 0 0", "sc-stars10")]
    // Regression: wins can exceed the 5-star display cap (10) while still unranked — must clamp
    // rather than index past the real sprite set's last entry.
    [InlineData("RATSX0 0 27 0 0 0 0 0 0", "sc-stars10")]
    public void GetProductIconKey_StarCraftUnranked_ReturnsWinStars(string statString, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetProductIconKey(statString));
    }

    [Theory]
    [InlineData("PXES")]
    [InlineData("RTSJ")]
    public void GetProductIconKey_BroodWarAndJapaneseShareTheSameStarBadge(string productCode)
    {
        Assert.Equal("sc-stars3", ChatIcon.GetProductIconKey($"{productCode}X0 0 3 0 0 0 0 0 0"));
    }

    // Regression: once a player is actually ladder-rated (rating > 0), the real client drops the
    // win-count star display entirely in favor of stamping the rating as text on the plain plate
    // (see GetLadderScore) — confirmed by the user, not assumed. A high win count must not still
    // show a 5-star badge once ranked.
    [Fact]
    public void GetProductIconKey_StarCraftRanked_ReturnsPlainPlateRegardlessOfWins()
    {
        Assert.Equal("sc-stars0", ChatIcon.GetProductIconKey("RATSX1500 0 9 0 0 0 0 0 0"));
    }

    // "NB2W" family: half an axe per win up to 4 axes (8 wins), a single sword at 9, two swords
    // from 10 until ladder-rated — exact thresholds from the user, not the "matches StarCraft"
    // guess this replaced.
    [Theory]
    [InlineData("NB2WX0 0 0 0 0 0 0 0 0", "war2-axes0")]
    [InlineData("NB2WX0 0 4 0 0 0 0 0 0", "war2-axes4")]
    [InlineData("NB2WX0 0 8 0 0 0 0 0 0", "war2-axes8")]
    [InlineData("NB2WX0 0 9 0 0 0 0 0 0", "war2-sword1")]
    [InlineData("NB2WX0 0 10 0 0 0 0 0 0", "war2-sword2")]
    [InlineData("NB2WX0 0 40 0 0 0 0 0 0", "war2-sword2")]
    public void GetProductIconKey_Warcraft2Unranked_ReturnsAxeOrSwordBadge(string statString, string expected)
    {
        Assert.Equal(expected, ChatIcon.GetProductIconKey(statString));
    }

    [Fact]
    public void GetProductIconKey_Warcraft2Ranked_ReturnsRankedPlateRegardlessOfWins()
    {
        Assert.Equal("war2-ranked", ChatIcon.GetProductIconKey("NB2WX1200 0 3 0 0 0 0 0 0"));
    }

    // Regression: a status icon must keep taking priority over the win/rank badge for StarCraft/
    // Warcraft II too, same rule already covered for Diablo above.
    [Fact]
    public void GetProductIconKey_StarCraftWithStatusIcon_StatusTakesPriorityOverStars()
    {
        var statString = "RATSX1500 0 9 0 0 0 0 0 0";
        Assert.Equal("sc", ChatIcon.GetProductIconKey(statString, (uint)UserFlags.Blizzard));
    }

    [Theory]
    [InlineData("RATSX0 0 0 0 0 0 0 0 0", null)]
    [InlineData("RATSX1500 0 9 0 0 0 0 0 0", 1500)]
    [InlineData("NB2WX850 0 3 0 0 0 0 0 0", 850)]
    // Regression: Diablo's rank badge (the dot picker) has no numeric score equivalent to stamp.
    [InlineData("LTRDX30 2 0 55 20 15 40 100 0", null)]
    [InlineData("war3", null)]
    public void GetLadderScore_ReturnsRatingOnlyForRankedIconClassProducts(string statString, int? expected)
    {
        Assert.Equal(expected, ChatIcon.GetLadderScore(statString));
    }
}
