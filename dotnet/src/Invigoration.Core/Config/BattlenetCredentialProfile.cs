namespace Invigoration.Core.Config;

/// <summary>
/// A named, user-managed Battle.net login — lets one or more bots (SC2 today;
/// StarCraft: Remastered / WarCraft III: Reforged once those exist, since the
/// same Battle.net account works across all of them) share one signed-in
/// session, or use deliberately separate ones. See BattlenetCredentialProfileStore.
/// </summary>
public sealed class BattlenetCredentialProfile
{
    /// <summary>Stable identity, assigned once at creation and never changed by a rename — this, not Name, is what a bot's config references and what the cached credential file is named after.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>User-editable label — free-text, never auto-overwritten (e.g. "Main", "Smurf"), so a deliberate rename always survives. Defaults to the first bot that auto-created this profile's DisplayName, purely as a starting point.</summary>
    public string Name { get; set; } = "New Profile";

    /// <summary>
    /// The actual signed-in Battle.net username (e.g. "Player#1234"), captured from Stimpak's
    /// AccountConnected event the first time this profile successfully connects — see
    /// BattlenetCredentialProfileStore.UpdateBattleTag. Null until that's happened at least
    /// once. This, not the free-text Name, is what actually tells two profiles for two
    /// different real accounts apart — see DisplayLabel.
    /// </summary>
    public string? BattleTag { get; set; }

    /// <summary>What the UI should show to identify this profile — the real signed-in username once known (the whole point of a profile is to track which actual account it is), falling back to the free-text Name for a profile that's never signed in yet.</summary>
    public string DisplayLabel => string.IsNullOrEmpty(BattleTag) ? Name : BattleTag;
}
