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

    public string Name { get; set; } = "New Profile";
}
