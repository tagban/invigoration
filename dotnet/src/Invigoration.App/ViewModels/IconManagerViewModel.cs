using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Config;
using System.Diagnostics;

namespace Invigoration.App.ViewModels;

public partial class IconSlotViewModel(string key, string displayName) : ObservableObject
{
    public string Key { get; } = key;

    public string DisplayName { get; } = displayName;

    [ObservableProperty]
    public partial Bitmap? PreviewImage { get; set; }

    [ObservableProperty]
    public partial bool HasOverride { get; set; }

    public void Refresh()
    {
        HasOverride = IconOverrideStore.GetOverridePath(Key) is not null;
        PreviewImage = GameIconLoader.Get(Key);
    }
}

/// <summary>Lets a user replace any bundled chat icon with their own image, stored via <see cref="IconOverrideStore"/>.</summary>
public partial class IconManagerViewModel : ViewModelBase
{
    /// <summary>
    /// Keys with a bundled 64x64 alternate available under Assets/GameIconsHD — the classic
    /// Battle.net "bnet-*" chat icon set (originally served from WC3 ladder pages); war3/w3tft
    /// specifically are Bnet-war3.png/Bnet-war3x.png pulled from warcraft.wiki.gg's chat-icon
    /// reference page (64x42, same source family as the rest of this set).
    /// </summary>
    private static readonly string[] HighResolutionKeys =
        ["blizz", "sysop", "mega", "ignore", "chat", "diablo", "diablo2", "d2exp", "sc", "scbw", "war2", "war3", "w3tft"];

    /// <summary>
    /// Keys with no distinct HD chat icon of their own upstream — rather than leaving these at
    /// the default 28x14 art (the only ones "Apply Bundled 64x64 Set" would otherwise skip),
    /// each borrows its closest relative's HD asset: same underlying game (StarCraft) for the
    /// Japanese release and the shareware trial, same idea for Diablo's shareware trial.
    /// </summary>
    private static readonly (string Key, string FallbackFrom)[] HighResolutionFallbacks =
    [
        ("jsc", "sc"),
        ("sware", "sc"),
        ("dshr", "diablo"),
    ];

    /// <summary>
    /// The optional modern-Battle.net alternate for every classic key that has a real official
    /// account.battle.net game icon — a separate opt-in set from the default 28x14 classic art,
    /// same idea as HighResolutionKeys/GameIconsHD but sourced from account.battle.net's own SVGs
    /// (rasterized under Assets/GameIconsBnet2, since nothing here renders SVG directly). Several
    /// keys intentionally share one source image, matching how Blizzard's own modern branding
    /// doesn't distinguish them: StarCraft/Brood War/the Japanese release/the shareware trial all
    /// point at one "StarCraft: Remastered" icon, and the same is true for Warcraft III/TFT and
    /// Diablo II/Lord of Destruction. sc2 and the new Bnet2Icons keys aren't here at all — they
    /// have no classic default worth preserving as an opt-in, so they use these images as their
    /// one and only default already (see Assets/GameIcons/sc2.png etc.)
    /// </summary>
    private static readonly (string Key, string SourceAssetKey)[] Bnet2ModernKeys =
    [
        ("diablo2", "diablo-ii"),
        ("d2exp", "diablo-ii"),
        ("war3", "warcraft-iii"),
        ("w3tft", "warcraft-iii"),
        ("war2", "warcraft-ii-remastered"),
        ("sc", "starcraft-remastered"),
        ("scbw", "starcraft-remastered"),
        ("jsc", "starcraft-remastered"),
        ("sware", "starcraft-remastered"),
        ("diablo", "diablo"),
        ("dshr", "diablo"),
    ];

    public ObservableCollection<IconSlotViewModel> GameIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> StatusIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> FriendIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> CustomIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> Bnet2Icons { get; } = [];

    /// <summary>Names of user-saved icon sets, each an ordinary folder under IconSetStore.Directory (inside the app's config folder, so it travels with any config backup).</summary>
    public ObservableCollection<string> SavedSets { get; } = [];

    [ObservableProperty]
    public partial string? SelectedSet { get; set; }

    [ObservableProperty]
    public partial string NewSetName { get; set; } = "";

    public IconManagerViewModel()
    {
        foreach (var (key, displayName) in IconCatalog.GameIcons)
        {
            GameIcons.Add(CreateSlot(key, displayName));
        }

        foreach (var (key, displayName) in IconCatalog.StatusIcons)
        {
            StatusIcons.Add(CreateSlot(key, displayName));
        }

        foreach (var (key, displayName) in IconCatalog.FriendIcons)
        {
            FriendIcons.Add(CreateSlot(key, displayName));
        }

        foreach (var (key, displayName) in IconCatalog.CustomIcons)
        {
            CustomIcons.Add(CreateSlot(key, displayName));
        }

        foreach (var (key, displayName) in IconCatalog.Bnet2Icons)
        {
            Bnet2Icons.Add(CreateSlot(key, displayName));
        }

        RefreshSavedSets();
    }

    /// <summary>Applies a picked file as the override for a slot. Public (not a RelayCommand) since the file picker itself has to run from code-behind — see IconManagerWindow.axaml.cs.</summary>
    public void ApplyIcon(IconSlotViewModel slot, string sourceFilePath)
    {
        IconOverrideStore.SetOverride(slot.Key, sourceFilePath);
        slot.Refresh();
    }

    [RelayCommand]
    private void ResetIcon(IconSlotViewModel slot)
    {
        IconOverrideStore.ClearOverride(slot.Key);
        slot.Refresh();
    }

    /// <summary>Applies the bundled 64x64 icon set to every key that has one (plus HighResolutionFallbacks for the handful that don't) — a concrete, one-click example of swapping in a larger icon set, using real Blizzard-hosted assets rather than a hypothetical.</summary>
    [RelayCommand]
    private void ApplyHighResolutionSet()
    {
        foreach (var key in HighResolutionKeys)
        {
            ApplyHighResolutionAsset(targetKey: key, sourceAssetKey: key);
        }

        foreach (var (key, fallbackFrom) in HighResolutionFallbacks)
        {
            ApplyHighResolutionAsset(targetKey: key, sourceAssetKey: fallbackFrom);
        }

        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons).Concat(Bnet2Icons))
        {
            slot.Refresh();
        }
    }

    private static void ApplyHighResolutionAsset(string targetKey, string sourceAssetKey)
    {
        var uri = new Uri($"avares://Invigoration.App/Assets/GameIconsHD/{sourceAssetKey}.gif");
        using var stream = AssetLoader.Open(uri);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        IconOverrideStore.SetOverrideBytes(targetKey, buffer.ToArray(), ".gif");
    }

    /// <summary>Applies the official account.battle.net modern icon set to every classic key that has one — see Bnet2ModernKeys.</summary>
    [RelayCommand]
    private void ApplyBnet2IconSet()
    {
        foreach (var (key, sourceAssetKey) in Bnet2ModernKeys)
        {
            var uri = new Uri($"avares://Invigoration.App/Assets/GameIconsBnet2/{sourceAssetKey}.png");
            using var stream = AssetLoader.Open(uri);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            IconOverrideStore.SetOverrideBytes(key, buffer.ToArray(), ".png");
        }

        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons).Concat(Bnet2Icons))
        {
            slot.Refresh();
        }
    }

    /// <summary>Clears every override, reverting all icons to the bundled classic 28x14 defaults in one action.</summary>
    [RelayCommand]
    private void ResetAllIcons()
    {
        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons).Concat(Bnet2Icons))
        {
            IconOverrideStore.ClearOverride(slot.Key);
            slot.Refresh();
        }
    }

    /// <summary>Snapshots the current set of overrides (whatever mix of custom/HD/default icons is active) under a name, so it can be swapped back to later or backed up as a folder.</summary>
    [RelayCommand]
    private void SaveCurrentAsSet()
    {
        if (string.IsNullOrWhiteSpace(NewSetName))
        {
            return;
        }

        IconSetStore.SaveCurrentAsSet(NewSetName);
        NewSetName = "";
        RefreshSavedSets();
    }

    [RelayCommand]
    private void ApplySelectedSet()
    {
        if (SelectedSet is null)
        {
            return;
        }

        IconSetStore.ApplySet(SelectedSet);
        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons).Concat(Bnet2Icons))
        {
            slot.Refresh();
        }
    }

    [RelayCommand]
    private void DeleteSelectedSet()
    {
        if (SelectedSet is null)
        {
            return;
        }

        IconSetStore.DeleteSet(SelectedSet);
        SelectedSet = null;
        RefreshSavedSets();
    }

    /// <summary>Opens the icon sets folder in the OS file explorer — each set is just a plain folder, so this is the whole "backup" story.</summary>
    [RelayCommand]
    private void OpenIconSetsFolder()
    {
        Directory.CreateDirectory(IconSetStore.Directory);
        Process.Start(new ProcessStartInfo(IconSetStore.Directory) { UseShellExecute = true });
    }

    private void RefreshSavedSets()
    {
        var selected = SelectedSet;
        SavedSets.Clear();
        foreach (var name in IconSetStore.ListSets())
        {
            SavedSets.Add(name);
        }

        SelectedSet = selected is not null && SavedSets.Contains(selected) ? selected : null;
    }

    private static IconSlotViewModel CreateSlot(string key, string displayName)
    {
        var slot = new IconSlotViewModel(key, displayName);
        slot.Refresh();
        return slot;
    }
}
