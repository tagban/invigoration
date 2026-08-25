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
    public const string Bnet1ClassicSetName = "Battle.net 1.0 Classic";
    public const string Wc3ClassicSetName = "Warcraft III Classic";
    public const string Bnet2SetName = "Battle.net 2.0";

    /// <summary>The three bundled (not user-saved) icon sets offered in the same apply dropdown as SavedSets — see AvailableIconSets.</summary>
    private static readonly string[] BundledSetNames = [Bnet1ClassicSetName, Wc3ClassicSetName, Bnet2SetName];

    /// <summary>
    /// "Battle.net 1.0 Classic" — the original classic.battle.net chat-icon set
    /// (classic.battle.net/info/icons.shtml), 28x14, sourced 2026-08-24. This is also what
    /// GameIconLoader's own bundled defaults are for every key here except blizz/sysop/mod-gavel/
    /// mega/ignore (which default to the sharper Warcraft III Classic art instead — see
    /// GameIconLoader's remarks) — selecting this set explicitly is how to get the small original
    /// look back for those five specifically. scbw ("W2sexp.png", the real StarCraft: Brood War
    /// badge — classic.battle.net's own /info/icons.shtml never actually distinguished it from
    /// plain StarCraft, a real bug fixed 2026-08-24 alongside Chat.ChatIcon.GetProductIconKey's
    /// matching PXES-mapping fix) and w3tft ("W2w3xp.png") both come from
    /// warcraft.wiki.gg/wiki/Warcraft_II_chat_icons instead — confirmed classic.battle.net
    /// itself never shipped one at all (its own site "never fully updated before it went down,"
    /// per the user); this wiki apparently caught a copy before that happened, filed oddly under
    /// the Warcraft II chat-icons page rather than Warcraft III's own. sc2 has no official
    /// classic-era icon at all (StarCraft II postdates this whole art style) — this one is
    /// user-made (2026-08-24, hand-drawn to match the aesthetic, dropped straight in at the same
    /// 28x14 size as the rest of this set) rather than sourced from anywhere. Not in Wc3ClassicSet
    /// below: that set is 64x64, and this icon is 28x14 — forcing it in would look inconsistently
    /// tiny/blurry next to the rest of that set.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Bnet1ClassicSet =
    [
        ("sc", "GameIconsClassic", "sc"), ("scbw", "GameIconsClassic", "scbw"), ("jsc", "GameIconsClassic", "jsc"),
        ("sware", "GameIconsClassic", "sware"), ("war2", "GameIconsClassic", "war2"),
        ("war3", "GameIconsClassic", "war3"), ("w3tft", "GameIconsClassic", "w3tft"),
        ("diablo", "GameIconsClassic", "diablo"), ("dshr", "GameIconsClassic", "dshr"),
        ("diablo2", "GameIconsClassic", "diablo2"), ("d2exp", "GameIconsClassic", "d2exp"),
        ("chat", "GameIconsClassic", "chat"), ("blizz", "GameIconsClassic", "blizz"),
        ("sysop", "GameIconsClassic", "sysop"), ("mod-gavel", "GameIconsClassic", "mod-gavel"),
        ("mega", "GameIconsClassic", "mega"), ("ignore", "GameIconsClassic", "ignore"),
        ("sc2", "GameIconsClassic", "sc2"),
    ];

    /// <summary>
    /// "Warcraft III Classic" — the 64x64 set under Assets/GameIconsHD, sourced directly from
    /// classic.battle.net/war3/images/battle.net/icons/ (the WC3 ladder site's own icon folder —
    /// true Blizzard-hosted originals, not a third-party re-host, confirmed 2026-08-24 per the
    /// WC3 ladder icons.shtml page listing every file in that folder). sysop uses
    /// "bnet-battlenet.gif" (the real Admin icon, not a fallback), mega uses "bnet-speaker.gif"
    /// (the real Speaker icon) — both explicit user corrections replacing an earlier
    /// Battle.net-1.0-Classic-borrowed placeholder. jsc/sware/dshr still have no distinct HD art
    /// upstream, so they borrow their closest relative (StarCraft/StarCraft/Diablo); war3/w3tft
    /// aren't in that same folder either (it's a ladder-status icon set, not per-product game
    /// icons) and stay sourced from wowpedia.fandom.com/wiki/Warcraft_III_chat_icons instead. sc2
    /// is user-made (2026-08-24, "bnet-sc2_war3_style.png") — StarCraft II postdates this whole
    /// art style, so there's no official upstream icon.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Wc3ClassicSet =
    [
        ("sc", "GameIconsHD", "sc"), ("scbw", "GameIconsHD", "scbw"), ("jsc", "GameIconsHD", "sc"),
        ("sware", "GameIconsHD", "sc"), ("war2", "GameIconsHD", "war2"), ("war3", "GameIconsHD", "war3"),
        ("w3tft", "GameIconsHD", "w3tft"), ("diablo", "GameIconsHD", "diablo"), ("dshr", "GameIconsHD", "diablo"),
        ("diablo2", "GameIconsHD", "diablo2"), ("d2exp", "GameIconsHD", "d2exp"), ("chat", "GameIconsHD", "chat"),
        ("blizz", "GameIconsHD", "blizz"), ("sysop", "GameIconsHD", "sysop"),
        ("mod-gavel", "GameIconsHD", "mod-gavel"), ("mega", "GameIconsHD", "mega"),
        ("ignore", "GameIconsHD", "ignore"), ("sc2", "GameIconsHD", "sc2"),
    ];

    /// <summary>
    /// "Battle.net 2.0" — the official account.battle.net game-icon set (rasterized SVGs under
    /// Assets/GameIconsBnet2), plus the Warcraft III Classic set's status badges per explicit
    /// request ("for Battle.net 2.0 we'll want to use the status badges from War3 set") — with one
    /// exception: mod-gavel uses "mod-gavel-glow.png", a green-glow variant of the same War3
    /// hammer (generated 2026-08-24 via a blurred green-tinted copy of the icon's own silhouette
    /// drawn behind the crisp original — "a gentle green glow around it," per request) rather than
    /// the plain one, to read as more at-home next to Bnet2's brighter modern art. sc2 uses the
    /// real official account.battle.net StarCraft II SVG (account.battle.net/static/images/
    /// game-icons/starcraft-ii.svg) — fixes a real bug where it had no entry in this set at all,
    /// so switching to Battle.net 2.0 after Battle.net 1.0 Classic/Warcraft III Classic left
    /// whichever classic-style sc2 icon was applied stuck in place instead of reverting. Several
    /// keys intentionally share one source image, matching how Blizzard's own modern branding
    /// doesn't distinguish them: StarCraft/Brood War/the Japanese release/the shareware trial all
    /// point at one "StarCraft: Remastered" icon, same idea for Warcraft III/TFT and Diablo II/
    /// Lord of Destruction.
    /// </summary>
    private static readonly (string Key, string Folder, string SourceKey)[] Bnet2Set =
    [
        ("diablo2", "GameIconsBnet2", "diablo-ii"),
        ("d2exp", "GameIconsBnet2", "diablo-ii"),
        ("war3", "GameIconsBnet2", "warcraft-iii"),
        ("w3tft", "GameIconsBnet2", "warcraft-iii"),
        ("war2", "GameIconsBnet2", "warcraft-ii-remastered"),
        ("sc", "GameIconsBnet2", "starcraft-remastered"),
        ("scbw", "GameIconsBnet2", "starcraft-remastered"),
        ("jsc", "GameIconsBnet2", "starcraft-remastered"),
        ("sware", "GameIconsBnet2", "starcraft-remastered"),
        ("sc2", "GameIconsBnet2", "starcraft-ii"),
        ("blizz", "GameIconsHD", "blizz"),
        ("sysop", "GameIconsHD", "sysop"),
        ("mod-gavel", "GameIconsBnet2", "mod-gavel-glow"),
        ("mega", "GameIconsHD", "mega"),
        ("ignore", "GameIconsHD", "ignore"),
        ("diablo", "GameIconsBnet2", "diablo"),
        ("dshr", "GameIconsBnet2", "diablo"),
    ];

    public ObservableCollection<IconSlotViewModel> GameIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> StatusIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> FriendIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> CustomIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> Bnet2Icons { get; } = [];

    /// <summary>Names of user-saved icon sets, each an ordinary folder under IconSetStore.Directory (inside the app's config folder, so it travels with any config backup).</summary>
    public ObservableCollection<string> SavedSets { get; } = [];

    /// <summary>The three bundled sets (BundledSetNames) plus every user-saved one (SavedSets), in that order — one combined Apply dropdown per explicit request, instead of two separate UI areas for "pick a bundled set" vs "pick a saved set".</summary>
    public ObservableCollection<string> AvailableIconSets { get; } = [];

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

    private static void ApplyIconSetAsset(string targetKey, string folder, string sourceKey)
    {
        var uri = new Uri($"avares://Invigoration.App/Assets/{folder}/{sourceKey}.png");
        using var stream = AssetLoader.Open(uri);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        IconOverrideStore.SetOverrideBytes(targetKey, buffer.ToArray(), ".png");
    }

    private void ApplyBundledSet((string Key, string Folder, string SourceKey)[] mapping)
    {
        foreach (var (key, folder, sourceKey) in mapping)
        {
            ApplyIconSetAsset(key, folder, sourceKey);
        }

        RefreshAllSlots();
    }

    private void RefreshAllSlots()
    {
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
        }

        RefreshAllSlots();
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

    /// <summary>
    /// One Apply command for the whole combined dropdown (AvailableIconSets) — dispatches to the
    /// three bundled sets or, for anything else, a user-saved one via IconSetStore. Replaces the
    /// old separate "Apply Bundled 64x64 Set"/"Apply Battle.net 2.0 Icon Set" buttons plus a
    /// second saved-sets-only dropdown, per explicit request for one unified list.
    /// </summary>
    [RelayCommand]
    private void ApplySelectedSet()
    {
        switch (SelectedSet)
        {
            case null:
                return;
            case Bnet1ClassicSetName:
                ApplyBundledSet(Bnet1ClassicSet);
                return;
            case Wc3ClassicSetName:
                ApplyBundledSet(Wc3ClassicSet);
                return;
            case Bnet2SetName:
                ApplyBundledSet(Bnet2Set);
                return;
            default:
                IconSetStore.ApplySet(SelectedSet);
                RefreshAllSlots();
                return;
        }
    }

    /// <summary>Only meaningful for a user-saved set — a no-op if a bundled set (BundledSetNames) is selected, since those aren't files to delete.</summary>
    [RelayCommand]
    private void DeleteSelectedSet()
    {
        if (SelectedSet is null || BundledSetNames.Contains(SelectedSet))
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

        AvailableIconSets.Clear();
        foreach (var name in BundledSetNames.Concat(SavedSets))
        {
            AvailableIconSets.Add(name);
        }

        SelectedSet = selected is not null && AvailableIconSets.Contains(selected) ? selected : null;
    }

    private static IconSlotViewModel CreateSlot(string key, string displayName)
    {
        var slot = new IconSlotViewModel(key, displayName);
        slot.Refresh();
        return slot;
    }
}
