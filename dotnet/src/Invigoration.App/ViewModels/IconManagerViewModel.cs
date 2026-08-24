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
    /// Keys with a bundled 64x64 alternate available under Assets/GameIconsHD
    /// — the same set classic.battle.net's WC3 ladder pages served
    /// (bnet-*.gif), double-plus the resolution of the default 28x14 set.
    /// Not every key has an HD source; the missing ones (war3/w3tft/jsc/
    /// sware/dshr) just aren't part of that original set.
    /// </summary>
    private static readonly string[] HighResolutionKeys =
        ["blizz", "sysop", "mega", "ignore", "chat", "diablo", "diablo2", "d2exp", "sc", "scbw", "war2"];

    public ObservableCollection<IconSlotViewModel> GameIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> StatusIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> FriendIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> CustomIcons { get; } = [];

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

    /// <summary>Applies the bundled 64x64 icon set to every key that has one — a concrete, one-click example of swapping in a larger icon set, using real Blizzard-hosted assets rather than a hypothetical.</summary>
    [RelayCommand]
    private void ApplyHighResolutionSet()
    {
        foreach (var key in HighResolutionKeys)
        {
            var uri = new Uri($"avares://Invigoration.App/Assets/GameIconsHD/{key}.gif");
            using var stream = AssetLoader.Open(uri);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            IconOverrideStore.SetOverrideBytes(key, buffer.ToArray(), ".gif");
        }

        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons))
        {
            slot.Refresh();
        }
    }

    /// <summary>Clears every override, reverting all icons to the bundled classic 28x14 defaults in one action.</summary>
    [RelayCommand]
    private void ResetAllIcons()
    {
        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons))
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
        foreach (var slot in GameIcons.Concat(StatusIcons).Concat(FriendIcons).Concat(CustomIcons))
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
