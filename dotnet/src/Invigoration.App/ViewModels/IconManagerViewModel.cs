using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
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
    public ObservableCollection<IconSlotViewModel> GameIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> StatusIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> FriendIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> CustomIcons { get; } = [];

    public ObservableCollection<IconSlotViewModel> Bnet2Icons { get; } = [];

    /// <summary>Names of user-saved icon sets, each an ordinary folder under IconSetStore.Directory (inside the app's config folder, so it travels with any config backup).</summary>
    public ObservableCollection<string> SavedSets { get; } = [];

    /// <summary>The three bundled sets plus every user-saved one (SavedSets), in that order — one combined Apply dropdown per explicit request, instead of two separate UI areas for "pick a bundled set" vs "pick a saved set".</summary>
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

        IconSetStore.ActiveSetName = "";
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
        if (SelectedSet is null)
        {
            return;
        }

        IconSets.Apply(SelectedSet);
        RefreshAllSlots();
    }

    /// <summary>Only meaningful for a user-saved set — a no-op if a bundled set is selected, since those aren't files to delete.</summary>
    [RelayCommand]
    private void DeleteSelectedSet()
    {
        if (SelectedSet is null || IconSets.IsBundled(SelectedSet))
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
        foreach (var name in IconSets.All())
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
