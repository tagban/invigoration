using CommunityToolkit.Mvvm.ComponentModel;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Thin wrapper around one saved HotlineServerProfile, purely to carry the "is a session currently
/// connected from this profile" indicator — HotlineServerProfile itself is a plain data class (no
/// INotifyPropertyChanged), shared as-is with the Core layer's store/serialization, so it can't
/// hold reactive UI-only state. Every editable field is still bound straight through to Profile
/// (see HotlineTrackerView.axaml's "Profile.Name" etc.) rather than duplicated here.
/// </summary>
public sealed partial class HotlineServerProfileViewModel : ViewModelBase
{
    public HotlineServerProfile Profile { get; }

    public HotlineServerProfileViewModel(HotlineServerProfile profile)
    {
        Profile = profile;
        IconId = profile.IconId;

        // Explicitly, not relying on the setter above: an icon of 0 is no change from the field's
        // default and would leave the preview blank.
        _ = RefreshIconPreviewAsync();
    }

    [ObservableProperty]
    public partial bool IsConnected { get; set; }

    /// <summary>
    /// This profile's Hotline icon number, mirrored here rather than bound straight to
    /// Profile.IconId so changing it can also refresh the preview — HotlineServerProfile is a plain
    /// data class with no change notification of its own.
    ///
    /// Until this existed there was no way to set a saved server's icon at all: the editor had a
    /// nickname box and nothing else, so every profile stayed on the default 414 (the generic
    /// page) no matter what the tracker's own default icon was set to. That setting only ever
    /// applied to connecting straight from the tracker's server list.
    /// </summary>
    [ObservableProperty]
    public partial int IconId { get; set; }

    partial void OnIconIdChanged(int value)
    {
        var clamped = (ushort)Math.Clamp(value, 0, ushort.MaxValue);
        Profile.IconId = clamped;
        _ = RefreshIconPreviewAsync(clamped);
    }

    /// <summary>The icon itself, so the number isn't the only thing to go on when picking one.</summary>
    [ObservableProperty]
    public partial Avalonia.Media.Imaging.Bitmap? IconPreview { get; set; }

    public async Task RefreshIconPreviewAsync(ushort? iconId = null) =>
        IconPreview = await Models.HotlineIconLoader.GetAsync(iconId ?? Profile.IconId).ConfigureAwait(true);
}
