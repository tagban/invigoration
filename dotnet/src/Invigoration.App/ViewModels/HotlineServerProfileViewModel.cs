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
    }

    [ObservableProperty]
    public partial bool IsConnected { get; set; }
}
