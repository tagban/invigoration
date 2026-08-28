using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;

namespace Invigoration.App.ViewModels;

/// <summary>
/// A Discord user seen talking through this server's relay bot recently — not a real Hotline
/// connection at all (no UserId/IconId/Flags from the protocol), so it's a deliberately separate,
/// simpler type from HotlineUserRowViewModel rather than forcing a synthetic HotlineUser through
/// the same shape. Listed in its own "Discord" section of the Users panel, per explicit request to
/// keep these visually separated from real server users. Expires — see
/// HotlineSessionViewModel.PruneStaleGhosts — after ~30 minutes of no relayed activity.
/// </summary>
public sealed partial class HotlineGhostUserViewModel : ViewModelBase
{
    private readonly HotlineTabViewModel _parent;

    public string Name { get; }

    public DateTimeOffset LastSeen { get; set; }

    public Bitmap? Icon => GameIconLoader.Get("discord-relay");

    public IBrush? HighlightBrush => _parent.Config.UserHighlightColors.TryGetValue(Name, out var hex)
        ? new SolidColorBrush(Color.Parse(hex))
        : null;

    public IBrush DisplayBrush => HighlightBrush ?? Brushes.Black;

    public HotlineGhostUserViewModel(HotlineTabViewModel parent, string name)
    {
        _parent = parent;
        Name = name;
        LastSeen = DateTimeOffset.UtcNow;
    }

    [RelayCommand]
    private void SetHighlightColor(string hex)
    {
        _parent.Config.UserHighlightColors[Name] = hex;
        _parent.NotifyConfigChanged();
        NotifyHighlightChanged();
    }

    [RelayCommand]
    private void ClearHighlightColor()
    {
        _parent.Config.UserHighlightColors.Remove(Name);
        _parent.NotifyConfigChanged();
        NotifyHighlightChanged();
    }

    private void NotifyHighlightChanged()
    {
        OnPropertyChanged(nameof(HighlightBrush));
        OnPropertyChanged(nameof(DisplayBrush));
    }
}
