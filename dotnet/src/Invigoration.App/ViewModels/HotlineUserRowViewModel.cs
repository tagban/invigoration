using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Wraps a HotlineUser for display — HotlineUser itself is a plain immutable record from the
/// protocol layer with no room for an asynchronously-loaded icon Bitmap, so this is the view-side
/// row that fetches one via HotlineIconLoader and exposes it as a bindable property once it
/// arrives (icons load after the row itself appears, not before — never blocks the user list on a
/// network fetch).
/// </summary>
public sealed partial class HotlineUserRowViewModel : ViewModelBase
{
    private readonly HotlineTabViewModel _parent;

    public HotlineUser User { get; }

    public ushort UserId => User.UserId;
    public string Name => User.Name;

    [ObservableProperty]
    public partial Bitmap? Icon { get; set; }

    /// <summary>A right-click "highlight this person" override — see HotlineTrackerConfig.UserHighlightColors' remarks. Null means no override; the chat log falls back to the normal Admin/Default rank color.</summary>
    public IBrush? HighlightBrush => _parent.Config.UserHighlightColors.TryGetValue(Name, out var hex)
        ? new SolidColorBrush(Color.Parse(hex))
        : null;

    /// <summary>
    /// What the Users list row itself renders the name in — the highlight color if set, otherwise
    /// the same Admin/Default rank coloring the chat log already uses (see
    /// HotlineSessionViewModel.AppendColorizedLine) so an admin shows the same color in both
    /// places; plain black for a non-admin (the panel's own light background assumes black as the
    /// default, unlike the dark chat log's white default). Previously fell straight through to
    /// black regardless of Admin status — a real bug, fixed per direct user report ("Usernames in
    /// Hotline are red in chat for admins but not in the userlist and should be").
    /// </summary>
    public IBrush DisplayBrush => HighlightBrush ?? (User.IsAdmin ? new SolidColorBrush(Color.Parse(_parent.Config.AdminColorHex)) : Brushes.Black);

    public HotlineUserRowViewModel(HotlineTabViewModel parent, HotlineUser user)
    {
        _parent = parent;
        User = user;
        _ = LoadIconAsync();
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

    private async Task LoadIconAsync()
    {
        var icon = await HotlineIconLoader.GetAsync(User.IconId).ConfigureAwait(true);
        Dispatcher.UIThread.Post(() => Icon = icon);
    }
}
