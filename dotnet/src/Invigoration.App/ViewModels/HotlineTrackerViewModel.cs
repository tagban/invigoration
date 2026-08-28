using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core.Hotline;

namespace Invigoration.App.ViewModels;

/// <summary>The Hotline tracker's permanent first sub-tab — a tracker query view plus the saved server profiles list. "Tracker" is a fixed header (unlike HotlineSessionViewModel's host:port), matching the original request's "the primary window starts with just the tracker".</summary>
public sealed partial class HotlineTrackerViewModel : ViewModelBase
{
    private readonly HotlineTabViewModel _parent;
    private List<HotlineTrackerServerEntry> _allServers = [];

    public string Title => "Tracker";
    public Avalonia.Media.IBrush HighlightBrush => HotlineTabViewModel.AccentBrush;
    public double HeaderFontSize => 13;
    public Avalonia.Media.IBrush HeaderForeground => Avalonia.Media.Brushes.White;
    public bool HasUnread => false;

    /// <summary>This tracker's own top-level tab name — editable here rather than through a separate "Add Hotline Tracker" dialog, same in-place-editing idiom as a saved server profile's own fields below.</summary>
    public string DisplayName
    {
        get => _parent.Config.DisplayName;
        set
        {
            if (_parent.Config.DisplayName == value)
            {
                return;
            }

            _parent.Config.DisplayName = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    public string TrackerHost
    {
        get => _parent.Config.TrackerHost;
        set
        {
            if (_parent.Config.TrackerHost == value)
            {
                return;
            }

            _parent.Config.TrackerHost = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>The nickname/icon used connecting straight from the tracker server list below — "to use on all servers", set once here rather than per-server. A saved profile's own Nickname/IconId still wins for a profile-based connect.</summary>
    public string DefaultNickname
    {
        get => _parent.Config.DefaultNickname;
        set
        {
            if (_parent.Config.DefaultNickname == value)
            {
                return;
            }

            _parent.Config.DefaultNickname = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    public int DefaultIconId
    {
        get => _parent.Config.DefaultIconId;
        set
        {
            var clamped = (ushort)Math.Clamp(value, 0, ushort.MaxValue);
            if (_parent.Config.DefaultIconId == clamped)
            {
                return;
            }

            _parent.Config.DefaultIconId = clamped;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>
    /// Chat username color for anyone with the Admin flag set — see HotlineTrackerConfig's remarks
    /// on why this is 2-tier only (Admin vs everyone else), not a 3rd "Mod" tier the protocol
    /// doesn't actually distinguish for other users.
    /// </summary>
    public Avalonia.Media.Color AdminColor
    {
        get => Avalonia.Media.Color.Parse(_parent.Config.AdminColorHex);
        set
        {
            var hex = value.ToString();
            if (_parent.Config.AdminColorHex == hex)
            {
                return;
            }

            _parent.Config.AdminColorHex = hex;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>Chat username color for everyone without the Admin flag.</summary>
    public Avalonia.Media.Color DefaultColor
    {
        get => Avalonia.Media.Color.Parse(_parent.Config.DefaultColorHex);
        set
        {
            var hex = value.ToString();
            if (_parent.Config.DefaultColorHex == hex)
            {
                return;
            }

            _parent.Config.DefaultColorHex = hex;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>Logs every inbound transaction on every session connected under this tracker — see HotlineTransactionClient.DebugLog's remarks. Grouped under the collapsed "Advanced" section.</summary>
    public bool Debug
    {
        get => _parent.Config.Debug;
        set
        {
            if (_parent.Config.Debug == value)
            {
                return;
            }

            _parent.Config.Debug = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>Off by default — shows the "Copy Log" button on every session connected under this tracker. Grouped under the collapsed "Advanced" section.</summary>
    public bool ShowCopyLogButton
    {
        get => _parent.Config.ShowCopyLogButton;
        set
        {
            if (_parent.Config.ShowCopyLogButton == value)
            {
                return;
            }

            _parent.Config.ShowCopyLogButton = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    /// <summary>Whether the whole tracker-setup block (name/host/identity/chat colors/advanced) is shown expanded — see HotlineTrackerConfig.SettingsExpanded's remarks. Bound TwoWay so manually collapsing/expanding it persists too, not just the automatic first-success collapse.</summary>
    public bool SettingsExpanded
    {
        get => _parent.Config.SettingsExpanded;
        set
        {
            if (_parent.Config.SettingsExpanded == value)
            {
                return;
            }

            _parent.Config.SettingsExpanded = value;
            OnPropertyChanged();
            _parent.NotifyConfigChanged();
        }
    }

    [ObservableProperty]
    public partial bool IsRefreshing { get; set; }

    [ObservableProperty]
    public partial string StatusText { get; set; } = "";

    [ObservableProperty]
    public partial string SearchText { get; set; } = "";

    partial void OnSearchTextChanged(string value) => ApplyFilter();

    public ObservableCollection<HotlineTrackerServerEntry> TrackerServers { get; } = [];

    public ObservableCollection<HotlineServerProfileViewModel> SavedProfiles { get; } = [];

    public HotlineTrackerViewModel(HotlineTabViewModel parent)
    {
        _parent = parent;
        ReloadProfiles();
        HotlineServerProfileStore.ProfilesChanged += ReloadProfiles;

        // A saved server row's connected indicator/quick-open (see RefreshConnectionStatus) needs
        // to react to every session connect and disconnect under this tracker — Items is where
        // both actually happen (HotlineTabViewModel.Connect/CloseSession add/remove from it).
        _parent.Items.CollectionChanged += (_, _) => RefreshConnectionStatus();
    }

    private void ReloadProfiles()
    {
        SavedProfiles.Clear();
        foreach (var profile in HotlineServerProfileStore.Profiles)
        {
            SavedProfiles.Add(new HotlineServerProfileViewModel(profile));
        }

        RefreshConnectionStatus();
    }

    /// <summary>Marks each saved profile row as connected if any currently-open session (see HotlineSessionViewModel.ProfileId) was started from it — an ad-hoc tracker-list connect has no ProfileId and never matches. Re-run whenever a session opens/closes or the profile list itself reloads.</summary>
    private void RefreshConnectionStatus()
    {
        var connectedProfileIds = _parent.Items.OfType<HotlineSessionViewModel>()
            .Select(s => s.ProfileId)
            .Where(id => id is not null)
            .ToHashSet();

        foreach (var profileVm in SavedProfiles)
        {
            profileVm.IsConnected = connectedProfileIds.Contains(profileVm.Profile.Id);
        }
    }

    /// <summary>Called by HotlineTabViewModel whenever the Tracker sub-tab becomes the selected item (including its very first appearance) — refreshes the server list on activation instead of requiring a manual Refresh click every time, per explicit request.</summary>
    public void OnTabActivated() => _ = RefreshTracker();

    [RelayCommand]
    private async Task RefreshTracker()
    {
        if (IsRefreshing)
        {
            return;
        }

        IsRefreshing = true;
        StatusText = "";
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            _allServers = [.. await HotlineTrackerClient.QueryAsync(TrackerHost, ct: cts.Token).ConfigureAwait(true)];
            ApplyFilter();
            StatusText = _allServers.Count == 0 ? "No servers found (or the tracker didn't respond)." : "";

            // "Settings for the tracker should collapse down once the tracker is setup for the
            // first time" — the tracker having actually returned a real server list is the
            // clearest signal it's genuinely working, not just that the host field is non-empty.
            if (_allServers.Count > 0 && SettingsExpanded)
            {
                SettingsExpanded = false;
            }
        }
        catch (Exception ex) when (ex is System.Net.Sockets.SocketException or OperationCanceledException)
        {
            StatusText = $"Couldn't reach {TrackerHost}: {ex.Message}";
        }
        finally
        {
            IsRefreshing = false;
        }
    }

    private void ApplyFilter()
    {
        TrackerServers.Clear();
        var query = SearchText.Trim();
        var filtered = query.Length == 0
            ? _allServers
            : _allServers.Where(s => s.Name.Contains(query, StringComparison.OrdinalIgnoreCase) || s.Description.Contains(query, StringComparison.OrdinalIgnoreCase));
        foreach (var server in filtered.OrderByDescending(s => s.UserCount))
        {
            TrackerServers.Add(server);
        }
    }

    [RelayCommand]
    private void ConnectToTrackerServer(HotlineTrackerServerEntry server) =>
        _parent.Connect(new HotlineConnectOptions(server.Address, server.Port, Login: "", Password: "", DefaultNickname, _parent.Config.DefaultIconId, server.Name));

    [RelayCommand]
    private void ConnectToProfile(HotlineServerProfile profile) =>
        _parent.Connect(new HotlineConnectOptions(profile.Host, profile.Port, profile.Login, profile.Password, profile.Nickname, profile.IconId, profile.Name, profile.AutoAcceptAgreement, profile.DiscordRelayUsername, profile.DiscordRelayPrefix, profile.ClientVersion, profile.SendClientVersion, profile.Id, profile.TriviaEnabled, profile.AdvertiseChatHistorySupport));

    /// <summary>The saved-servers row's quick action when already connected — switches to the existing session tab instead of opening a second, redundant connection to the same server.</summary>
    [RelayCommand]
    private void OpenConnectedSession(HotlineServerProfile profile)
    {
        var session = _parent.Items.OfType<HotlineSessionViewModel>().FirstOrDefault(s => s.ProfileId == profile.Id);
        if (session is not null)
        {
            _parent.SelectedItem = session;
        }
    }

    [RelayCommand]
    private void AddProfile() => HotlineServerProfileStore.CreateAndSave("New Server", "", HotlineConstants.DefaultServerPort);

    [RelayCommand]
    private void SaveProfile(HotlineServerProfile profile) => HotlineServerProfileStore.Save();

    [RelayCommand]
    private void DeleteProfile(HotlineServerProfile profile) => HotlineServerProfileStore.Delete(profile.Id);

    [RelayCommand]
    private void RemoveTracker() => _parent.RequestRemove();

    /// <summary>Called once at startup by MainWindowViewModel — connects every saved profile with AutoConnect on, same idea as AutoConnectStartupBotsAsync but scoped to Hotline profiles.</summary>
    public void AutoConnectStartupProfiles()
    {
        foreach (var profile in SavedProfiles.Select(vm => vm.Profile).Where(p => p.AutoConnect).ToList())
        {
            _parent.Connect(new HotlineConnectOptions(profile.Host, profile.Port, profile.Login, profile.Password, profile.Nickname, profile.IconId, profile.Name, profile.AutoAcceptAgreement, profile.DiscordRelayUsername, profile.DiscordRelayPrefix, profile.ClientVersion, profile.SendClientVersion, profile.Id, profile.TriviaEnabled, profile.AdvertiseChatHistorySupport));
        }
    }
}
