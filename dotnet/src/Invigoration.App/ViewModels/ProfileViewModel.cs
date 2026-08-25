using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.Core;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Backs ProfileWindow — a real Battle.net 1.0 profile (Sex/Age/Location/Description), fetched
/// via BotEngine.RequestProfileAsync (SID_READUSERDATA) on open. Editable only when viewing your
/// own account (SID_WRITEUSERDATA can only ever write the logged-in account's own profile).
/// </summary>
public sealed partial class ProfileViewModel : ViewModelBase
{
    private readonly BotEngine _engine;

    public string Account { get; }

    public bool IsOwnAccount { get; }

    [ObservableProperty]
    public partial bool IsLoading { get; set; } = true;

    [ObservableProperty]
    public partial string Sex { get; set; } = "";

    [ObservableProperty]
    public partial string Age { get; set; } = "";

    [ObservableProperty]
    public partial string Location { get; set; } = "";

    [ObservableProperty]
    public partial string Description { get; set; } = "";

    [ObservableProperty]
    public partial string StatusMessage { get; set; } = "";

    public ProfileViewModel(BotEngine engine, string account)
    {
        _engine = engine;
        Account = account;
        IsOwnAccount = string.Equals(account, engine.OwnChatIdentity ?? engine.Config.Username, StringComparison.OrdinalIgnoreCase);
        _ = LoadAsync();
    }

    private async Task LoadAsync()
    {
        var profile = await _engine.RequestProfileAsync(Account).ConfigureAwait(true);
        IsLoading = false;
        if (profile is null)
        {
            StatusMessage = "No response from Battle.net — profile lookups may not be supported on this connection.";
            return;
        }

        Sex = profile.Sex;
        Age = profile.Age;
        Location = profile.Location;
        Description = profile.Description;
    }

    [RelayCommand]
    private async Task SaveAsync()
    {
        StatusMessage = "Saving...";
        await _engine.WriteProfileAsync(Sex, Location, Description).ConfigureAwait(true);
        StatusMessage = "Saved.";
    }
}
