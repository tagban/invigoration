using System.Collections.ObjectModel;
using System.Text;
using Avalonia.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Invigoration.App.Models;
using Invigoration.Core.Config;
using Stimpak;

namespace Invigoration.App.ViewModels;

/// <summary>
/// Manages the shared (cross-bot) list of named Battle.net logins — see
/// BattlenetCredentialProfileStore. Reusing one across bots (e.g. an SC2 bot
/// and a future WC3:Reforged bot on the same account) shares that one
/// signed-in session; separate profiles keep separate logins.
/// </summary>
public partial class BattlenetCredentialProfilesViewModel : ObservableObject
{
    public ObservableCollection<BattlenetCredentialProfileViewModel> Profiles { get; }

    [ObservableProperty]
    public partial string? StatusMessage { get; set; }

    public BattlenetCredentialProfilesViewModel()
    {
        Profiles = new ObservableCollection<BattlenetCredentialProfileViewModel>(
            BattlenetCredentialProfileStore.Profiles.Select(p => new BattlenetCredentialProfileViewModel(p)));
    }

    [RelayCommand]
    private void AddProfile()
    {
        var profile = BattlenetCredentialProfileStore.CreateAndSave("New Profile");
        Profiles.Add(new BattlenetCredentialProfileViewModel(profile));
    }

    [RelayCommand]
    private void RemoveProfile(BattlenetCredentialProfileViewModel profile)
    {
        var botsUsingIt = new ConfigStore().Load()
            .Where(b => b.BattlenetCredentialProfileId == profile.Profile.Id)
            .Select(b => b.DisplayName)
            .ToList();
        if (botsUsingIt.Count > 0)
        {
            StatusMessage = $"\"{profile.Name}\" is still used by: {string.Join(", ", botsUsingIt)}. Change those bots' Battle.net Profile first.";
            return;
        }

        StatusMessage = null;
        BattlenetCredentialProfileStore.Delete(profile.Profile.Id);
        Profiles.Remove(profile);
    }

    [RelayCommand]
    private void Save() => BattlenetCredentialProfileStore.Save();
}

/// <summary>Wraps a BattlenetCredentialProfile for the "Manage Battle.net Profiles" editor, adding a standalone sign-in action.</summary>
public partial class BattlenetCredentialProfileViewModel : ObservableObject
{
    public BattlenetCredentialProfile Profile { get; }

    [ObservableProperty]
    public partial string Name { get; set; }

    [ObservableProperty]
    public partial bool IsSignedIn { get; set; }

    [ObservableProperty]
    public partial bool IsSigningIn { get; set; }

    public BattlenetCredentialProfileViewModel(BattlenetCredentialProfile profile)
    {
        Profile = profile;
        Name = profile.Name;
        RefreshIsSignedIn();
    }

    partial void OnNameChanged(string value) => Profile.Name = value;

    private void RefreshIsSignedIn() => IsSignedIn = BattlenetCredentialProfileStore.HasCachedCredential(Profile.Id);

    /// <summary>
    /// Standalone verification/sign-in: spins up a throwaway StimpakClient
    /// against this profile's own credential file (the exact one
    /// BotEngine.Sc2.cs will use for any bot assigned to this profile),
    /// forces the interactive flow, and waits for either a successful
    /// connect or a failure before disposing. Uses the same
    /// Sc2LoginChallenge popup BotTabView's real connect path falls back to
    /// when Stimpak's own native auth window isn't available.
    /// </summary>
    public async Task SignInAsync(TopLevel owner)
    {
        IsSigningIn = true;
        try
        {
            using var client = new StimpakClient(BattlenetCredentialProfileStore.CredentialFilePath(Profile.Id));
            var outcome = new TaskCompletionSource();
            client.EventReceived += next =>
            {
                switch (next)
                {
                    case StageChanged { Stage: Stage.Connected }:
                        outcome.TrySetResult();
                        break;
                    case AuthenticationRequired auth:
                        _ = SubmitChallengeAsync(client, auth, owner, outcome);
                        break;
                    case SessionFailed failed:
                        outcome.TrySetException(new InvalidOperationException(failed.Message));
                        break;
                    case SessionEnded:
                        outcome.TrySetException(new InvalidOperationException("Sign-in ended before completing."));
                        break;
                }
            };

            client.Connect(forceInteractive: true);

            using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(3));
            await using (timeout.Token.Register(() => outcome.TrySetCanceled()).ConfigureAwait(false))
            {
                await outcome.Task.ConfigureAwait(false);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Sign-in failed: {ex.Message}", ex);
        }
        finally
        {
            IsSigningIn = false;
            RefreshIsSignedIn();
        }
    }

    private static async Task SubmitChallengeAsync(StimpakClient client, AuthenticationRequired auth, TopLevel owner, TaskCompletionSource outcome)
    {
        try
        {
            var credential = await Sc2LoginChallenge.ShowAsync(owner, new Uri(auth.Url)).ConfigureAwait(false);
            client.SubmitAuth(auth.AuthId, Encoding.UTF8.GetString(credential));
        }
        catch (Exception ex)
        {
            outcome.TrySetException(ex);
        }
    }
}
