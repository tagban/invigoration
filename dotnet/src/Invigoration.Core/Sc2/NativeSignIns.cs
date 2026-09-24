using System.Collections.Concurrent;
using Invigoration.Core.Config;

namespace Invigoration.Core.Sc2;

/// <summary>
/// One Battle.net sign-in for every game on a profile. Each game logs in as itself with its own
/// saved sign-in, but a signed-in session can ask Battle.net for another game's too, so the first
/// game to sign in saves them for the rest. And when bots on one profile all need a web sign-in at
/// once (auto-connect at startup, nothing saved yet), only one window opens: the others wait for
/// it, then use the sign-in it saved for them.
/// </summary>
public static class NativeSignIns
{
    /// <summary>The games with a native connection, by program code.</summary>
    public static IReadOnlyList<string> Programs { get; } = [NativeSc2ChatClient.Program, NativeScrChatClient.Program];

    private static readonly ConcurrentDictionary<string, SemaphoreSlim> Gates = new();

    public static string GameName(string program) => program switch
    {
        NativeSc2ChatClient.Program => "StarCraft II",
        NativeScrChatClient.Program => "StarCraft: Remastered",
        _ => program,
    };

    /// <summary>
    /// Takes the profile's web sign-in turn, waiting while another game on it has the window open.
    /// Throws <see cref="SavedSignInArrivedException"/> instead when, by then, a new saved sign-in for
    /// <paramref name="program"/> has turned up (the other game got one for it): the caller starts
    /// over with that rather than asking the user again.
    /// </summary>
    public static async Task<IDisposable> BeginWebSignInAsync(string profileId, string program, byte[]? presented, CancellationToken cancellationToken)
    {
        var gate = Gates.GetOrAdd(profileId, _ => new SemaphoreSlim(1, 1));
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        var saved = BattlenetCredentialProfileStore.LoadNativeCredential(profileId, program);
        if (saved is not null && (presented is null || !saved.AsSpan().SequenceEqual(presented)))
        {
            gate.Release();
            throw new SavedSignInArrivedException(GameName(program));
        }

        return new Turn(gate);
    }

    /// <summary>
    /// Asks the signed-in session for a saved sign-in for each other game this profile has none
    /// for yet, and saves it. Never replaces one: that could log a running bot's next attempt out.
    /// </summary>
    public static async Task IssueMissingAsync(
        string profileId,
        string signedInAs,
        Func<string, CancellationToken, Task<byte[]?>> generate,
        Action<string> trace,
        CancellationToken cancellationToken)
    {
        foreach (var program in Programs.Where(p => p != signedInAs))
        {
            if (BattlenetCredentialProfileStore.LoadNativeCredential(profileId, program) is not null)
            {
                continue;
            }

            try
            {
                if (await generate(program, cancellationToken).ConfigureAwait(false) is { Length: > 0 } credential)
                {
                    BattlenetCredentialProfileStore.SaveNativeCredential(profileId, program, credential);
                    trace($"Also saved a sign-in for {GameName(program)} on this profile.");
                }
                else
                {
                    trace($"Battle.net issued no sign-in for {GameName(program)}.");
                }
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                trace($"Couldn't get a sign-in for {GameName(program)}: {ex.Message}");
            }
        }
    }

    private sealed class Turn(SemaphoreSlim gate) : IDisposable
    {
        private int _released;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _released, 1) == 0)
            {
                gate.Release();
            }
        }
    }
}

/// <summary>Another game on the profile signed in meanwhile and saved a sign-in for this one; start over with it.</summary>
public sealed class SavedSignInArrivedException(string game)
    : Exception($"Another game on this Battle.net profile just signed in and saved a sign-in for {game}; using that.");
