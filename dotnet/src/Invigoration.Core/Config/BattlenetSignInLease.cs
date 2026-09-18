using System.Diagnostics.CodeAnalysis;

namespace Invigoration.Core.Config;

/// <summary>
/// One user at a time for a Battle.net profile's saved sign-in: a connecting or connected bot, or
/// the Battle.net Profiles window's Sign In. Held across every bot in this app and every copy of the
/// app running.
/// </summary>
/// <remarks>
/// <para>Every Battle.net login replaces the saved credential with a new one. When a login presents a
/// credential Battle.net has already replaced, Stimpak deletes the file before asking the user to
/// sign in again. So two logins from one file at once end with the older one deleting the credential
/// the newer one just saved. It's one Battle.net account as well, whose chat holds one session: each
/// login takes it from the other.</para>
/// <para>Not tied to any one game. StarCraft II, StarCraft: Remastered and Warcraft III: Reforged all
/// sign in as the same Battle.net account, and so will anything else that does.</para>
/// <para>Other copies of the app are kept out by an OS lock on a file next to the credential (.NET's
/// FileShare.None is flock on macOS and Linux). The OS lets go of it when a process exits, crashed or
/// not, so a lease is never left behind.</para>
/// </remarks>
public sealed class BattlenetSignInLease : IDisposable
{
    private static readonly Lock SyncRoot = new();
    private static readonly Dictionary<string, BattlenetSignInLease> Held = [];

    private readonly FileStream _lockFile;
    private readonly TaskCompletionSource _released = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private int _disposed;

    private BattlenetSignInLease(string profileId, object owner, string ownerName, FileStream lockFile)
    {
        ProfileId = profileId;
        Owner = owner;
        OwnerName = ownerName;
        _lockFile = lockFile;
    }

    public string ProfileId { get; }

    /// <summary>Who took it, so its owner can recognise its own lease.</summary>
    public object Owner { get; }

    /// <summary>Who took it, for the message a would-be user gets: a bot's name, say.</summary>
    public string OwnerName { get; }

    /// <summary>Completes once this lease is let go of.</summary>
    public Task Released => _released.Task;

    /// <summary>
    /// Takes the profile's sign-in if nobody has it. Otherwise <paramref name="holder"/> is whoever in
    /// this app does, or null when it's another copy of the app.
    /// </summary>
    public static bool TryAcquire(
        string profileId,
        object owner,
        string ownerName,
        [NotNullWhen(true)] out BattlenetSignInLease? lease,
        out BattlenetSignInLease? holder)
    {
        lock (SyncRoot)
        {
            lease = null;
            if (Held.TryGetValue(profileId, out holder))
            {
                return false;
            }

            FileStream lockFile;
            try
            {
                var path = LockFilePath(profileId);
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                lockFile = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return false;
            }

            lease = new BattlenetSignInLease(profileId, owner, ownerName, lockFile);
            Held[profileId] = lease;
            return true;
        }
    }

    /// <summary>The file whose OS lock keeps other copies of the app out. Left in place afterwards, empty; removing it could race another copy opening it.</summary>
    public static string LockFilePath(string profileId) =>
        BattlenetCredentialProfileStore.CredentialFilePath(profileId) + ".lock";

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
        {
            return;
        }

        lock (SyncRoot)
        {
            _lockFile.Dispose();
            if (Held.TryGetValue(ProfileId, out var current) && ReferenceEquals(current, this))
            {
                Held.Remove(ProfileId);
            }
        }

        _released.TrySetResult();
    }
}
