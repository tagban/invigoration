namespace Invigoration.Core.Trivia;

/// <summary>
/// Seeds the Trivia config folder with the base question packs on first run,
/// downloaded from this project's own GitHub repo, so they land as normal
/// ".txt" files a user can freely edit/delete/add to instead of only living
/// in <see cref="TriviaBank"/>'s embedded fallback pack. Only ever runs when
/// the folder has no ".txt" files yet — an existing file (whether downloaded
/// before, hand-edited, or entirely custom) is never touched or overwritten.
/// </summary>
public static class TriviaPackDownloader
{
    private const string BaseUrl = "https://raw.githubusercontent.com/tagban/invigoration/main/dotnet/TriviaPacks/";

    /// <summary>File names under dotnet/TriviaPacks/ in the repo — must stay in sync with what's actually committed there.</summary>
    private static readonly string[] PackFiles =
    [
        "Diablo.txt", "Warcraft.txt", "StarCraft.txt", "Blizzard.txt", "PopCulture.txt", "Music.txt", "Apple.txt",
    ];

    /// <summary>
    /// Best-effort: any failure (offline, GitHub unreachable, a file 404s,
    /// etc.) is caught and reported via <paramref name="onError"/> rather
    /// than thrown — this must never block app startup or leave trivia
    /// broken, since <see cref="TriviaBank"/> already falls back to its own
    /// embedded pack whenever the Trivia folder has no ".txt" files.
    /// </summary>
    public static async Task EnsureDownloadedAsync(Action<string>? onError = null)
    {
        try
        {
            Directory.CreateDirectory(TriviaBank.Directory);
            if (Directory.GetFiles(TriviaBank.Directory, "*.txt").Length > 0)
            {
                return;
            }

            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            var downloads = PackFiles.Select(async file =>
            {
                var text = await http.GetStringAsync(BaseUrl + file).ConfigureAwait(false);
                return (file, text);
            });

            var results = await Task.WhenAll(downloads).ConfigureAwait(false);
            foreach (var (file, text) in results)
            {
                await File.WriteAllTextAsync(Path.Combine(TriviaBank.Directory, file), text).ConfigureAwait(false);
            }
        }
        catch (Exception ex)
        {
            onError?.Invoke($"Trivia: couldn't download the base question packs ({ex.Message}) — using the built-in fallback pack instead.");
        }
    }
}
