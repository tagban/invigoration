using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

/// <summary>
/// Parses the actual dotnet/TriviaPacks/*.txt files committed to the repo —
/// the ones TriviaPackDownloader fetches from GitHub into a fresh install's
/// Trivia folder — so a typo breaking one of these in production (silently
/// dropped by TriviaBank.LoadAll's per-line try/catch) gets caught here
/// instead of only showing up as a smaller-than-expected question count
/// after a real download.
/// </summary>
public class TriviaPacksTests
{
    private static string FindTriviaPacksDirectory()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "TriviaPacks");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate dotnet/TriviaPacks by walking up from the test binary's output directory.");
    }

    [Fact]
    public void AllPackFiles_ParseWithoutErrors_AndAreNonEmpty()
    {
        var directory = FindTriviaPacksDirectory();
        var files = Directory.GetFiles(directory, "*.txt");
        Assert.NotEmpty(files);

        foreach (var file in files)
        {
            var defaultCategory = Path.GetFileNameWithoutExtension(file);
            var lineNumber = 0;
            foreach (var rawLine in File.ReadAllLines(file))
            {
                lineNumber++;
                var line = rawLine.Trim();
                if (line.Length == 0)
                {
                    continue;
                }

                var exception = Record.Exception(() => TriviaQuestion.Parse(line, defaultCategory));
                Assert.True(exception is null, $"{Path.GetFileName(file)}:{lineNumber} failed to parse: {exception?.Message}");
            }
        }
    }
}
