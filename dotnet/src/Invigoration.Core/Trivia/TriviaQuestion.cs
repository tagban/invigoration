using System.Text;

namespace Invigoration.Core.Trivia;

/// <summary>
/// One trivia question with its accepted answers and progressively-revealing
/// hints, ported from BNU`Bot's TriviaItem
/// (github.com/tagban/bnubot/tree/master/BNUBot/src/net/bnubot/bot/trivia) —
/// same file format and hint-masking logic, so an existing BNU`Bot trivia
/// pack can be dropped into the Trivia config folder unmodified.
/// </summary>
public sealed class TriviaQuestion
{
    public string Category { get; }

    public string QuestionText { get; }

    public IReadOnlyList<string> Answers { get; }

    private readonly IReadOnlyList<string> _answersNormalized;

    /// <summary>Fully masked — shown alongside the question itself.</summary>
    public string Hint0 { get; }

    /// <summary>Partially revealed — shown after 10 seconds.</summary>
    public string Hint1 { get; }

    /// <summary>Most revealed — shown after 20 seconds.</summary>
    public string Hint2 { get; }

    /// <summary>True for an A/B/C/D question built via <see cref="CreateMultipleChoice"/> — no hint reveals (every option is already visible), and always scores the "before first hint" tier since there's no hint progression to speak of.</summary>
    public bool IsMultipleChoice { get; }

    /// <summary>The 2-4 displayed options in shuffled order, each paired with its assigned letter — empty unless <see cref="IsMultipleChoice"/>.</summary>
    public IReadOnlyList<(char Letter, string Text)> Choices { get; } = [];

    public TriviaQuestion(string category, string questionText, IReadOnlyList<string> answers)
    {
        if (answers.Count == 0 || answers.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("A trivia question needs at least one non-empty answer.", nameof(answers));
        }

        Category = category;
        QuestionText = questionText;
        Answers = answers;
        _answersNormalized = answers.Select(Normalize).ToList();
        (Hint0, Hint1, Hint2) = BuildHints(answers[0]);
    }

    private TriviaQuestion(string category, string questionText, IReadOnlyList<(char Letter, string Text)> choices, char correctLetter)
    {
        Category = category;
        QuestionText = questionText;
        IsMultipleChoice = true;
        Choices = choices;
        // Accepts both the letter and the actual option text as a correct answer,
        // reusing the existing Normalize+match plumbing unchanged.
        Answers = [choices.First(c => c.Letter == correctLetter).Text, correctLetter.ToString()];
        _answersNormalized = Answers.Select(Normalize).ToList();
        Hint0 = Hint1 = Hint2 = ""; // never shown/used for multiple choice — every option is already visible
    }

    /// <summary>
    /// Builds an A/B/C/D question: correctAnswer plus 1-3 wrong options,
    /// shuffled into a random display order so the correct choice isn't
    /// always "A". Answers are expected to already be short (1-2 words) —
    /// this doesn't enforce that itself, since it's a content-authoring
    /// concern (respecting BNCS's per-message length limit when all options
    /// are sent on one line), not a parsing one.
    /// </summary>
    public static TriviaQuestion CreateMultipleChoice(string category, string questionText, string correctAnswer, IReadOnlyList<string> wrongAnswers)
    {
        if (string.IsNullOrWhiteSpace(correctAnswer))
        {
            throw new ArgumentException("A multiple-choice question needs a correct answer.", nameof(correctAnswer));
        }
        if (wrongAnswers.Count is < 1 or > 3)
        {
            throw new ArgumentException("A multiple-choice question needs 1-3 wrong answers (2-4 choices total).", nameof(wrongAnswers));
        }

        var order = new List<string> { correctAnswer };
        order.AddRange(wrongAnswers);
        Shuffle(order);

        var letters = "ABCD";
        var choices = order.Select((text, i) => (Letter: letters[i], Text: text)).ToList();
        var correctLetter = choices.First(c => c.Text == correctAnswer).Letter;

        return new TriviaQuestion(category, questionText, choices, correctLetter);
    }

    private static void Shuffle<T>(IList<T> list)
    {
        var random = Random.Shared;
        for (var i = list.Count - 1; i > 0; i--)
        {
            var j = random.Next(i + 1);
            (list[i], list[j]) = (list[j], list[i]);
        }
    }

    /// <summary>True if the given chat text matches any accepted answer once both sides are normalized — e.g. "it's Tyrael!" still matches "Tyrael".</summary>
    public bool TryMatchAnswer(string text, out string matchedAnswer)
    {
        var normalized = Normalize(text);
        if (normalized.Length > 0)
        {
            for (var i = 0; i < _answersNormalized.Count; i++)
            {
                if (_answersNormalized[i] == normalized)
                {
                    matchedAnswer = Answers[i];
                    return true;
                }
            }
        }

        matchedAnswer = "";
        return false;
    }

    /// <summary>Strips everything but letters/digits and folds case, matching BNU`Bot's HexDump.getAlphaNumerics + equalsIgnoreCase comparison.</summary>
    public static string Normalize(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            if (char.IsLetterOrDigit(c))
            {
                builder.Append(char.ToUpperInvariant(c));
            }
        }

        return builder.ToString();
    }

    /// <summary>
    /// Masks the primary answer with '?' in a 2-hidden/1-shown-per-3-characters
    /// pattern across three stages, revealing progressively more each stage;
    /// non-alphanumeric characters (spaces, punctuation) are never masked, so
    /// word boundaries stay visible. Ported 1:1 from TriviaItem.makeHints().
    /// </summary>
    private static (string Hint0, string Hint1, string Hint2) BuildHints(string answer)
    {
        var hint0 = new StringBuilder();
        var hint1 = new StringBuilder();
        var hint2 = new StringBuilder();
        var numHidden = 0;

        foreach (var c in answer)
        {
            if (char.IsLetterOrDigit(c))
            {
                numHidden++;
                if (numHidden % 3 < 2)
                {
                    hint0.Append('?');
                    hint1.Append('?');
                    hint2.Append(numHidden % 3 < 1 ? '?' : c);
                }
                else
                {
                    hint0.Append('?');
                    hint1.Append(c);
                    hint2.Append(c);
                }
            }
            else
            {
                hint0.Append(c);
                hint1.Append(c);
                hint2.Append(c);
            }
        }

        return (hint0.ToString(), hint1.ToString(), hint2.ToString());
    }

    /// <summary>
    /// Parses one line in either of BNU`Bot's two formats:
    /// "/category/answer1/answer2//question text" (explicit category, multiple answers), or
    /// "question text*answer1*answer2" (falls back to '|' as the delimiter if
    /// '*' doesn't split the line into exactly two parts, so either character
    /// can appear in the question/answers when the other is used as delimiter).
    /// "Scramble*word" is a special case: the question becomes the word's
    /// letters shuffled, with the original word as the one accepted answer.
    /// "MC*category*question*correctAnswer*wrong1*wrong2*wrong3" (2-4 total
    /// options — one correct, 1-3 wrong) is a separate, new format for
    /// multiple-choice questions — deliberately not folded into the existing
    /// "multiple accepted answers" semantics above, since those mean
    /// something different (any listed answer counts as correct, not "these
    /// are the choices, only the first one is right").
    /// Throws FormatException on a line matching neither shape, mirroring
    /// TriviaItem's IllegalArgumentException — callers should catch this per
    /// line and skip/log rather than aborting the whole file.
    /// </summary>
    public static TriviaQuestion Parse(string line, string defaultCategory)
    {
        if (line.Length == 0)
        {
            throw new FormatException("Empty line.");
        }

        if (line.StartsWith("MC*", StringComparison.Ordinal))
        {
            var fields = line.Split('*');
            if (fields.Length < 5 || fields.Length > 7)
            {
                throw new FormatException($"Expected MC*category*question*correctAnswer*wrong1[*wrong2*wrong3]: {line}");
            }

            var mcCategory = fields[1];
            var mcQuestion = fields[2];
            var correctAnswer = fields[3];
            var wrongAnswers = fields.Skip(4).ToList();
            return CreateMultipleChoice(mcCategory, mcQuestion, correctAnswer, wrongAnswers);
        }

        if (line[0] == '/')
        {
            var parts = line.Split("//", 2);
            if (parts.Length != 2)
            {
                throw new FormatException($"Missing '//' before the question text: {line}");
            }

            var categoryAndAnswers = parts[0].Split('/');
            if (categoryAndAnswers.Length < 3)
            {
                throw new FormatException($"Expected /category/answer1[/answer2...]: {line}");
            }

            var category = categoryAndAnswers[1];
            var answers = categoryAndAnswers.Skip(2).ToList();
            return new TriviaQuestion(category, parts[1], answers);
        }

        var splitChar = '*';
        var split = line.Split(splitChar, 2);
        if (split.Length != 2)
        {
            splitChar = '|';
            split = line.Split(splitChar, 2);
        }

        if (split.Length != 2)
        {
            throw new FormatException($"Expected \"question*answer\" (or with '|'): {line}");
        }

        if (split[0] == "Scramble")
        {
            return new TriviaQuestion(defaultCategory, "Scramble: " + Scramble(split[1]), [split[1]]);
        }

        var answersFromLine = split[1].Split(splitChar);
        return new TriviaQuestion(defaultCategory, split[0], answersFromLine);
    }

    private static string Scramble(string word)
    {
        var chars = word.ToCharArray();
        var random = Random.Shared;
        for (var i = chars.Length - 1; i > 0; i--)
        {
            var j = random.Next(i + 1);
            (chars[i], chars[j]) = (chars[j], chars[i]);
        }

        return new string(chars);
    }
}
