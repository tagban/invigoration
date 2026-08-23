using Invigoration.Core.Trivia;

namespace Invigoration.Core.Tests;

public class TriviaQuestionParseTests
{
    [Fact]
    public void Parse_AsteriskFormat_SingleAnswer()
    {
        var q = TriviaQuestion.Parse("What year was Diablo released?*1996", "Trivia");

        Assert.Equal("Trivia", q.Category);
        Assert.Equal("What year was Diablo released?", q.QuestionText);
        Assert.Equal(["1996"], q.Answers);
    }

    [Fact]
    public void Parse_AsteriskFormat_MultipleAnswers()
    {
        var q = TriviaQuestion.Parse("Name a Prime Evil.*Diablo*Mephisto*Baal", "Diablo");

        Assert.Equal(["Diablo", "Mephisto", "Baal"], q.Answers);
    }

    [Fact]
    public void Parse_PipeFallback_WhenLineHasNoAsterisk()
    {
        var q = TriviaQuestion.Parse("What comes after 9?|10", "Trivia");

        Assert.Equal("What comes after 9?", q.QuestionText);
        Assert.Equal(["10"], q.Answers);
    }

    [Fact]
    public void Parse_SlashFormat_ExplicitCategoryAndMultipleAnswers()
    {
        var q = TriviaQuestion.Parse("/Diablo/Diablo/Mephisto/Baal//Name a Prime Evil.", "ignored");

        Assert.Equal("Diablo", q.Category);
        Assert.Equal("Name a Prime Evil.", q.QuestionText);
        Assert.Equal(["Diablo", "Mephisto", "Baal"], q.Answers);
    }

    [Fact]
    public void Parse_Scramble_ProducesShuffledQuestionWithOriginalAsAnswer()
    {
        var q = TriviaQuestion.Parse("Scramble*Tristram", "Trivia");

        Assert.StartsWith("Scramble: ", q.QuestionText);
        Assert.Equal(["Tristram"], q.Answers);
    }

    [Fact]
    public void Parse_MalformedLine_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() => TriviaQuestion.Parse("no delimiter here", "Trivia"));
    }

    [Fact]
    public void Parse_EmptyLine_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() => TriviaQuestion.Parse("", "Trivia"));
    }

    [Fact]
    public void Parse_MultipleChoice_FourChoices()
    {
        var q = TriviaQuestion.Parse("MC*Blizzard*Which game has the Secret Cow Level?*Diablo II*StarCraft*Warcraft III*Hearthstone", "ignored");

        Assert.True(q.IsMultipleChoice);
        Assert.Equal("Blizzard", q.Category);
        Assert.Equal("Which game has the Secret Cow Level?", q.QuestionText);
        Assert.Equal(4, q.Choices.Count);
        Assert.Contains(q.Choices, c => c.Text == "Diablo II");
        Assert.Contains(q.Choices, c => c.Text == "StarCraft");
        Assert.Contains(q.Choices, c => c.Text == "Warcraft III");
        Assert.Contains(q.Choices, c => c.Text == "Hearthstone");
        // Letters are exactly A/B/C/D for a 4-choice question, assigned by shuffled position.
        Assert.Equal(['A', 'B', 'C', 'D'], q.Choices.Select(c => c.Letter).OrderBy(l => l));
    }

    [Fact]
    public void Parse_MultipleChoice_MinimumTwoChoices()
    {
        var q = TriviaQuestion.Parse("MC*Diablo*What class of enemy is Diablo?*Demon*Undead", "ignored");

        Assert.True(q.IsMultipleChoice);
        Assert.Equal(2, q.Choices.Count);
        Assert.Equal(['A', 'B'], q.Choices.Select(c => c.Letter).OrderBy(l => l));
    }

    [Fact]
    public void Parse_MultipleChoice_TooFewFields_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() => TriviaQuestion.Parse("MC*Category*Question*OnlyCorrectAnswer", "ignored"));
    }

    [Fact]
    public void Parse_MultipleChoice_TooManyChoices_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() =>
            TriviaQuestion.Parse("MC*Category*Question*Correct*Wrong1*Wrong2*Wrong3*Wrong4", "ignored"));
    }
}

public class TriviaQuestionMatchTests
{
    [Theory]
    [InlineData("Tyrael", true)]
    [InlineData("tyrael", true)]
    [InlineData("Tyrael!", true)]
    [InlineData("  Tyrael  ", true)]
    [InlineData("it's Tyrael!", false)]
    [InlineData("Diablo", false)]
    [InlineData("", false)]
    public void TryMatchAnswer_NormalizesCaseAndPunctuation(string text, bool expectedMatch)
    {
        // Matching is whole-message, not substring — an extra word like "it's" still misses,
        // matching BNU`Bot's own full-string equalsIgnoreCase comparison after normalization.
        var q = TriviaQuestion.Parse("Who is the archangel of Justice?*Tyrael", "Diablo");

        var matched = q.TryMatchAnswer(text, out var answer);

        Assert.Equal(expectedMatch, matched);
        if (expectedMatch)
        {
            Assert.Equal("Tyrael", answer);
        }
    }

    [Fact]
    public void TryMatchAnswer_MatchesAnyAcceptedAnswer()
    {
        var q = TriviaQuestion.Parse("Name a Prime Evil.*Diablo*Mephisto*Baal", "Diablo");

        Assert.True(q.TryMatchAnswer("mephisto", out var answer));
        Assert.Equal("Mephisto", answer);
    }

    [Fact]
    public void Hints_RevealProgressivelyMoreOfPrimaryAnswer()
    {
        var q = TriviaQuestion.Parse("Test question*Tristram", "Trivia");

        var hidden0 = q.Hint0.Count(c => c == '?');
        var hidden1 = q.Hint1.Count(c => c == '?');
        var hidden2 = q.Hint2.Count(c => c == '?');

        Assert.True(hidden0 >= hidden1);
        Assert.True(hidden1 >= hidden2);
        Assert.Equal(q.Hint0.Length, q.Hint1.Length);
        Assert.Equal(q.Hint1.Length, q.Hint2.Length);
    }

    [Fact]
    public void TryMatchAnswer_MultipleChoice_AcceptsEitherLetterOrFullText()
    {
        var q = TriviaQuestion.Parse("MC*Diablo*What class of enemy is Diablo?*Demon*Undead", "ignored");
        var correctLetter = q.Choices.First(c => c.Text == "Demon").Letter;

        Assert.True(q.TryMatchAnswer("Demon", out var byText));
        Assert.Equal("Demon", byText);

        Assert.True(q.TryMatchAnswer(correctLetter.ToString(), out var byLetter));
        Assert.Equal(correctLetter.ToString(), byLetter);

        Assert.False(q.TryMatchAnswer("Undead", out _)); // the wrong option's own text should never match
    }

    [Fact]
    public void CreateMultipleChoice_CorrectAnswerIsNotAlwaysTheFirstLetter()
    {
        // Shuffling is random, so run it enough times that landing on "A"
        // every single time would be implausible (1-in-4^N) if the shuffle
        // were actually broken/no-op.
        var everAssignedNonA = false;
        for (var i = 0; i < 30; i++)
        {
            var q = TriviaQuestion.CreateMultipleChoice("cat", "q", "Correct", ["Wrong1", "Wrong2", "Wrong3"]);
            var correctLetter = q.Choices.First(c => c.Text == "Correct").Letter;
            if (correctLetter != 'A')
            {
                everAssignedNonA = true;
                break;
            }
        }

        Assert.True(everAssignedNonA, "Correct answer landed on 'A' in all 30 attempts — shuffle looks broken.");
    }
}

public class TriviaSessionTests
{
    private static TriviaQuestion Question(string q = "Q*A") => TriviaQuestion.Parse(q, "Trivia");

    [Fact]
    public void Start_MakesSessionEnabledWithFullPool()
    {
        var session = new TriviaSession();

        session.Start([Question(), Question("Q2*B")]);

        Assert.True(session.IsEnabled);
        Assert.Equal(2, session.QuestionsRemaining);
    }

    [Fact]
    public void AskNext_DrainsPoolAndReturnsNullWhenEmpty()
    {
        var session = new TriviaSession();
        session.Start([Question()]);

        var first = session.AskNext();
        var second = session.AskNext();

        Assert.NotNull(first);
        Assert.Null(second);
        Assert.Equal(0, session.QuestionsRemaining);
    }

    [Fact]
    public void TryMatchAnswer_NoActiveQuestion_NeverMatches()
    {
        var session = new TriviaSession();
        session.Start([Question("What?*Answer")]);
        // Deliberately not calling AskNext() — no Current question yet.

        var matched = session.TryMatchAnswer("Answer", out _);

        Assert.False(matched);
    }

    [Fact]
    public void TryMatchAnswer_AnyoneCanAnswer_NoJoinStepRequired()
    {
        var session = new TriviaSession();
        session.Start([Question("What?*Answer")]);
        session.AskNext();

        var matched = session.TryMatchAnswer("answer", out var text);

        Assert.True(matched);
        Assert.Equal("Answer", text);
    }

    [Fact]
    public void RecordAnswered_ResetsUnansweredStreak()
    {
        var session = new TriviaSession();
        session.Start([Question()]);

        session.RecordTimeout();
        session.RecordTimeout();
        session.RecordAnswered();

        Assert.Equal(0, session.UnansweredStreak);
    }

    [Fact]
    public void Stop_ClearsEnabledAndCurrentQuestion()
    {
        var session = new TriviaSession();
        session.Start([Question()]);
        session.AskNext();

        session.Stop();

        Assert.False(session.IsEnabled);
        Assert.Null(session.Current);
    }
}
