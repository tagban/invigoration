namespace Invigoration.Core.Trivia;

/// <summary>
/// Question source: every ".txt" file dropped in the Trivia folder inside
/// the app's config directory (reachable via "Open Config Folder"), normally
/// seeded on first run by <see cref="TriviaPackDownloader"/> downloading the
/// base packs from GitHub so they land as normal, freely editable files
/// rather than something baked into the app. If the folder has no ".txt"
/// files at all yet — the download hasn't run, is still in flight, or
/// failed and the user is offline — a small pack embedded directly in code
/// (<see cref="BundledEntries"/>) is used instead, purely so trivia always
/// has *something* to play rather than nothing. Once any ".txt" file
/// exists, the embedded pack is not used at all (avoids duplicate
/// questions once the downloaded/edited files are the real source). Custom
/// files use the same line format as BNU`Bot's trivia packs, so an existing
/// one can be copied in unmodified.
/// </summary>
public static class TriviaBank
{
    public static string Directory => Path.Combine(Config.ConfigStore.DefaultConfigDirectory(), "Trivia");

    /// <summary>Loads every ".txt" file under the Trivia folder, or the embedded fallback pack if there are none yet. A line that fails to parse is skipped (reported via onParseError) rather than aborting the whole file.</summary>
    public static List<TriviaQuestion> LoadAll(Action<string>? onParseError = null)
    {
        var questions = new List<TriviaQuestion>();

        // Created (not just checked) so the folder shows up under Open Config
        // Folder right away, even before any file has landed in it.
        System.IO.Directory.CreateDirectory(Directory);
        var files = System.IO.Directory.GetFiles(Directory, "*.txt");

        if (files.Length == 0)
        {
            foreach (var (category, line) in BundledEntries)
            {
                TryAdd(questions, line, category, onParseError, source: "(bundled fallback)");
            }

            return questions;
        }

        foreach (var file in files)
        {
            var defaultCategory = Path.GetFileNameWithoutExtension(file);
            foreach (var rawLine in File.ReadAllLines(file))
            {
                var line = rawLine.Trim();
                if (line.Length == 0)
                {
                    continue;
                }

                TryAdd(questions, line, defaultCategory, onParseError, source: Path.GetFileName(file));
            }
        }

        return questions;
    }

    private static void TryAdd(List<TriviaQuestion> questions, string line, string category, Action<string>? onParseError, string source)
    {
        try
        {
            questions.Add(TriviaQuestion.Parse(line, category));
        }
        catch (FormatException ex)
        {
            onParseError?.Invoke($"{source}: {ex.Message}");
        }
    }

    /// <summary>
    /// Starter pack: general well-known facts about the games (release
    /// years, character/faction names) rather than reproduced game text.
    /// "*"-delimited: "question*answer1*answer2...". Category names double
    /// as what "!trivia &lt;category&gt;" matches against (case-insensitive),
    /// so keep them short and consistent (see BotEngine.Trivia.cs).
    /// </summary>
    private static readonly (string Category, string Line)[] BundledEntries =
    [
        ("Diablo", "What year was the original Diablo released?*1996"),
        ("Diablo", "What is the name of the hero's hometown in Diablo I?*Tristram"),
        ("Diablo", "Diablo is also known as the Lord of what?*Terror"),
        ("Diablo", "What is the name of Diablo II's expansion?*Lord of Destruction"),
        ("Diablo", "Which Diablo II class specializes in bows and javelins?*Amazon"),
        ("Diablo", "What is the name of the archangel of Justice, a recurring character across the Diablo series?*Tyrael"),
        ("Diablo", "In Diablo II, what is the capital city of Act II, set in the desert?*Lut Gholein"),
        ("Diablo", "Name one of the three Prime Evils.*Diablo*Mephisto*Baal"),
        ("Diablo", "How many difficulty levels does Diablo II have?*3*Three"),
        ("Diablo", "What is the name of the one-legged boy merchant found in Diablo II's Rogue Encampment?*Wirt"),
        ("Diablo", "What year was Diablo III released?*2012"),
        ("Diablo", "What year was Diablo IV released?*2023"),
        ("Diablo", "Name one of the five original Diablo III character classes.*Barbarian*Witch Doctor*Wizard*Monk*Demon Hunter"),
        ("Diablo", "In Diablo III's opening, who is the mysterious figure that arrives at New Tristram, later revealed to be Diablo's host?*The Dark Wanderer*Dark Wanderer"),
        ("Warcraft", "What year was the original Warcraft: Orcs & Humans released?*1994"),
        ("Warcraft", "What year was Warcraft III: Reign of Chaos released?*2002"),
        ("Warcraft", "What is the name of Warcraft III's expansion?*The Frozen Throne"),
        ("Warcraft", "What race is Thrall, the famous Warcraft warchief?*Orc"),
        ("Warcraft", "Who is the fallen paladin prince who becomes the Lich King's champion in Warcraft III?*Arthas"),
        ("Warcraft", "Name one of the two playable factions in Warcraft II: Tides of Darkness.*Humans*Orcs"),
        ("Warcraft", "Who is the archdruid and central Night Elf hero of the Warcraft III campaign, brother of Illidan?*Malfurion*Malfurion Stormrage"),
        ("Warcraft", "What is the undead faction commonly called in Warcraft III?*Scourge*Undead"),
        ("Warcraft", "In Warcraft III, where do you recruit neutral heroes to join your side?*Tavern*Taverns"),
        ("Warcraft", "What race was Sylvanas Windrunner before she became undead?*High Elf*Elf"),
        ("Warcraft", "What year was World of Warcraft: The Burning Crusade released?*2007"),
        ("Warcraft", "What is World of Warcraft's second expansion, centered on the Lich King, called?*Wrath of the Lich King"),
        ("Warcraft", "What is the capital city of the Horde in World of Warcraft?*Orgrimmar"),
        ("Warcraft", "What is the capital city of the human Alliance kingdom, featured since Warcraft III?*Stormwind"),
        ("Warcraft", "Which orc spirit does Arthas merge with atop Icecrown Glacier to become the Lich King?*Ner'zhul*Nerzhul"),
        ("Warcraft", "Which dreadlord corrupts Arthas Menethil early in Warcraft III's human campaign?*Mal'Ganis*Malganis"),
        ("StarCraft", "What year was the original StarCraft released?*1998"),
        ("StarCraft", "What is the name of StarCraft's expansion pack?*Brood War"),
        ("StarCraft", "Name one of the three playable races in StarCraft.*Terran*Protoss*Zerg"),
        ("StarCraft", "What is the name of the Zerg's original hive-mind entity?*The Overmind*Overmind"),
        ("StarCraft", "What title does Sarah Kerrigan take after being infested by the Zerg?*Queen of Blades"),
        ("StarCraft", "What is the Protoss home planet called?*Aiur"),
        ("StarCraft", "What year was StarCraft II: Wings of Liberty released?*2010"),
        ("StarCraft", "Name the Terran resistance leader and central protagonist of the StarCraft saga, a former marshal.*Jim Raynor*Raynor"),
        ("StarCraft", "What weapon do Protoss Zealots wield?*Psi Blades*Psionic Blades"),
        ("StarCraft", "Besides minerals, what other resource do you gather in StarCraft?*Vespene Gas*Gas"),
        ("StarCraft", "What year was StarCraft: Brood War released?*1998"),
        ("StarCraft", "What year was StarCraft II: Heart of the Swarm released?*2013"),
        ("StarCraft", "What year was StarCraft II: Legacy of the Void released?*2015"),
        ("StarCraft", "Who is the Terran Dominion emperor and primary antagonist of StarCraft II?*Arcturus Mengsk*Mengsk"),
        ("StarCraft", "What is the name of the Dark Templar's hidden homeworld?*Shakuras"),
        ("Blizzard", "What was Blizzard Entertainment's original studio name when it was founded in 1991?*Silicon & Synapse*Silicon and Synapse"),
        ("Blizzard", "What company acquired Silicon & Synapse (Blizzard's original name) in 1994?*Davidson & Associates*Davidson and Associates"),
        ("Blizzard", "What is the name of Blizzard's classic puzzle-platformer starring three Vikings stranded across time?*The Lost Vikings*Lost Vikings"),
        ("Blizzard", "Name one of the three original playable Vikings in The Lost Vikings.*Erik*Erik the Swift*Baleog*Baleog the Fierce*Olaf*Olaf the Stout"),
        ("Blizzard", "What is the name of Blizzard's 1993 combat racing game?*Rock n' Roll Racing*Rock and Roll Racing"),
        ("Blizzard", "What is the name of Blizzard's 1994 side-scrolling action-platformer game?*Blackthorne"),
        ("Blizzard", "What year did Battle.net launch alongside the original Diablo?*1996"),
        ("Blizzard", "In Diablo II, what item combination famously opens the Secret Cow Level?*Wirt's Leg*Wirt's Leg and a Tome of Town Portal"),
        ("Blizzard", "What is Diablo III's secret non-cow-themed joke level called?*Whimsyshire"),
        ("Blizzard", "What date in 2004 did World of Warcraft officially launch?*November 23*November 23, 2004*November 23rd"),
        ("Blizzard", "What year was Hearthstone released?*2014"),
        ("Blizzard", "What year was Overwatch released?*2016"),
        ("Blizzard", "What is the name of Blizzard's 2015 multiplayer online battle arena featuring heroes from across its franchises?*Heroes of the Storm"),
        ("Blizzard", "What was the codename of Blizzard's cancelled StarCraft stealth-action spin-off?*StarCraft: Ghost*Ghost"),
        ("Blizzard", "Which publisher merged with Blizzard's parent company in 2008 to form Activision Blizzard?*Activision"),
        ("Pop Culture", "Who directed the movie \"Jaws\" (1975)?*Steven Spielberg"),
        ("Pop Culture", "Who played Iron Man in the Marvel Cinematic Universe?*Robert Downey Jr*Robert Downey Jr."),
        ("Pop Culture", "What band performed the song \"Bohemian Rhapsody\"?*Queen"),
        ("Pop Culture", "In \"The Lord of the Rings,\" what region is Frodo's home?*The Shire*Shire"),
        ("Pop Culture", "Who created the Mickey Mouse character?*Walt Disney"),
        ("Pop Culture", "What year did the first iPhone release?*2007"),
        ("Pop Culture", "What is the name of the wizarding school in the Harry Potter series?*Hogwarts"),
        ("Pop Culture", "Who wrote the \"A Song of Ice and Fire\" novels that inspired \"Game of Thrones\"?*George R.R. Martin*George RR Martin*GRRM"),
        ("Pop Culture", "What is the fictional African nation in Marvel's \"Black Panther\"?*Wakanda"),
        ("Pop Culture", "Who played the Joker in 2008's \"The Dark Knight\"?*Heath Ledger"),
        ("Music", "What British band released the album \"Abbey Road\" in 1969?*The Beatles*Beatles"),
        ("Music", "Who is known as the \"King of Pop\"?*Michael Jackson"),
        ("Music", "What is the name of Beethoven's famous symphony known for its four-note opening motif?*Symphony No. 5*Fifth Symphony*5th Symphony"),
        ("Music", "Which artist released the album \"Thriller\" in 1982?*Michael Jackson"),
        ("Music", "What instrument is Jimi Hendrix best known for playing?*Guitar*Electric Guitar"),
        ("Music", "What Swedish pop group released \"Dancing Queen\"?*ABBA"),
        ("Music", "Who composed the opera \"The Magic Flute\"?*Wolfgang Amadeus Mozart*Mozart"),
        ("Music", "Freddie Mercury was the lead singer of which band?*Queen"),
        ("Music", "What genre of music is Nashville, Tennessee best known for?*Country*Country Music"),
        ("Music", "Who sang \"Like a Rolling Stone\"?*Bob Dylan"),
    ];
}
