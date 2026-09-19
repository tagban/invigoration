using Invigoration.Core.Config;

namespace Invigoration.Core.Tests;

/// <summary>
/// Which icon set is showing, remembered so a bot with no set of its own can tick the one it's
/// getting, and so switching to a bot whose set is already showing copies nothing.
/// </summary>
public class IconSetStoreTests
{
    [Fact]
    public void ApplyingASavedSet_RecordsItAsShowing_AndSaysSo()
    {
        var name = $"set-{Guid.NewGuid():N}";
        IconSetStore.SaveCurrentAsSet(name);
        var changes = 0;
        void Count() => changes++;
        IconSetStore.ActiveSetChanged += Count;

        try
        {
            IconSetStore.ApplySet(name);

            Assert.Equal(name, IconSetStore.ActiveSetName);
            Assert.Equal(1, changes);

            // The same set again changes nothing, so nothing needs to redraw.
            IconSetStore.ActiveSetName = name;
            Assert.Equal(1, changes);
        }
        finally
        {
            IconSetStore.ActiveSetChanged -= Count;
            IconSetStore.DeleteSet(name);
        }
    }

    [Fact]
    public void TheSetShowing_IsKeptOnDisk()
    {
        IconSetStore.ActiveSetName = "Battle.net 1.0 Classic";

        Assert.Equal("Battle.net 1.0 Classic", File.ReadAllText(Path.Combine(ConfigStore.DefaultConfigDirectory(), "icon-set.txt")));
        Assert.Equal("Battle.net 1.0 Classic", IconSetStore.ActiveSetName);
    }
}
