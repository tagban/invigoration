namespace Invigoration.Core.Config;

/// <summary>
/// Named, swappable collections of icon overrides, stored as subfolders under
/// %AppData%/Invigoration/IconSets/&lt;name&gt;/&lt;key&gt;.&lt;ext&gt; — each set
/// mirrors <see cref="IconOverrideStore"/>'s flat key-to-file layout, so
/// saving a set is just copying the current override files out, applying one
/// is copying them back in, and a set is just an ordinary folder a user can
/// zip or copy elsewhere to back it up.
/// </summary>
public static class IconSetStore
{
    public static string Directory => Path.Combine(ConfigStore.DefaultConfigDirectory(), "IconSets");

    /// <summary>Raised whenever a set is saved or deleted, so a UI can refresh its list.</summary>
    public static event Action? SetsChanged;

    public static IReadOnlyList<string> ListSets()
    {
        if (!System.IO.Directory.Exists(Directory))
        {
            return [];
        }

        return System.IO.Directory.GetDirectories(Directory)
            .Select(Path.GetFileName)
            .Where(name => !string.IsNullOrEmpty(name))
            .Select(name => name!)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>Snapshots every file currently applied as an override into a new (or replaced) named set.</summary>
    public static void SaveCurrentAsSet(string name)
    {
        var target = Path.Combine(Directory, SanitizeName(name));
        if (System.IO.Directory.Exists(target))
        {
            System.IO.Directory.Delete(target, recursive: true);
        }

        System.IO.Directory.CreateDirectory(target);

        if (System.IO.Directory.Exists(IconOverrideStore.Directory))
        {
            foreach (var file in System.IO.Directory.GetFiles(IconOverrideStore.Directory))
            {
                File.Copy(file, Path.Combine(target, Path.GetFileName(file)), overwrite: true);
            }
        }

        SetsChanged?.Invoke();
    }

    /// <summary>Applies a saved set: clears every current override, then copies the set's files in as the new overrides.</summary>
    public static void ApplySet(string name)
    {
        var source = Path.Combine(Directory, SanitizeName(name));
        if (!System.IO.Directory.Exists(source))
        {
            return;
        }

        if (System.IO.Directory.Exists(IconOverrideStore.Directory))
        {
            foreach (var file in System.IO.Directory.GetFiles(IconOverrideStore.Directory))
            {
                File.Delete(file);
                IconOverrideStore.NotifyOverrideChanged(Path.GetFileNameWithoutExtension(file));
            }
        }

        System.IO.Directory.CreateDirectory(IconOverrideStore.Directory);
        foreach (var file in System.IO.Directory.GetFiles(source))
        {
            var fileName = Path.GetFileName(file);
            File.Copy(file, Path.Combine(IconOverrideStore.Directory, fileName), overwrite: true);
            IconOverrideStore.NotifyOverrideChanged(Path.GetFileNameWithoutExtension(fileName));
        }
    }

    public static void DeleteSet(string name)
    {
        var target = Path.Combine(Directory, SanitizeName(name));
        if (System.IO.Directory.Exists(target))
        {
            System.IO.Directory.Delete(target, recursive: true);
        }

        SetsChanged?.Invoke();
    }

    private static string SanitizeName(string name)
    {
        name = name.Trim();
        foreach (var invalid in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalid, '_');
        }

        return name;
    }
}
