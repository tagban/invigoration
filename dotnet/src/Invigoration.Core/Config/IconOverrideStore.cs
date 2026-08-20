namespace Invigoration.Core.Config;

/// <summary>
/// User-replaceable chat icon files at %AppData%/Invigoration/Icons, keyed
/// by the same icon key <c>Chat.ChatIcon</c>/the App project's GameIconLoader
/// use (e.g. "sc", "mod-gavel"). A file here overrides the bundled default
/// with the same key; clearing it reverts to the default. Mirrors
/// <see cref="ColorSchemeLibrary"/>'s folder-of-files approach.
/// </summary>
public static class IconOverrideStore
{
    private static readonly string[] SupportedExtensions = [".png", ".gif", ".jpg", ".jpeg", ".bmp"];

    public static string Directory => Path.Combine(ConfigStore.DefaultConfigDirectory(), "Icons");

    /// <summary>Raised whenever an override is set or cleared, with the affected key, so a UI can invalidate any cached bitmap for it.</summary>
    public static event Action<string>? OverridesChanged;

    /// <summary>The full path to this key's override file, or null if none exists.</summary>
    public static string? GetOverridePath(string key)
    {
        if (!System.IO.Directory.Exists(Directory))
        {
            return null;
        }

        foreach (var extension in SupportedExtensions)
        {
            var candidate = Path.Combine(Directory, key + extension);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    /// <summary>Copies <paramref name="sourceFilePath"/> in as the override for <paramref name="key"/>, replacing any prior override for it (of any supported extension).</summary>
    public static void SetOverride(string key, string sourceFilePath)
    {
        System.IO.Directory.CreateDirectory(Directory);
        ClearOverride(key);
        var extension = Path.GetExtension(sourceFilePath);
        if (string.IsNullOrEmpty(extension) || !SupportedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            extension = ".png";
        }

        File.Copy(sourceFilePath, Path.Combine(Directory, key + extension), overwrite: true);
        OverridesChanged?.Invoke(key);
    }

    /// <summary>Writes raw image bytes in directly as the override for <paramref name="key"/> — used to apply a bundled alternate icon set without round-tripping through a temp file.</summary>
    public static void SetOverrideBytes(string key, byte[] imageBytes, string extension)
    {
        System.IO.Directory.CreateDirectory(Directory);
        ClearOverride(key);
        if (!extension.StartsWith('.'))
        {
            extension = "." + extension;
        }

        File.WriteAllBytes(Path.Combine(Directory, key + extension), imageBytes);
        OverridesChanged?.Invoke(key);
    }

    /// <summary>Raises OverridesChanged for a key whose override file was written directly rather than through SetOverride/SetOverrideBytes — used by <see cref="IconSetStore"/> after copying a saved set's files in.</summary>
    public static void NotifyOverrideChanged(string key) => OverridesChanged?.Invoke(key);

    public static void ClearOverride(string key)
    {
        var existing = GetOverridePath(key);
        if (existing is not null)
        {
            File.Delete(existing);
            OverridesChanged?.Invoke(key);
        }
    }
}
