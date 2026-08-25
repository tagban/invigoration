using System.Reflection;
using System.Runtime.InteropServices;

namespace Invigoration.Core;

/// <summary>
/// Registers an explicit native-library resolver for Stimpak's P/Invoke calls, bypassing the
/// OS's own shared-library search order entirely — see StimpakPackage.props for the full story
/// on why that's needed. Short version: Stimpak's managed assembly is named "Stimpak.dll" and
/// its native library "stimpak.dll" — same name apart from casing — and even when the native
/// library is staged correctly (nested under runtimes/&lt;rid&gt;/native/, never colliding with
/// the managed assembly on disk), the OS's *default* search order still checks the flat app
/// directory — where the managed assembly sits — before it ever probes that nested folder, so a
/// bare P/Invoke call finds the wrong file and fails with EntryPointNotFoundException. This
/// sidesteps that ambiguity by telling .NET exactly which file to load, unconditionally.
///
/// Must run once, before the first StimpakClient is constructed anywhere in the process — see
/// BotEngine.Sc2.cs's ConnectSc2Async, the only call site that needs it.
/// </summary>
public static class StimpakNativeResolver
{
    private static bool _registered;

    public static void Register()
    {
        if (_registered)
        {
            return;
        }

        _registered = true;
        NativeLibrary.SetDllImportResolver(typeof(Stimpak.StimpakClient).Assembly, Resolve);
    }

    /// <summary>Returning IntPtr.Zero (for any name other than Stimpak's own "stimpak") is not a failure — it's the documented signal for .NET to fall back to its own default resolution for that call.</summary>
    private static IntPtr Resolve(string libraryName, Assembly assembly, DllImportSearchPath? searchPath)
    {
        if (libraryName != "stimpak")
        {
            return IntPtr.Zero;
        }

        var fileName = OperatingSystem.IsWindows() ? "stimpak.dll"
            : OperatingSystem.IsMacOS() ? "libstimpak.dylib"
            : "libstimpak.so";

        var path = Path.Combine(AppContext.BaseDirectory, "runtimes", RuntimeInformation.RuntimeIdentifier, "native", fileName);
        return NativeLibrary.Load(path);
    }
}
