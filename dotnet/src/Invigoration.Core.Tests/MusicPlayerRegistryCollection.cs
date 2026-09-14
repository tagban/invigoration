namespace Invigoration.Core.Tests;

/// <summary>MusicPlayerRegistry.Controller is one process-wide slot: test classes that set it run one at a time, so one can't swap in its controller halfway through another's test.</summary>
[CollectionDefinition(Name)]
public sealed class MusicPlayerRegistryCollection
{
    public const string Name = "MusicPlayerRegistry";
}
