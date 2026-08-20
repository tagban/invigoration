namespace Invigoration.Sc2.Native;

/// <summary>Battlenet::Authentication::ModuleId + ModuleData — a 40-byte module identifier plus its opaque payload.</summary>
public sealed record ModuleId(byte[] Usage, byte[] Identity);

public sealed record ModuleInput(ModuleId Id, byte[] Data);

/// <summary>Auth/18 Authentication::Configuration.</summary>
public sealed record ResumeConfiguration(bool UseS3Depot);

/// <summary>Auth/1 Authentication::ResumeResponse's result.</summary>
public abstract record ResumeResult
{
    private ResumeResult()
    {
    }

    public sealed record Success(IReadOnlyList<ModuleInput> FinalRequests, int PingTimeoutSeconds, RegulatorInfo? RegulatorRules) : ResumeResult;

    public sealed record Failure(byte[]? Strings, ResumeFailureReason Reason) : ResumeResult;
}

public abstract record ResumeFailureReason
{
    private ResumeFailureReason()
    {
    }

    public sealed record Update : ResumeFailureReason;

    public sealed record Failed(ushort ErrorCode, int WaitSeconds) : ResumeFailureReason;

    public sealed record VersionCheckDisconnect : ResumeFailureReason;
}

/// <summary>Battlenet::Regulator::Info.</summary>
public abstract record RegulatorInfo
{
    private RegulatorInfo()
    {
    }

    public sealed record None : RegulatorInfo;

    public sealed record LeakyBucket(uint Threshold, uint Rate) : RegulatorInfo;
}
