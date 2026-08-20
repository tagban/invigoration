namespace Invigoration.Sc2.Native;

/// <summary>The server sent a Connection/Boom record — an explicit native-protocol error terminating the session. Mirrors core/src/native/auth.rs's native_server_error / Error::NativeServerRejected.</summary>
public sealed class NativeServerRejectedException(ushort errorCode)
    : Exception($"Native Battle.net server rejected the connection (error code {errorCode}).")
{
    public ushort ErrorCode { get; } = errorCode;
}

/// <summary>Sunken rejected the Resume attempt with a structured error+wait pair. Mirrors Error::NativeResumeRejected.</summary>
public sealed class NativeResumeRejectedException(ushort errorCode, int waitSeconds)
    : Exception($"Native Battle.net resume was rejected (error code {errorCode}, retry after {waitSeconds}s).")
{
    public ushort ErrorCode { get; } = errorCode;
    public int WaitSeconds { get; } = waitSeconds;
}
