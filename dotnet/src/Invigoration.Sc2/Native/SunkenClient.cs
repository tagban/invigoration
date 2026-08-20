using System.Net;
using System.Net.Sockets;
using Invigoration.Sc2.Front;
using Invigoration.Sc2.Wire;

namespace Invigoration.Sc2.Native;

/// <summary>An established, encrypted Sunken session, ready for chat-service bootstrap (toon select, channel join).</summary>
public sealed record SunkenSession(RecordStream Stream, ResumeResult.Success Details);

/// <summary>
/// Opens the Sunken TCP connection and drives Resume through to an
/// encrypted <see cref="RecordStream"/>. Ported from
/// core/src/native/client.rs's Connector::authenticate_with_handoff,
/// including its exact record sequence and the "skip one optional
/// RegulatorUpdate before the real ResumeResponse" behavior.
///
/// Like <see cref="Front.FrontClient"/>, this has NOT been exercised against
/// a live Battle.net server.
/// </summary>
public static class SunkenClient
{
    private const byte ConnectionBoomCommand = 1;
    private const byte ConnectionRegulatorCommand = 11;
    private const byte AuthConfigurationCommand = 18;
    private const byte AuthProofCommand = 2;
    private const byte AuthResumeCommand = 1;

    public static async Task<SunkenSession> ConnectAsync(SunkenHandoff handoff, int defaultPort = 1119, Action<string>? onStage = null, CancellationToken cancellationToken = default)
    {
        if (handoff.LogonResponse is not { Length: > 0 } logonResponseBlob)
        {
            throw new InvalidOperationException("Sunken handoff has no logon_response payload — cannot determine the game-account region for Resume.");
        }

        // ResumeRequest.game_account_region must come from LogonResponse3's own
        // m_gameAccountRegion field, not the top-level GameUtilities "account_region"
        // attribute (a different field, used only for a consistency check upstream).
        // Confirmed from core/src/native/protocol.rs's decode_logon_parameters — using
        // the wrong source here causes the server to silently close the connection.
        var logonParameters = LogonResponse3Decoder.Decode(logonResponseBlob);
        onStage?.Invoke($"LogonResponse3: account_region={logonParameters.AccountRegion} game_account_region={logonParameters.GameAccountRegion} game_account_name='{logonParameters.GameAccountName}'");
        if (logonParameters.AccountRegion != handoff.AccountRegion)
        {
            throw new InvalidOperationException($"Front logon account region is inconsistent between GameUtilities ({handoff.AccountRegion}) and LogonResponse3 ({logonParameters.AccountRegion}).");
        }

        if (logonParameters.GameAccountName != handoff.GameAccountName)
        {
            throw new InvalidOperationException($"Front logon game-account name is inconsistent between GameUtilities ('{handoff.GameAccountName}') and LogonResponse3 ('{logonParameters.GameAccountName}').");
        }

        var (host, port) = handoff.Endpoint(defaultPort);
        onStage?.Invoke($"Opening TCP connection to {host}:{port}...");
        var tcpClient = new TcpClient { NoDelay = true };
        await tcpClient.ConnectAsync(host, port, cancellationToken).ConfigureAwait(false);
        onStage?.Invoke("TCP connected.");

        var peerAddress = ((IPEndPoint)tcpClient.Client.RemoteEndPoint!).Address;
        var thumbprintContext = NativeCrypto.ThumbprintContextForPeer(peerAddress.ToString());

        var stream = new RecordStream(tcpClient.GetStream());

        onStage?.Invoke("Sending Auth/1 ResumeRequest...");
        await stream.SendAsync(
            ResumeHandshake.EncodeResumeRequest(handoff.AccountMail, logonParameters.GameAccountRegion, handoff.GameAccountName),
            cancellationToken).ConfigureAwait(false);

        onStage?.Invoke("Waiting for Auth/18 Configuration...");
        await ReceiveRecordAsync(stream, (command, slot, reader) =>
        {
            RequireRoute(command, slot, ResumeHandshake.AuthenticationSlot, AuthConfigurationCommand);
            return ResumeHandshake.DecodeConfiguration(reader);
        }, cancellationToken).ConfigureAwait(false);
        onStage?.Invoke("Got Configuration.");

        onStage?.Invoke("Waiting for Auth/2 ProofRequest...");
        var proofRequestModules = await ReceiveRecordAsync(stream, (command, slot, reader) =>
        {
            RequireRoute(command, slot, ResumeHandshake.AuthenticationSlot, AuthProofCommand);
            return ResumeHandshake.DecodeProofRequest(reader);
        }, cancellationToken).ConfigureAwait(false);
        onStage?.Invoke($"Got ProofRequest ({proofRequestModules.Count} modules).");

        var answer = ResumeHandshake.AnswerProofRequest(proofRequestModules, handoff.SessionKey, thumbprintContext);
        onStage?.Invoke("Thumbprint verified. Sending Auth/2 ProofResponse...");
        await stream.SendAsync(ResumeHandshake.EncodeProofResponse(answer.SessionProofOutput), cancellationToken).ConfigureAwait(false);

        onStage?.Invoke("Waiting for Auth/1 ResumeResponse...");
        ResumeResult? resumeResult = null;
        while (resumeResult is null)
        {
            resumeResult = await ReceiveRecordAsync(stream, (command, slot, reader) =>
            {
                if (slot == ResumeHandshake.ConnectionSlot && command == ConnectionRegulatorCommand)
                {
                    ResumeHandshake.DecodeRegulatorUpdate(reader);
                    return (ResumeResult?)null;
                }

                RequireRoute(command, slot, ResumeHandshake.AuthenticationSlot, AuthResumeCommand);
                return ResumeHandshake.DecodeResumeResponse(reader);
            }, cancellationToken).ConfigureAwait(false);
        }

        if (resumeResult is ResumeResult.Failure failure)
        {
            throw failure.Reason switch
            {
                ResumeFailureReason.Failed f => new NativeResumeRejectedException(f.ErrorCode, f.WaitSeconds),
                ResumeFailureReason.VersionCheckDisconnect => new InvalidOperationException("Sunken rejected Resume: client version is out of date."),
                _ => new InvalidOperationException("Sunken rejected Resume: an update is required."),
            };
        }

        onStage?.Invoke("Got ResumeResponse (success). Validating server proof...");
        var success = (ResumeResult.Success)resumeResult;
        ResumeHandshake.ValidateServerProof(success, answer.Session);
        onStage?.Invoke("Server proof verified. Enabling encryption...");

        await stream.SendAsync(ResumeHandshake.EncodeEnableEncryption(), cancellationToken).ConfigureAwait(false);
        var (inboundKey, outboundKey) = NativeCrypto.DeriveTransportRc4Keys(answer.Session.TransportKey);
        stream.EnableEncryption(new Rc4State(inboundKey), new Rc4State(outboundKey));

        return new SunkenSession(stream, success);
    }

    /// <summary>Polls the stream for the next complete record, checking every one for a Connection/Boom termination before handing it to <paramref name="decode"/>.</summary>
    private static async Task<T> ReceiveRecordAsync<T>(RecordStream stream, Func<byte, byte?, BitReader, T> decode, CancellationToken cancellationToken)
    {
        while (true)
        {
            var completed = stream.TryDecodeRecord(
                (command, slot, reader) =>
                {
                    if (slot == ResumeHandshake.ConnectionSlot && command == ConnectionBoomCommand)
                    {
                        throw new NativeServerRejectedException((ushort)reader.Read(16));
                    }

                    return decode(command, slot, reader);
                },
                out var result);

            if (completed)
            {
                return result!;
            }

            if (!await stream.FillAsync(cancellationToken).ConfigureAwait(false))
            {
                throw new IOException("Sunken connection closed before a complete record arrived.");
            }
        }
    }

    private static void RequireRoute(byte command, byte? slot, byte expectedSlot, byte expectedCommand)
    {
        if (slot != expectedSlot || command != expectedCommand)
        {
            throw new InvalidOperationException($"Unexpected native record route slot={slot} command={command} (expected slot={expectedSlot} command={expectedCommand}).");
        }
    }
}
