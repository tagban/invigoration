using System.Net.Http;
using System.Net.Http.Json;

namespace Invigoration.Core.Config;

/// <summary>
/// Submits opt-in chat-gem tallies (<see cref="ChatGemTallyStore"/>) to the leaderboard endpoint.
///
/// **Inert until <see cref="Endpoint"/> is filled in.** It ships empty on purpose: no endpoint has
/// been chosen yet, and an unconfigured build must not transmit anything. Everything else — the
/// consent gate, the flush triggers, the payload shape — is wired and testable without it.
///
/// What gets sent, and only after the user has explicitly agreed via the first-click prompt: the
/// "username@server" account key, the "yyyy-MM" month, and the cumulative activation count for
/// that month. No IP is collected by this client (whatever host receives the request necessarily
/// sees the connection's source address, as with any HTTP request), no chat content, no other
/// account data.
///
/// Cumulative, never a delta: a lost or failed send costs nothing because the next one carries
/// the full figure, and the receiving end can take max(stored, submitted) per account per month.
/// That same monotonic rule is what makes replayed or forged submissions cheap to clamp — worth
/// remembering that a client-authoritative counter can't be made unforgeable (the app is open
/// source; any signing key would ship with it), so the defence is server-side sanity caps and
/// the fact that every submission carries an account name that can simply be zeroed out.
/// </summary>
public sealed class ChatGemTallySender(HttpClient? httpClient = null)
{
    /// <summary>The leaderboard endpoint. Empty = feature disabled, nothing is ever sent.</summary>
    public const string Endpoint = "";

    /// <summary>Kept short so a flush triggered by app close or disconnect can never noticeably delay either.</summary>
    private static readonly TimeSpan RequestTimeout = TimeSpan.FromSeconds(4);

    private readonly HttpClient _httpClient = httpClient ?? new HttpClient { Timeout = RequestTimeout };

    public static bool IsConfigured => !string.IsNullOrWhiteSpace(Endpoint);

    /// <summary>True only when an endpoint exists AND the user opted in — both must hold before anything leaves the machine.</summary>
    public static bool CanSubmit => IsConfigured && ChatGemTallyStore.SharingEnabled;

    /// <summary>
    /// Sends every account whose count is ahead of what's been acknowledged. Safe to call on any
    /// trigger (idle, periodic, disconnect, app close) and safe to call when there's nothing
    /// pending — it no-ops. Never throws: a leaderboard submission failing is not a reason to
    /// disrupt a chat client, and the un-acknowledged count simply goes out with the next flush.
    /// </summary>
    public async Task FlushAsync(CancellationToken cancellationToken = default)
    {
        if (!CanSubmit)
        {
            return;
        }

        foreach (var (accountKey, tally) in ChatGemTallyStore.PendingSubmissions())
        {
            try
            {
                var payload = new ChatGemSubmission(accountKey, tally.Month, tally.Count);
                using var response = await _httpClient
                    .PostAsJsonAsync(Endpoint, payload, cancellationToken)
                    .ConfigureAwait(false);

                if (response.IsSuccessStatusCode)
                {
                    ChatGemTallyStore.MarkSubmitted(accountKey, tally.Count);
                }
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or OperationCanceledException)
            {
                // Endpoint down, DNS failure, timeout, or shutdown cancelling us mid-flight —
                // leave Submitted where it is so the next flush retries with the same cumulative
                // figure. Deliberately not logged to the chat window: this is a background
                // cosmetic feature and shouldn't put noise in front of the user.
                return;
            }
        }
    }
}

/// <summary>The submission body: account identity, the month it belongs to, and that month's cumulative activation count.</summary>
public sealed record ChatGemSubmission(string Account, string Month, int Count);
