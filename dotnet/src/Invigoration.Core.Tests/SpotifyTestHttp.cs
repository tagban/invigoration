using System.Net;
using System.Text;

namespace Invigoration.Core.Tests;

/// <summary>A stand-in for Spotify: answers each request from a routing function and records what was asked.</summary>
internal sealed class SpotifyTestHttp : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, string, HttpResponseMessage> _respond;

    public SpotifyTestHttp(Func<HttpRequestMessage, string, HttpResponseMessage> respond) => _respond = respond;

    public List<(HttpMethod Method, string Url, string Body, string? Authorization)> Requests { get; } = [];

    public HttpClient Client() => new(this);

    public static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    public static HttpResponseMessage Status(HttpStatusCode status) => new(status) { Content = new StringContent("") };

    public static HttpResponseMessage Tokens(string access = "access-1", string? refresh = null, int expiresIn = 3600) =>
        Json($$"""{"access_token":"{{access}}","token_type":"Bearer","expires_in":{{expiresIn}}{{(refresh is null ? "" : $",\"refresh_token\":\"{refresh}\"")}}}""");

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var body = request.Content is null ? "" : await request.Content.ReadAsStringAsync(cancellationToken);
        lock (Requests)
        {
            Requests.Add((request.Method, request.RequestUri!.ToString(), body, request.Headers.Authorization?.ToString()));
        }

        return _respond(request, body);
    }
}
