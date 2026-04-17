// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace PerplexitySampleAgent;

/// <summary>
/// Minimal MCP client that speaks JSON-RPC over Streamable HTTP.
/// Ported from the Python Perplexity sample's _McpSession.
/// </summary>
public sealed class McpSession : IAsyncDisposable
{
    private readonly string _url;
    private readonly string _serverName;
    private readonly string _authToken;
    private readonly ILogger _logger;
    private readonly HttpClient _http;
    private string? _sessionId;
    private int _reqId;

    public McpSession(string url, string authToken, string serverName, ILogger logger)
    {
        _url = url;
        _serverName = serverName;
        _authToken = authToken;
        _logger = logger;
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
    }

    public async Task<JsonElement> InitializeAsync(CancellationToken cancellationToken = default)
    {
        var result = await RpcAsync("initialize", new
        {
            protocolVersion = "2025-03-26",
            capabilities = new { },
            clientInfo = new { name = "perplexity-agent-dotnet", version = "0.1.0" }
        }, cancellationToken).ConfigureAwait(false);

        // Notify the server that the client is ready (fire-and-forget)
        await NotifyAsync("notifications/initialized", cancellationToken: cancellationToken).ConfigureAwait(false);
        return result;
    }

    public async Task<List<JsonElement>> ListToolsAsync(CancellationToken cancellationToken = default)
    {
        var result = await RpcAsync("tools/list", new { }, cancellationToken).ConfigureAwait(false);
        if (result.TryGetProperty("tools", out var toolsArr) && toolsArr.ValueKind == JsonValueKind.Array)
        {
            var tools = new List<JsonElement>();
            foreach (var tool in toolsArr.EnumerateArray())
            {
                tools.Add(tool.Clone());
            }
            return tools;
        }
        return [];
    }

    public async Task<string> CallToolAsync(string name, Dictionary<string, JsonElement> arguments, CancellationToken cancellationToken = default)
    {
        var result = await RpcAsync("tools/call", new { name, arguments }, cancellationToken).ConfigureAwait(false);
        if (result.TryGetProperty("content", out var contentArr) && contentArr.ValueKind == JsonValueKind.Array)
        {
            var texts = new List<string>();
            foreach (var c in contentArr.EnumerateArray())
            {
                if (c.TryGetProperty("type", out var typeVal) && typeVal.GetString() == "text" &&
                    c.TryGetProperty("text", out var textVal))
                {
                    texts.Add(textVal.GetString() ?? string.Empty);
                }
            }
            if (texts.Count > 0) return string.Join("\n", texts);
        }
        return result.GetRawText();
    }

    public async ValueTask DisposeAsync()
    {
        _http.Dispose();
        await Task.CompletedTask;
    }

    // -- Transport ----------------------------------------------------------

    private async Task<JsonElement> RpcAsync(string method, object parameters, CancellationToken cancellationToken)
    {
        _reqId++;
        var body = new
        {
            jsonrpc = "2.0",
            id = _reqId,
            method,
            @params = parameters
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, _url);
        request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _authToken);
        if (_sessionId != null)
        {
            request.Headers.TryAddWithoutValidation("Mcp-Session-Id", _sessionId);
        }

        var resp = await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        if (resp.Headers.TryGetValues("mcp-session-id", out var sessionValues))
        {
            _sessionId = sessionValues.FirstOrDefault();
        }

        return await ParseResponseAsync(resp, cancellationToken).ConfigureAwait(false);
    }

    private async Task NotifyAsync(string method, object? parameters = null, CancellationToken cancellationToken = default)
    {
        var body = new
        {
            jsonrpc = "2.0",
            method,
            @params = parameters ?? new { }
        };

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Post, _url);
            request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _authToken);
            if (_sessionId != null)
            {
                request.Headers.TryAddWithoutValidation("Mcp-Session-Id", _sessionId);
            }

            await _http.SendAsync(request, cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            // Notifications are best-effort
        }
    }

    private async Task<JsonElement> ParseResponseAsync(HttpResponseMessage resp, CancellationToken cancellationToken)
    {
        var contentType = resp.Content.Headers.ContentType?.MediaType ?? string.Empty;

        if (contentType.Contains("text/event-stream"))
        {
            var text = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            return ParseSse(text);
        }

        var json = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        if (root.TryGetProperty("error", out var error))
        {
            throw new InvalidOperationException($"MCP error from '{_serverName}': {error.GetRawText()}");
        }

        if (root.TryGetProperty("result", out var result))
        {
            return result.Clone();
        }

        return JsonDocument.Parse("{}").RootElement;
    }

    private static JsonElement ParseSse(string text)
    {
        foreach (var line in text.Split('\n'))
        {
            if (line.StartsWith("data: "))
            {
                try
                {
                    using var doc = JsonDocument.Parse(line[6..]);
                    if (doc.RootElement.TryGetProperty("result", out var result))
                    {
                        return result.Clone();
                    }
                }
                catch (JsonException)
                {
                    continue;
                }
            }
        }
        return JsonDocument.Parse("{}").RootElement;
    }
}
