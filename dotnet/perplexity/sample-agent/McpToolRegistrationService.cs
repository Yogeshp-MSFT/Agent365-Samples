// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Microsoft.Agents.A365.Tooling.Services;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Builder.App.UserAuth;
using Microsoft.Extensions.Logging;

namespace PerplexitySampleAgent;

/// <summary>
/// Delegate for executing a named tool with JSON arguments.
/// </summary>
public delegate Task<string> ToolExecutor(string name, Dictionary<string, JsonElement> arguments);

/// <summary>
/// Discovers MCP servers via the A365 Tooling SDK, connects via JSON-RPC,
/// and returns OpenAI Responses-API-compatible tool definitions plus an execute callback.
/// Ported from the Python Perplexity sample's McpToolRegistrationService.
/// </summary>
public interface IMcpToolRegistrationService
{
    Task<(IReadOnlyList<JsonElement> Tools, ToolExecutor Executor)> GetMcpToolsAsync(
        string agenticAppId,
        UserAuthorization auth,
        string authHandlerName,
        ITurnContext context,
        string? authTokenOverride = null,
        CancellationToken cancellationToken = default);
}

public sealed class McpToolRegistrationService : IMcpToolRegistrationService, IAsyncDisposable
{
    private readonly ILogger<McpToolRegistrationService> _logger;
    private readonly IMcpToolServerConfigurationService _configService;

    // Cached state — survives across turns
    private readonly List<McpSession> _sessions = [];
    private readonly Dictionary<string, McpSession> _toolMap = [];
    private readonly List<JsonElement> _openAiTools = [];
    private bool _initialized;

    // Retry configuration
    private const int MaxRetries = 2;
    private const double RetryBaseDelaySeconds = 1.0;

    public McpToolRegistrationService(
        IMcpToolServerConfigurationService configService,
        ILogger<McpToolRegistrationService> logger)
    {
        _configService = configService;
        _logger = logger;
    }

    public async Task<(IReadOnlyList<JsonElement> Tools, ToolExecutor Executor)> GetMcpToolsAsync(
        string agenticAppId,
        UserAuthorization auth,
        string authHandlerName,
        ITurnContext context,
        string? authTokenOverride = null,
        CancellationToken cancellationToken = default)
    {
        if (_initialized && _openAiTools.Count > 0)
        {
            _logger.LogInformation("Returning {Count} cached MCP tools from {Sessions} sessions",
                _openAiTools.Count, _sessions.Count);
            return (_openAiTools, MakeExecutor());
        }

        var authToken = authTokenOverride;
        if (string.IsNullOrEmpty(authToken))
        {
            authToken = await auth.GetTurnTokenAsync(context, authHandlerName).ConfigureAwait(false);
        }

        _logger.LogInformation("Listing MCP tool servers for agent {AgentId}", agenticAppId);
        var mcpServerConfigs = await _configService.ListToolServersAsync(agenticAppId, authToken!).ConfigureAwait(false);
        _logger.LogInformation("Loaded {Count} MCP server configurations", mcpServerConfigs?.Count ?? 0);

        if (mcpServerConfigs == null || mcpServerConfigs.Count == 0)
        {
            _initialized = true;
            return (_openAiTools, MakeExecutor());
        }

        // Connect to all MCP servers in parallel
        var tasks = mcpServerConfigs.Select(cfg => ConnectServerAsync(cfg, authToken!, cancellationToken)).ToList();
        var results = await Task.WhenAll(tasks).ConfigureAwait(false);

        foreach (var result in results)
        {
            if (result == null) continue;
            var (session, tools) = result.Value;
            _sessions.Add(session);

            foreach (var tool in tools)
            {
                var name = tool.TryGetProperty("name", out var nameVal) ? nameVal.GetString() ?? string.Empty : string.Empty;
                if (string.IsNullOrEmpty(name)) continue;

                // Sanitize inputSchema for Perplexity strictness
                JsonElement? rawSchema = tool.TryGetProperty("inputSchema", out var schemaVal) ? schemaVal : null;
                var parameters = SanitizeSchema(rawSchema);

                var description = tool.TryGetProperty("description", out var descVal) ? descVal.GetString() ?? string.Empty : string.Empty;

                // Create Responses API tool definition as JsonElement
                _openAiTools.Add(JsonSerializer.SerializeToElement(new Dictionary<string, object?>
                {
                    ["type"] = "function",
                    ["name"] = name,
                    ["description"] = description,
                    ["parameters"] = parameters
                }));

                _toolMap[name] = session;
            }
        }

        if (_openAiTools.Count == 0)
        {
            _logger.LogInformation("No MCP tools discovered — running without tools");
        }
        else
        {
            _logger.LogInformation("Registered {ToolCount} MCP tools from {ServerCount} servers",
                _openAiTools.Count, _sessions.Count);
        }

        _initialized = true;
        return (_openAiTools, MakeExecutor());
    }

    private async Task<(McpSession Session, List<JsonElement> Tools)?> ConnectServerAsync(
        object serverConfig,
        string authToken,
        CancellationToken cancellationToken)
    {
        // Extract URL and name from the server config using reflection
        // (the SDK config objects expose these as properties)
        var configType = serverConfig.GetType();
        var urlProp = configType.GetProperty("Url") ?? configType.GetProperty("url");
        var nameProp = configType.GetProperty("McpServerName") ?? configType.GetProperty("mcpServerName");
        var uniqueProp = configType.GetProperty("McpServerUniqueName") ?? configType.GetProperty("mcpServerUniqueName");

        var rawUrl = urlProp?.GetValue(serverConfig)?.ToString();
        var rawName = nameProp?.GetValue(serverConfig)?.ToString() ?? string.Empty;
        var rawUnique = uniqueProp?.GetValue(serverConfig)?.ToString() ?? string.Empty;

        string? serverUrl;
        if (!string.IsNullOrEmpty(rawUrl))
            serverUrl = rawUrl;
        else if (rawUnique.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            serverUrl = rawUnique;
        else
            serverUrl = null;

        var serverName = !string.IsNullOrEmpty(rawName) ? rawName :
            (!rawUnique.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? rawUnique : "unknown");

        _logger.LogInformation("MCP server '{Name}' -> {Url}", serverName, serverUrl ?? "(no URL)");

        if (string.IsNullOrEmpty(serverUrl))
        {
            _logger.LogWarning("Skipping MCP server '{Name}' — no URL configured.", serverName);
            return null;
        }

        try
        {
            var session = new McpSession(serverUrl, authToken, serverName, _logger);
            await session.InitializeAsync(cancellationToken).ConfigureAwait(false);
            var tools = await session.ListToolsAsync(cancellationToken).ConfigureAwait(false);
            _logger.LogInformation("Server '{Name}' exposes {Count} tools", serverName, tools.Count);
            return (session, tools);
        }
        catch (Exception exc)
        {
            _logger.LogWarning("Failed to connect to MCP server '{Name}' at {Url}: {Error}", serverName, serverUrl, exc.Message);
            return null;
        }
    }

    private ToolExecutor MakeExecutor()
    {
        var toolMap = _toolMap;
        var svc = this;

        return async (string name, Dictionary<string, JsonElement> arguments) =>
        {
            if (!toolMap.TryGetValue(name, out var session))
            {
                return $"Error: unknown tool '{name}'";
            }

            string? lastError = null;
            for (int attempt = 0; attempt <= MaxRetries; attempt++)
            {
                try
                {
                    return await session.CallToolAsync(name, arguments).ConfigureAwait(false);
                }
                catch (HttpRequestException exc)
                {
                    lastError = exc.Message;
                    svc._logger.LogWarning("Connection error on attempt {Attempt}/{Max} for tool '{Tool}': {Error}",
                        attempt + 1, MaxRetries + 1, name, exc.Message);
                }
                catch (Exception exc)
                {
                    svc._logger.LogError("Tool call '{Tool}' failed: {Error}", name, exc.Message);
                    return $"Error executing tool '{name}': {exc.Message}";
                }

                if (attempt < MaxRetries)
                {
                    var delay = RetryBaseDelaySeconds * Math.Pow(2, attempt) + Random.Shared.NextDouble() * 0.5;
                    svc._logger.LogInformation("Retrying tool '{Tool}' in {Delay:F2}s…", name, delay);
                    await Task.Delay(TimeSpan.FromSeconds(delay)).ConfigureAwait(false);
                }
            }

            // All retries exhausted — invalidate cache
            svc._logger.LogError("Tool '{Tool}' failed after {Max} attempts — clearing MCP cache", name, MaxRetries + 1);
            svc._initialized = false;
            return $"Error executing tool '{name}': {lastError}";
        };
    }

    /// <summary>
    /// Sanitize an MCP inputSchema for Perplexity's stricter requirements.
    /// </summary>
    private static Dictionary<string, object?> SanitizeSchema(JsonElement? raw)
    {
        var empty = new Dictionary<string, object?> { ["type"] = "object", ["properties"] = new Dictionary<string, object?>() };

        if (raw == null || raw.Value.ValueKind != JsonValueKind.Object) return empty;

        if (!raw.Value.TryGetProperty("type", out var typeVal) || typeVal.GetString() != "object")
            return empty;

        return CleanSchema(raw.Value);
    }

    private static readonly HashSet<string> UnsupportedKeys =
    [
        "$defs", "$ref", "additionalProperties", "allOf", "anyOf",
        "oneOf", "not", "$schema", "definitions"
    ];

    private static Dictionary<string, object?> CleanSchema(JsonElement schema)
    {
        var cleaned = new Dictionary<string, object?>();

        foreach (var prop in schema.EnumerateObject())
        {
            if (UnsupportedKeys.Contains(prop.Name)) continue;

            // Remove empty "required" arrays
            if (prop.Name == "required" && prop.Value.ValueKind == JsonValueKind.Array && prop.Value.GetArrayLength() == 0)
                continue;

            if (prop.Name == "properties" && prop.Value.ValueKind == JsonValueKind.Object)
            {
                var cleanedProps = new Dictionary<string, object?>();
                foreach (var subProp in prop.Value.EnumerateObject())
                {
                    if (subProp.Value.ValueKind == JsonValueKind.Object)
                    {
                        cleanedProps[subProp.Name] = CleanSchema(subProp.Value);
                    }
                }
                cleaned["properties"] = cleanedProps;
            }
            else if (prop.Name == "items" && prop.Value.ValueKind == JsonValueKind.Object)
            {
                cleaned["items"] = CleanSchema(prop.Value);
            }
            else
            {
                cleaned[prop.Name] = JsonSerializer.Deserialize<object>(prop.Value.GetRawText());
            }
        }

        if (!cleaned.ContainsKey("properties"))
        {
            cleaned["properties"] = new Dictionary<string, object?>();
        }

        return cleaned;
    }

    public async ValueTask DisposeAsync()
    {
        foreach (var session in _sessions)
        {
            try { await session.DisposeAsync(); } catch { }
        }
        _sessions.Clear();
        _toolMap.Clear();
        _openAiTools.Clear();
        _initialized = false;
    }
}
