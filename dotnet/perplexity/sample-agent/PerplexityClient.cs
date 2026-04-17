// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Diagnostics;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace PerplexitySampleAgent;

/// <summary>
/// Async client for Perplexity AI using the Agent API (Responses API).
/// Calls the Perplexity Responses API directly via HttpClient.
/// Supports multi-turn function calling with MCP tools.
/// </summary>
public sealed class PerplexityClient : IDisposable
{
    private const string PerplexityBaseUrl = "https://api.perplexity.ai/v1/responses";
    private const int MaxToolRounds = 8;
    private const int MaxTotalSeconds = 90;
    private const int PerRoundTimeoutSeconds = 30;

    /// <summary>
    /// Maximum tools to send per API call. Reduces payload size and model confusion.
    /// </summary>
    private const int MaxToolsPerCall = 40;

    private readonly HttpClient _http;
    private readonly string _model;
    private readonly string _systemPrompt;
    private readonly ILogger _logger;

    public PerplexityClient(string apiKey, string model, string systemPrompt, ILogger logger)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(apiKey);
        _model = model ?? "perplexity/sonar";
        _systemPrompt = systemPrompt ?? string.Empty;
        _logger = logger;

        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(PerRoundTimeoutSeconds + 5) };
        _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    /// <summary>
    /// Delegate for executing an MCP tool by name with arguments.
    /// </summary>
    public delegate Task<string> ToolExecutorDelegate(string toolName, Dictionary<string, JsonElement> arguments);

    /// <summary>
    /// Send a user message to Perplexity and return the response.
    /// When tools and a toolExecutor are provided, runs a multi-turn tool-call loop.
    /// </summary>
    public async Task<string> InvokeAsync(
        string userMessage,
        IReadOnlyList<JsonElement>? tools = null,
        ToolExecutorDelegate? toolExecutor = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Invoking Perplexity model={Model} (tools={ToolCount})", _model, tools?.Count ?? 0);

        // Filter tools to only the most relevant ones for this message
        var allTools = tools;
        if (tools != null && tools.Count > MaxToolsPerCall)
        {
            tools = FilterRelevantTools(userMessage, tools, MaxToolsPerCall);
            _logger.LogInformation("Filtered {Total} tools down to {Filtered} relevant tools", allTools!.Count, tools.Count);
        }

        var inputItems = new List<JsonElement>();
        inputItems.Add(CreateUserMessageItem(userMessage));

        var invokeStart = Stopwatch.StartNew();
        string lastText = string.Empty;
        string? pendingResourceId = null;
        bool resourceFinalized = false;
        bool retriedWithNudge = false;
        // Track consecutive failures of the same tool to break retry loops
        string? lastFailedTool = null;
        int consecutiveFailures = 0;
        const int MaxConsecutiveFailures = 2;
        bool breakRetryLoop = false;

        for (int round = 0; round < MaxToolRounds; round++)
        {
            if (invokeStart.Elapsed.TotalSeconds > MaxTotalSeconds)
            {
                _logger.LogWarning("Wall-clock limit ({Limit}s) hit after {Rounds} rounds", MaxTotalSeconds, round);
                break;
            }

            if (breakRetryLoop)
                break;

            JsonElement response;
            try
            {
                using var roundCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                roundCts.CancelAfter(TimeSpan.FromSeconds(PerRoundTimeoutSeconds));

                var t0 = Stopwatch.StartNew();
                response = await CallResponsesApiAsync(inputItems, tools, roundCts.Token).ConfigureAwait(false);
                _logger.LogInformation("Perplexity API round {Round} took {Elapsed:F1}s (total {Total:F1}s)",
                    round + 1, t0.Elapsed.TotalSeconds, invokeStart.Elapsed.TotalSeconds);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                _logger.LogWarning("Perplexity API round {Round} timed out ({Timeout}s) — returning partial answer",
                    round + 1, PerRoundTimeoutSeconds);
                break;
            }
            catch (HttpRequestException apiErr) when (tools != null && IsToolRejectionError(apiErr))
            {
                _logger.LogWarning("Tool-call API error — falling back to text-only: {Error}", apiErr.Message);
                tools = null;
                var ctx = ToolsAsContext(allTools);
                if (!string.IsNullOrEmpty(ctx))
                {
                    inputItems.Add(CreateUserMessageItem($"{userMessage}\n\n{ctx}"));
                }
                response = await CallResponsesApiAsync(inputItems, null, cancellationToken).ConfigureAwait(false);
            }

            // Collect function calls and text from output
            var functionCalls = new List<JsonElement>();
            var textParts = new List<string>();

            if (response.TryGetProperty("output", out var outputArr) && outputArr.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in outputArr.EnumerateArray())
                {
                    var itemType = item.TryGetProperty("type", out var typeVal) ? typeVal.GetString() : null;

                    if (itemType == "function_call")
                    {
                        functionCalls.Add(item.Clone());
                    }
                    else if (itemType == "message")
                    {
                        if (item.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var c in content.EnumerateArray())
                            {
                                if (c.TryGetProperty("text", out var textVal) && !string.IsNullOrEmpty(textVal.GetString()))
                                {
                                    textParts.Add(textVal.GetString()!);
                                }
                            }
                        }
                    }
                }
            }

            if (textParts.Count > 0)
            {
                lastText = string.Join("\n", textParts);
            }

            // No function calls → final text response
            if (functionCalls.Count == 0 || toolExecutor == null)
            {
                // Re-prompt: model returned text without calling tools on round 1
                if (round == 0 && tools != null && toolExecutor != null && !retriedWithNudge && UserWantsAction(userMessage))
                {
                    retriedWithNudge = true;
                    var nudge = "Do NOT describe what you would do. You MUST call the appropriate tool " +
                                "right now to complete the user's request. Use the tools provided.";
                    inputItems.Add(CreateUserMessageItem($"{userMessage}\n\n[SYSTEM: {nudge}]"));
                    _logger.LogInformation("Model returned text without tool calls — re-prompting with nudge");
                    continue;
                }

                // Auto-finalize: resource was created but never sent
                if (pendingResourceId != null && !resourceFinalized && toolExecutor != null && UserWantsToSend(userMessage))
                {
                    var sendTool = FindFinalizeTool(allTools);
                    if (sendTool != null)
                    {
                        _logger.LogInformation("Auto-finalizing resource via '{Tool}' (model stopped short)", sendTool);
                        try
                        {
                            // Determine the ID parameter name from the tool schema
                            var idParam = FindIdParam(sendTool, allTools);
                            var args = new Dictionary<string, JsonElement>
                            {
                                [idParam] = JsonSerializer.SerializeToElement(pendingResourceId)
                            };
                            await toolExecutor(sendTool, args).ConfigureAwait(false);
                            resourceFinalized = true;
                        }
                        catch (Exception sendErr)
                        {
                            _logger.LogWarning("Auto-finalize failed: {Error}", sendErr.Message);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(lastText)) return lastText;
                if (response.TryGetProperty("output_text", out var outputText) && !string.IsNullOrEmpty(outputText.GetString()))
                    return outputText.GetString()!;
                return lastText;
            }

            // Tool-call round: append output items then function results
            if (response.TryGetProperty("output", out var outputItems) && outputItems.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in outputItems.EnumerateArray())
                {
                    inputItems.Add(item.Clone());
                }
            }

            foreach (var fc in functionCalls)
            {
                var funcName = fc.TryGetProperty("name", out var nameVal) ? nameVal.GetString() ?? string.Empty : string.Empty;
                var funcArgs = fc.TryGetProperty("arguments", out var argsVal) ? argsVal.GetString() ?? "{}" : "{}";
                var callId = fc.TryGetProperty("call_id", out var callIdVal) ? callIdVal.GetString() ?? string.Empty : string.Empty;

                Dictionary<string, JsonElement> arguments;
                try
                {
                    arguments = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(funcArgs) ?? [];
                }
                catch (JsonException)
                {
                    arguments = [];
                }

                // Strip empty {} / [] values that Perplexity sends for unfilled fields —
                // MCP servers reject them (e.g. "Invalid date format" for startDateTime:{}).
                arguments = StripEmptyValues(arguments);

                // Enrich AFTER stripping so enrichment can fill missing required fields
                arguments = EnrichArguments(funcName, arguments, userMessage, allTools, pendingResourceId);

                _logger.LogInformation("Executing MCP tool: {Tool} (round {Round})", funcName, round + 1);
                _logger.LogInformation("Tool arguments: {Args}", JsonSerializer.Serialize(arguments));

                var result = await toolExecutor(funcName, arguments).ConfigureAwait(false);
                _logger.LogInformation("Tool result (first 500 chars): {Result}", result?.Length > 500 ? result[..500] : result);

                // Track consecutive failures to break hopeless retry loops
                var isError = result != null && result.StartsWith("Error", StringComparison.OrdinalIgnoreCase);
                if (isError && funcName == lastFailedTool)
                {
                    consecutiveFailures++;
                    if (consecutiveFailures >= MaxConsecutiveFailures)
                    {
                        _logger.LogWarning("Tool '{Tool}' failed {Count} times consecutively — breaking retry loop", funcName, consecutiveFailures);
                        inputItems.Add(CreateFunctionCallOutputItem(callId, result ?? string.Empty));
                        breakRetryLoop = true;
                        break; // Exit the foreach; breakRetryLoop stops the outer for-loop
                    }
                }
                else
                {
                    consecutiveFailures = isError ? 1 : 0;
                    lastFailedTool = isError ? funcName : null;
                }

                // Track resource creation/finalization generically
                var toolLower = funcName.ToLowerInvariant();
                // Detect "create" tools — track the resource ID for auto-finalize
                if (Regex.IsMatch(toolLower, @"create|new|add|book|schedule"))
                {
                    var resourceId = ExtractResourceId(result);
                    if (resourceId != null)
                    {
                        pendingResourceId = resourceId;
                        _logger.LogInformation("Tracked created resource: {Id} (from {Tool})",
                            resourceId.Length > 40 ? resourceId[..40] : resourceId, funcName);
                    }
                }
                // Detect "send/submit/finalize" tools
                if (Regex.IsMatch(toolLower, @"send|submit|publish|finalize|confirm|dispatch"))
                {
                    resourceFinalized = true;
                }

                inputItems.Add(CreateFunctionCallOutputItem(callId, result ?? string.Empty));
            }
        }

        // Exhausted rounds — make final summary call without tools
        try
        {
            _logger.LogInformation("Max rounds/time reached — making final summary call");
            using var summaryCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            summaryCts.CancelAfter(TimeSpan.FromSeconds(PerRoundTimeoutSeconds));

            var summary = await CallResponsesApiAsync(inputItems, null, summaryCts.Token).ConfigureAwait(false);
            var summaryText = ExtractText(summary);
            if (!string.IsNullOrEmpty(summaryText)) return summaryText;
        }
        catch (Exception summaryErr)
        {
            _logger.LogWarning("Final summary call failed: {Error}", summaryErr.Message);
        }

        if (!string.IsNullOrEmpty(lastText)) return lastText;
        return "I ran out of time processing your request. The actions may have partially completed — please check and try again if needed.";
    }

    public void Dispose()
    {
        _http.Dispose();
    }

    // ------------------------------------------------------------------
    // HTTP Transport
    // ------------------------------------------------------------------

    private async Task<JsonElement> CallResponsesApiAsync(
        List<JsonElement> inputItems,
        IReadOnlyList<JsonElement>? tools,
        CancellationToken cancellationToken)
    {
        var requestBody = new Dictionary<string, object>
        {
            ["model"] = _model,
            ["input"] = inputItems,
            ["instructions"] = _systemPrompt
        };

        if (tools != null && tools.Count > 0)
        {
            requestBody["tools"] = tools;
        }

        var json = JsonSerializer.Serialize(requestBody);
        using var content = new StringContent(json, Encoding.UTF8, "application/json");
        using var response = await _http.PostAsync(PerplexityBaseUrl, content, cancellationToken).ConfigureAwait(false);

        var responseBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException($"Perplexity API error ({response.StatusCode}): {responseBody}");
        }

        using var doc = JsonDocument.Parse(responseBody);
        return doc.RootElement.Clone();
    }

    // ------------------------------------------------------------------
    // JSON Helpers
    // ------------------------------------------------------------------

    private static JsonElement CreateUserMessageItem(string message)
    {
        var item = new { type = "message", role = "user", content = new[] { new { type = "input_text", text = message } } };
        return JsonSerializer.SerializeToElement(item);
    }

    private static JsonElement CreateFunctionCallOutputItem(string callId, string output)
    {
        var item = new { type = "function_call_output", call_id = callId, output };
        return JsonSerializer.SerializeToElement(item);
    }

    private static string ExtractText(JsonElement response)
    {
        if (response.TryGetProperty("output", out var outputArr) && outputArr.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in outputArr.EnumerateArray())
            {
                var itemType = item.TryGetProperty("type", out var t) ? t.GetString() : null;
                if (itemType == "message" && item.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
                {
                    foreach (var c in content.EnumerateArray())
                    {
                        if (c.TryGetProperty("text", out var textVal) && !string.IsNullOrEmpty(textVal.GetString()))
                        {
                            return textVal.GetString()!;
                        }
                    }
                }
            }
        }
        if (response.TryGetProperty("output_text", out var outputText))
        {
            return outputText.GetString() ?? string.Empty;
        }
        return string.Empty;
    }

    // ------------------------------------------------------------------
    // Argument Enrichment & Helpers
    // ------------------------------------------------------------------

    private static bool IsToolRejectionError(Exception ex)
    {
        var msg = ex.Message.ToLowerInvariant();
        return msg.Contains("not supported") || msg.Contains("unrecognized") ||
               msg.Contains("tool") || msg.Contains("parameter") || msg.Contains("function");
    }

    /// <summary>
    /// Safety net: fill fields the model left empty by inspecting tool schemas.
    /// Pass 1: String content fields (subject/body) from user message.
    /// Pass 2: String ID fields from a previously created resource.
    /// Pass 3: Non-string recipient/address fields built from schema + extracted email.
    /// Fully generic — no tool names or MCP server assumptions.
    /// </summary>
    private static Dictionary<string, JsonElement> EnrichArguments(
        string toolName,
        Dictionary<string, JsonElement> arguments,
        string userMessage,
        IReadOnlyList<JsonElement>? tools,
        string? pendingResourceId = null)
    {
        // Find the schema for this tool
        Dictionary<string, JsonElement>? schemaProps = null;
        if (tools != null)
        {
            foreach (var t in tools)
            {
                var name = t.TryGetProperty("name", out var n) ? n.GetString() : null;
                if (name == toolName)
                {
                    JsonElement paramRoot = default;
                    if (t.TryGetProperty("parameters", out paramRoot) || t.TryGetProperty("inputSchema", out paramRoot))
                    {
                        if (paramRoot.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object)
                        {
                            schemaProps = new Dictionary<string, JsonElement>();
                            foreach (var prop in props.EnumerateObject())
                                schemaProps[prop.Name] = prop.Value;
                        }
                    }
                    break;
                }
            }
        }

        if (schemaProps == null) return arguments;

        // --- Pass 1: String content fields (subject/body) ---
        // Enriches both empty values AND missing required fields.
        var content = ExtractContent(userMessage);
        if (!string.IsNullOrEmpty(content))
        {
            string[] subjectHints = ["subject", "title"];
            string[] bodyHints = ["body", "comment", "content"];

            foreach (var (fieldName, fieldDef) in schemaProps)
            {
                var fieldType = fieldDef.TryGetProperty("type", out var ft) ? ft.GetString() : null;
                if (fieldType != "string") continue;

                if (arguments.TryGetValue(fieldName, out var existing) && !IsEmptyOrNull(existing))
                    continue;

                var fieldLower = fieldName.ToLowerInvariant();
                var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString() ?? "" : "";
                var descLower = desc.ToLowerInvariant();

                if (Regex.IsMatch(descLower, @":\s*\w+\s+or\s+\w+|'[^']+',?\s*'[^']+'"))
                    continue;
                if (Regex.IsMatch(fieldLower, @"type|format|encoding|provider|mode"))
                    continue;

                // Match "subject", "title", "fileName", or exactly "name" (not displayName, firstName, etc.)
                if (subjectHints.Any(h => fieldLower.Contains(h)) || fieldLower == "name" || fieldLower == "filename")
                {
                    arguments[fieldName] = JsonSerializer.SerializeToElement(content);
                }
                else if (bodyHints.Any(h => fieldLower.Contains(h) || descLower.Contains(h)))
                {
                    arguments[fieldName] = JsonSerializer.SerializeToElement(content);
                }
            }
        }

        // --- Pass 1.5: DateTime fields ---
        // If the user provided a date/time, fill any missing datetime fields.
        var extractedDt = ExtractDateTime(userMessage);
        if (extractedDt.HasValue)
        {
            foreach (var (fieldName, fieldDef) in schemaProps)
            {
                if (arguments.TryGetValue(fieldName, out var existing) && !IsEmptyOrNull(existing))
                    continue;
                var fieldType = fieldDef.TryGetProperty("type", out var ftDt) ? ftDt.GetString() : null;
                if (fieldType != "string") continue;

                if (IsDateTimeField(fieldName, fieldDef))
                {
                    var fieldLower = fieldName.ToLowerInvariant();
                    if (fieldLower.Contains("end"))
                    {
                        // Default end = start + 30 min (or extracted duration)
                        var durMins = ParseDurationMinutes(ExtractDuration(userMessage)) ?? 30;
                        arguments[fieldName] = JsonSerializer.SerializeToElement(
                            extractedDt.Value.AddMinutes(durMins).ToString("yyyy-MM-ddTHH:mm:ss"));
                    }
                    else
                    {
                        arguments[fieldName] = JsonSerializer.SerializeToElement(
                            extractedDt.Value.ToString("yyyy-MM-ddTHH:mm:ss"));
                    }
                }
            }
        }

        // --- Pass 1.6: Duration fields (ISO 8601 duration like PT30M) ---
        foreach (var (fieldName, fieldDef) in schemaProps)
        {
            if (arguments.TryGetValue(fieldName, out var existing) && !IsEmptyOrNull(existing))
                continue;
            var fieldType = fieldDef.TryGetProperty("type", out var ftDur) ? ftDur.GetString() : null;
            if (fieldType != "string") continue;

            if (IsDurationField(fieldName, fieldDef))
            {
                var dur = ExtractDuration(userMessage) ?? "PT30M";
                arguments[fieldName] = JsonSerializer.SerializeToElement(dur);
            }
        }

        // --- Pass 2: ID backfill — fill empty identifier fields from pending resource ---
        if (!string.IsNullOrEmpty(pendingResourceId))
        {
            foreach (var (fieldName, fieldDef) in schemaProps)
            {
                if (arguments.TryGetValue(fieldName, out var existing) && !IsEmptyOrNull(existing))
                    continue;
                var fieldType = fieldDef.TryGetProperty("type", out var ftId) ? ftId.GetString() : null;
                if (fieldType != "string") continue;
                if (IsIdField(fieldName, fieldDef))
                {
                    arguments[fieldName] = JsonSerializer.SerializeToElement(pendingResourceId);
                }
            }
        }

        // --- Pass 3: Recipient enrichment — fill empty address/recipient fields ---
        // Only fill cc/bcc if the user explicitly mentioned them.
        var email = ExtractEmail(userMessage);
        if (!string.IsNullOrEmpty(email))
        {
            var msgLower = userMessage.ToLowerInvariant();
            var userMentionsCc = Regex.IsMatch(msgLower, @"\bcc\b");
            var userMentionsBcc = Regex.IsMatch(msgLower, @"\bbcc\b");

            foreach (var (fieldName, fieldDef) in schemaProps)
            {
                if (arguments.TryGetValue(fieldName, out var existing) && !IsEmptyOrNull(existing))
                    continue;
                var fieldType = fieldDef.TryGetProperty("type", out var ftR) ? ftR.GetString() : null;
                if (fieldType == "string") continue; // handled in pass 1

                var fieldLower = fieldName.ToLowerInvariant();
                if (fieldLower == "cc" && !userMentionsCc) continue;
                if (fieldLower == "bcc" && !userMentionsBcc) continue;

                if (IsRecipientField(fieldName, fieldDef))
                {
                    var built = BuildValueFromSchema(fieldDef, email);
                    if (built != null)
                    {
                        arguments[fieldName] = built.Value;
                    }
                }
            }
        }

        return arguments;
    }

    /// <summary>
    /// Check if a JsonElement is effectively empty (null, empty string, empty object, empty array).
    /// </summary>
    private static bool IsEmptyOrNull(JsonElement val)
    {
        return val.ValueKind == JsonValueKind.Null
            || val.ValueKind == JsonValueKind.Undefined
            || (val.ValueKind == JsonValueKind.String && string.IsNullOrEmpty(val.GetString()))
            || (val.ValueKind == JsonValueKind.Object && !val.EnumerateObject().Any())
            || (val.ValueKind == JsonValueKind.Array && val.GetArrayLength() == 0);
    }

    /// <summary>
    /// Remove arguments whose values are empty ({}, [], null, "").
    /// Perplexity sends {} for every unfilled field; MCP servers reject these
    /// (e.g. "Invalid date format" for startDateTime:{}).
    /// </summary>
    private static Dictionary<string, JsonElement> StripEmptyValues(Dictionary<string, JsonElement> arguments)
    {
        var cleaned = new Dictionary<string, JsonElement>(arguments.Count);
        foreach (var (key, value) in arguments)
        {
            if (!IsEmptyOrNull(value))
                cleaned[key] = value;
        }
        return cleaned;
    }

    /// <summary>
    /// Filter tools to the most relevant ones for a given user message.
    /// Scores each tool by keyword overlap (with synonym expansion) between
    /// user message and tool name + description. Returns top N tools.
    /// </summary>
    private static readonly string[][] SynonymGroups =
    [
        ["mail", "email", "message", "draft", "send", "inbox", "reply", "forward"],
        ["meeting", "event", "calendar", "schedule", "appointment", "invite", "book"],
        ["team", "teams", "chat", "channel", "conversation"],
        ["file", "document", "onedrive", "sharepoint", "folder", "upload", "download"],
        ["search", "find", "look", "query"],
        ["create", "new", "add", "make", "compose", "write"],
        ["delete", "remove", "cancel", "decline"],
        ["update", "edit", "modify", "change", "rename"],
        ["list", "get", "read", "show", "view", "fetch"],
        ["excel", "spreadsheet", "workbook", "sheet", "cell"],
        ["word", "document", "doc"],
        ["powerpoint", "presentation", "slide"],
    ];

    private static IReadOnlyList<JsonElement> FilterRelevantTools(
        string userMessage, IReadOnlyList<JsonElement> tools, int maxTools)
    {
        var msgWords = Regex.Split(userMessage.ToLowerInvariant(), @"[^\w@.]+")
            .Where(w => w.Length > 1)
            .ToHashSet();

        // Expand user words with synonyms
        var expandedWords = new HashSet<string>(msgWords);
        foreach (var group in SynonymGroups)
        {
            if (msgWords.Any(w => group.Contains(w)))
            {
                foreach (var syn in group)
                    expandedWords.Add(syn);
            }
        }

        var scored = new List<(JsonElement Tool, int Score)>(tools.Count);

        foreach (var tool in tools)
        {
            var name = (tool.TryGetProperty("name", out var n) ? n.GetString() : "") ?? "";
            var desc = (tool.TryGetProperty("description", out var d) ? d.GetString() : "") ?? "";

            // Split camelCase/PascalCase into words: "CreateDraftMessage" → "create draft message"
            var nameWords = Regex.Split(
                Regex.Replace(name, @"([a-z])([A-Z])", "$1 $2"),
                @"[^\w]+").Select(w => w.ToLowerInvariant()).Where(w => w.Length > 1).ToHashSet();

            var descWords = Regex.Split(desc.ToLowerInvariant(), @"[^\w]+")
                .Where(w => w.Length > 2).ToHashSet();

            int score = 0;

            // Direct + synonym matches against tool name words (high value)
            foreach (var tw in nameWords)
            {
                if (expandedWords.Contains(tw)) score += 3;
            }

            // Direct + synonym matches against description words
            foreach (var dw in descWords)
            {
                if (expandedWords.Contains(dw)) score += 1;
            }

            scored.Add((tool, score));
        }

        return scored
            .OrderByDescending(s => s.Score)
            .Take(maxTools)
            .Select(s => s.Tool)
            .ToList();
    }

    /// <summary>
    /// Extract an email address from the user message.
    /// </summary>
    private static string? ExtractEmail(string text)
    {
        var match = Regex.Match(text, @"[\w.+-]+@[\w-]+\.[\w.]+", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : null;
    }

    /// <summary>
    /// Check if a field name/description suggests it holds an identifier.
    /// </summary>
    private static bool IsIdField(string fieldName, JsonElement fieldDef)
    {
        // Match camelCase ID patterns: "id", "messageId", "event_id", "draftID"
        if (Regex.IsMatch(fieldName, @"(?:^id$|Id$|_id$|ID$)"))
            return true;
        var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString() ?? "" : "";
        return Regex.IsMatch(desc, @"\bid\b|identifier|unique.*key", RegexOptions.IgnoreCase);
    }

    /// <summary>
    /// Check if a field name/description suggests it holds recipients/attendees.
    /// </summary>
    private static bool IsRecipientField(string fieldName, JsonElement fieldDef)
    {
        var lower = fieldName.ToLowerInvariant();
        var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString()?.ToLowerInvariant() ?? "" : "";
        return Regex.IsMatch(lower, @"^(to|cc|bcc)$|recipient|attendee|participant|invitee")
            || Regex.IsMatch(desc, @"recipient|attendee|participant");
    }

    /// <summary>
    /// Build a correctly-shaped JSON value from a schema definition,
    /// placing the email at the deepest string field whose name/description
    /// suggests an email address. Works recursively through array/object nesting.
    /// </summary>
    private static JsonElement? BuildValueFromSchema(JsonElement schemaDef, string email)
    {
        var type = schemaDef.TryGetProperty("type", out var t) ? t.GetString() : null;

        if (type == "array")
        {
            if (schemaDef.TryGetProperty("items", out var items))
            {
                var itemVal = BuildValueFromSchema(items, email);
                if (itemVal != null)
                {
                    using var doc = JsonDocument.Parse("[" + itemVal.Value.GetRawText() + "]");
                    return doc.RootElement.Clone();
                }
            }
            // Fallback: array of strings
            return JsonSerializer.SerializeToElement(new[] { email });
        }

        if (type == "object")
        {
            if (schemaDef.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object)
            {
                var obj = new Dictionary<string, JsonElement>();
                foreach (var prop in props.EnumerateObject())
                {
                    var propType = prop.Value.TryGetProperty("type", out var pt) ? pt.GetString() : null;
                    if (propType == "string" && IsEmailLeaf(prop.Name, prop.Value))
                    {
                        obj[prop.Name] = JsonSerializer.SerializeToElement(email);
                    }
                    else if (propType == "object" || propType == "array")
                    {
                        var nested = BuildValueFromSchema(prop.Value, email);
                        if (nested != null) obj[prop.Name] = nested.Value;
                    }
                }
                if (obj.Count > 0) return JsonSerializer.SerializeToElement(obj);
            }
            return null;
        }

        if (type == "string")
            return JsonSerializer.SerializeToElement(email);

        return null;
    }

    /// <summary>
    /// Check if a string field is the leaf node that should hold an email address.
    /// </summary>
    private static bool IsEmailLeaf(string fieldName, JsonElement fieldDef)
    {
        var lower = fieldName.ToLowerInvariant();
        var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString()?.ToLowerInvariant() ?? "" : "";
        return lower.Contains("address") || lower.Contains("email") || lower.Contains("mail")
            || desc.Contains("email") || desc.Contains("address");
    }

    private static string? ExtractContent(string userMessage)
    {
        string[] patterns =
        [
            @"(?:saying|say)\s+(.+?)(?:\s+and\s+send|\s+right\s+away|$)",
            @"(?:with\s+(?:message|body|text|content|subject))\s+(.+?)$",
            @"(?:that\s+says?)\s+(.+?)$",
            // "named X" / "called X" — stop at "and add/with/on/at/for/from/in" but not "and" alone
            @"(?:named?|called)\s+(.+?)(?:\s+and\s+(?:add|invite|include|remove|with|on|at|for|from|in)\b|\s+(?:on|at|for|from|with|in)\s+|$)",
            @"(?:about)\s+(.+?)$",
            @"(?:titled?)\s+(.+?)$",
        ];

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(userMessage, pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[match.Groups.Count - 1].Value.Trim();
            }
        }
        return null;
    }

    /// <summary>
    /// Extract a DateTime from the user message.
    /// Handles: "16/4/2026 at 9 pm", "today at 3pm", "tomorrow at 14:00", "at 9 pm".
    /// </summary>
    private static DateTime? ExtractDateTime(string text)
    {
        // Pattern 1: Explicit date (d/m/y or m/d/y) + time
        var match = Regex.Match(text,
            @"(\d{1,2})[/\-.]\s*(\d{1,2})[/\-.]\s*(\d{2,4})\s+(?:at\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm)?",
            RegexOptions.IgnoreCase);
        if (match.Success)
        {
            int p1 = int.Parse(match.Groups[1].Value);
            int p2 = int.Parse(match.Groups[2].Value);
            int year = int.Parse(match.Groups[3].Value);
            if (year < 100) year += 2000;
            int hour = int.Parse(match.Groups[4].Value);
            int minute = match.Groups[5].Success && match.Groups[5].Value.Length > 0
                ? int.Parse(match.Groups[5].Value) : 0;
            var ampm = match.Groups[6].Success ? match.Groups[6].Value.ToLowerInvariant() : "";

            if (ampm == "pm" && hour < 12) hour += 12;
            if (ampm == "am" && hour == 12) hour = 0;

            // If p1 > 12, it must be the day (dd/mm format)
            int day, month;
            if (p1 > 12) { day = p1; month = p2; }
            else if (p2 > 12) { day = p2; month = p1; }
            else { day = p1; month = p2; } // Ambiguous — assume d/m

            try { return new DateTime(year, month, day, hour, minute, 0); }
            catch { return null; }
        }

        // Pattern 2: "today"/"tomorrow" + time
        match = Regex.Match(text,
            @"\b(today|tomorrow)\b.*?(?:at\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm)",
            RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var baseDate = match.Groups[1].Value.Equals("tomorrow", StringComparison.OrdinalIgnoreCase)
                ? DateTime.Now.Date.AddDays(1)
                : DateTime.Now.Date;
            int hour = int.Parse(match.Groups[2].Value);
            int minute = match.Groups[3].Success && match.Groups[3].Value.Length > 0
                ? int.Parse(match.Groups[3].Value) : 0;
            var ampm = match.Groups[4].Value.ToLowerInvariant();
            if (ampm == "pm" && hour < 12) hour += 12;
            if (ampm == "am" && hour == 12) hour = 0;
            return baseDate.AddHours(hour).AddMinutes(minute);
        }

        // Pattern 3: Standalone time "at X am/pm" (use today)
        match = Regex.Match(text, @"\bat\s+(\d{1,2})(?::(\d{2}))?\s*(am|pm)", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            int hour = int.Parse(match.Groups[1].Value);
            int minute = match.Groups[2].Success && match.Groups[2].Value.Length > 0
                ? int.Parse(match.Groups[2].Value) : 0;
            var ampm = match.Groups[3].Value.ToLowerInvariant();
            if (ampm == "pm" && hour < 12) hour += 12;
            if (ampm == "am" && hour == 12) hour = 0;
            return DateTime.Now.Date.AddHours(hour).AddMinutes(minute);
        }

        return null;
    }

    /// <summary>
    /// Extract an ISO 8601 duration string from the user message.
    /// Handles: "2 hours", "90 minutes", "1 hour and 30 minutes".
    /// Returns null if no duration is mentioned (caller defaults to PT30M).
    /// </summary>
    private static string? ExtractDuration(string text)
    {
        // "2 hours", "1 hour 30 minutes", "1.5 hours"
        var match = Regex.Match(text,
            @"(\d+(?:\.\d+)?)\s*(?:hours?|hrs?)\s*(?:and\s+)?(?:(\d+)\s*(?:minutes?|mins?))?",
            RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var hours = double.Parse(match.Groups[1].Value);
            var mins = match.Groups[2].Success ? int.Parse(match.Groups[2].Value) : 0;
            var totalMins = (int)(hours * 60) + mins;
            return totalMins >= 60
                ? $"PT{totalMins / 60}H" + (totalMins % 60 > 0 ? $"{totalMins % 60}M" : "")
                : $"PT{totalMins}M";
        }

        match = Regex.Match(text, @"(\d+)\s*(?:minutes?|mins?)", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return $"PT{match.Groups[1].Value}M";
        }

        return null;
    }

    /// <summary>
    /// Parse an ISO 8601 duration string into minutes. Returns null if unparseable.
    /// </summary>
    private static int? ParseDurationMinutes(string? duration)
    {
        if (string.IsNullOrEmpty(duration)) return null;
        int total = 0;
        var hMatch = Regex.Match(duration, @"(\d+)H");
        if (hMatch.Success) total += int.Parse(hMatch.Groups[1].Value) * 60;
        var mMatch = Regex.Match(duration, @"(\d+)M");
        if (mMatch.Success) total += int.Parse(mMatch.Groups[1].Value);
        return total > 0 ? total : null;
    }

    /// <summary>
    /// Check if a field is a date/time field that should be enriched.
    /// Matches: startDateTime, endDateTime, dueDate, etc.
    /// Excludes: timezone, createdTime, dateFormat, etc.
    /// </summary>
    private static bool IsDateTimeField(string fieldName, JsonElement fieldDef)
    {
        var lower = fieldName.ToLowerInvariant();
        // Exclude metadata/zone/format fields
        if (Regex.IsMatch(lower, @"zone|format|created|modified|updated|locale"))
            return false;
        // Contains "datetime" anywhere
        if (lower.Contains("datetime")) return true;
        var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString()?.ToLowerInvariant() ?? "" : "";
        // Contains "date" with ISO/scheduling context in description
        if (lower.Contains("date") && (desc.Contains("iso") || desc.Contains("time") || desc.Contains("yyyy") || desc.Contains("schedule")))
            return true;
        // Starts with start/end/due/begin + description confirms date/time
        if (Regex.IsMatch(lower, @"^(start|end|begin|due)") && (desc.Contains("date") || desc.Contains("time") || desc.Contains("iso") || desc.Contains("when")))
            return true;
        return false;
    }

    /// <summary>
    /// Check if a field is a duration field (ISO 8601 like PT30M, PT1H).
    /// </summary>
    private static bool IsDurationField(string fieldName, JsonElement fieldDef)
    {
        var lower = fieldName.ToLowerInvariant();
        if (lower.Contains("duration")) return true;
        var desc = fieldDef.TryGetProperty("description", out var d) ? d.GetString()?.ToLowerInvariant() ?? "" : "";
        return lower.Contains("length")
            && (desc.Contains("pt") || desc.Contains("iso") || desc.Contains("duration") || desc.Contains("minute") || desc.Contains("hour"));
    }

    private static bool UserWantsToSend(string userMessage)
    {
        var msg = userMessage.ToLowerInvariant();
        return Regex.IsMatch(msg, @"\b(send|mail|email|schedule|create|book|invite|forward|reply)\b")
            && !Regex.IsMatch(msg, @"\bdraft\b");
    }

    private static bool UserWantsAction(string userMessage)
    {
        var msg = userMessage.ToLowerInvariant();
        return Regex.IsMatch(msg, @"\b(send|mail|email|schedule|create|book|set\s+up|arrange|cancel|delete|remove|move|forward|reply|update|add|invite)\b");
    }

    /// <summary>
    /// Extract a resource ID from a tool result, searching common response shapes.
    /// MCP results may contain non-JSON trailers (e.g. CorrelationId lines),
    /// so we try each line separately.
    /// </summary>
    private static string? ExtractResourceId(string? result)
    {
        if (string.IsNullOrEmpty(result)) return null;

        string[] idKeys = ["messageId", "id", "eventId", "itemId", "draftId", "resourceId"];

        // Try each line — MCP responses often append metadata (CorrelationId) after the JSON
        foreach (var line in result.Split('\n'))
        {
            var trimmed = line.Trim();
            if (!trimmed.StartsWith("{")) continue;

            try
            {
                using var doc = JsonDocument.Parse(trimmed);
                var root = doc.RootElement;

                var dataEl = root.TryGetProperty("data", out var d) ? d : default;

                foreach (var key in idKeys)
                {
                    if (dataEl.ValueKind == JsonValueKind.Object
                        && dataEl.TryGetProperty(key, out var val)
                        && val.ValueKind == JsonValueKind.String
                        && !string.IsNullOrEmpty(val.GetString()))
                    {
                        return val.GetString();
                    }
                    if (root.TryGetProperty(key, out var rootVal)
                        && rootVal.ValueKind == JsonValueKind.String
                        && !string.IsNullOrEmpty(rootVal.GetString()))
                    {
                        return rootVal.GetString();
                    }
                }
            }
            catch
            {
                // Not valid JSON on this line — try next
            }
        }

        return null;
    }

    /// <summary>
    /// Find the best send/submit/finalize tool from the schema list.
    /// Scores candidates to prefer draft-specific tools over generic send tools.
    /// </summary>
    private static string? FindFinalizeTool(IReadOnlyList<JsonElement>? tools)
    {
        if (tools == null) return null;

        string? bestName = null;
        int bestScore = 0;

        foreach (var t in tools)
        {
            var rawName = t.TryGetProperty("name", out var n) ? n.GetString() : null;
            if (rawName == null) continue;
            var name = rawName.ToLowerInvariant();
            var desc = (t.TryGetProperty("description", out var d) ? d.GetString() : null)?.ToLowerInvariant() ?? "";

            // Skip "self" tools — these send to user's own notes, not external recipients
            if (name.Contains("self") || name.Contains("note")) continue;

            int score = 0;
            // Best: explicitly about sending a draft
            if (Regex.IsMatch(name, @"send.*draft") || desc.Contains("send") && desc.Contains("draft"))
                score = 3;
            // Good: send + message (but not self)
            else if (Regex.IsMatch(name, @"send.*mail|send.*message|send.*email"))
                score = 2;
            // OK: submit/publish/dispatch
            else if (Regex.IsMatch(name, @"submit|publish|dispatch"))
                score = 1;

            if (score > bestScore)
            {
                bestScore = score;
                bestName = rawName;
            }
        }
        return bestName;
    }

    /// <summary>
    /// Find the ID parameter name for a tool from its schema.
    /// </summary>
    private static string FindIdParam(string toolName, IReadOnlyList<JsonElement>? tools)
    {
        if (tools == null) return "id";
        foreach (var t in tools)
        {
            var name = t.TryGetProperty("name", out var n) ? n.GetString() : null;
            if (name == toolName)
            {
                JsonElement paramRoot = default;
                if (t.TryGetProperty("parameters", out paramRoot) || t.TryGetProperty("inputSchema", out paramRoot))
                {
                    if (paramRoot.TryGetProperty("properties", out var props) && props.ValueKind == JsonValueKind.Object)
                    {
                        // Prefer required ID fields
                        if (paramRoot.TryGetProperty("required", out var req) && req.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var r in req.EnumerateArray())
                            {
                                var rName = r.GetString() ?? "";
                                if (rName.ToLowerInvariant().Contains("id")) return rName;
                            }
                        }
                        // Fall back to any property with "id" in its name
                        foreach (var p in props.EnumerateObject())
                        {
                            if (p.Name.ToLowerInvariant().Contains("id")) return p.Name;
                        }
                    }
                }
            }
        }
        return "id";
    }

    private static string ToolsAsContext(IReadOnlyList<JsonElement>? tools)
    {
        if (tools == null || tools.Count == 0) return string.Empty;
        var sb = new StringBuilder();
        sb.AppendLine("[Available tools for context — these are MCP tools the system has access to:");
        foreach (var t in tools)
        {
            var name = t.TryGetProperty("name", out var n) ? n.GetString() ?? "tool" : "tool";
            var desc = t.TryGetProperty("description", out var d) ? d.GetString() ?? "" : "";
            sb.AppendLine($"- {name}: {desc}");
        }
        sb.Append(']');
        return sb.ToString();
    }
}
