# Perplexity Sample Agent (.NET)

A .NET implementation of the Perplexity AI sample agent using the Microsoft Agent 365 SDK. This agent leverages Perplexity's OpenAI-compatible Responses API with function calling support and MCP (Model Context Protocol) tools for rich integration with Microsoft 365 services.

## Architecture

This sample demonstrates:

- **Perplexity AI Integration**: Uses the OpenAI .NET SDK with a custom base URL pointing to `https://api.perplexity.ai/v1`, enabling access to Perplexity's Responses API with function calling.
- **Custom MCP Client**: A lightweight JSON-RPC client (`McpSession`) that connects to MCP servers over Streamable HTTP, supporting initialization, tool discovery, and tool execution.
- **MCP Tool Registration**: Discovers MCP servers via the A365 Tooling SDK, connects in parallel, sanitizes tool schemas for Perplexity's stricter requirements, and provides cached tool definitions with retry logic.
- **Multi-Turn Tool Loop**: The `PerplexityClient` runs up to 8 rounds of tool calls with a 90-second wall-clock limit, including auto-finalize for create→send workflows, argument enrichment, and nudge retry when the model describes actions instead of executing them.
- **A365 Observability**: Full OpenTelemetry integration with `BaggageBuilder` scopes, custom metrics, and agentic token caching for the observability exporter.
- **Notifications**: Handles email and Word comment notifications from the A365 platform.

## Project Structure

```
sample-agent/
├── Agent/
│   └── MyAgent.cs              # AgentApplication — message/notification handlers, typing loop
├── telemetry/
│   ├── AgentMetrics.cs          # Custom ActivitySource, Meter, counters, histograms
│   ├── AgentOTELExtensions.cs   # OpenTelemetry builder configuration
│   └── A365OtelWrapper.cs       # BaggageBuilder + observability token cache wrapper
├── appPackage/
│   └── manifest.json            # Teams app manifest
├── AspNetExtensions.cs          # JWT bearer token validation
├── McpSession.cs                # JSON-RPC over Streamable HTTP client
├── McpToolRegistrationService.cs # MCP server discovery + tool registration
├── PerplexityClient.cs          # Perplexity Responses API client with tool-call loop
├── Program.cs                   # ASP.NET Core startup and DI
├── appsettings.json             # Production configuration
├── appsettings.Playground.json  # Development/Playground configuration
├── ToolingManifest.json         # MCP server definitions (Mail + Calendar)
└── PerplexitySampleAgent.csproj # Project file
```

## Prerequisites

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) or later
- A [Perplexity API key](https://docs.perplexity.ai/) — set in `appsettings.json` or as environment variable `PERPLEXITY_API_KEY`
- Agent 365 registration (run `a365 config init`) for MCP tool and auth setup

## Configuration

### Perplexity API

Set your API key in `appsettings.json`:

```json
{
  "Perplexity": {
    "ApiKey": "your-api-key-here",
    "Model": "perplexity/sonar"
  }
}
```

Or via environment variable:

```
PERPLEXITY_API_KEY=your-api-key-here
PERPLEXITY_MODEL=perplexity/sonar
```

Available models: `perplexity/sonar`, `openai/gpt-5.4`, `anthropic/claude-sonnet-4-6`, etc. See [Perplexity Agent API Models](https://docs.perplexity.ai/docs/agent-api/models).

### Agent 365 Service Connection

Follow the same registration flow as other .NET A365 samples. After running `a365 config init`, fill in `appsettings.json` or `appsettings.Playground.json` with:
- `TokenValidation:Audiences` — your app registration Client ID
- `Connections:ServiceConnection:Settings` — client credentials

## Running

```bash
cd dotnet/perplexity/sample-agent
dotnet run
```

The agent starts on `http://localhost:3978` in development mode. Use the Agent 365 Playground or Teams to interact.

## Key Differences from Agent Framework Sample

| Aspect | Agent Framework Sample | Perplexity Sample |
|--------|----------------------|-------------------|
| LLM Provider | Azure OpenAI via `IChatClient` | Perplexity via OpenAI SDK with custom endpoint |
| Tool Integration | `Tooling.Extensions.AgentFramework` + `IChatClient` | Custom `McpSession` + `PerplexityClient` tool loop |
| Response Handling | Streaming via `ChatClientAgent.RunStreamingAsync` | Non-streaming with multi-turn tool loop |
| Local Tools | Weather lookup, DateTime | None (MCP tools only) |
| Tool Calling | Automatic via `IChatClient.UseFunctionInvocation()` | Manual loop in `PerplexityClient.InvokeAsync()` |

## Reference

This .NET sample is a port of the [Python Perplexity sample](../../python/perplexity/sample-agent/), following the same architecture and system prompt pattern.
