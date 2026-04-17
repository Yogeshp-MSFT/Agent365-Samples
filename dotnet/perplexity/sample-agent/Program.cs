// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using PerplexitySampleAgent;
using PerplexitySampleAgent.Agent;
using PerplexitySampleAgent.telemetry;
using Microsoft.Agents.A365.Observability;
using Microsoft.Agents.A365.Tooling.Services;
using Microsoft.Agents.Builder;
using Microsoft.Agents.Core;
using Microsoft.Agents.Hosting.AspNetCore;
using Microsoft.Agents.Storage;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// Setup OpenTelemetry
builder.ConfigureOpenTelemetry();

builder.Configuration.AddUserSecrets(Assembly.GetExecutingAssembly());
builder.Services.AddControllers();
builder.Services.AddHttpClient("WebClient", client => client.Timeout = TimeSpan.FromSeconds(600));
builder.Services.AddHttpContextAccessor();
builder.Logging.AddConsole();

// **********  Configure A365 Services **********
// Configure observability
builder.Services.AddAgenticTracingExporter(clusterCategory: "production");

// Add A365 Tooling Server integration (using our custom implementation, not the framework extension)
builder.Services.AddSingleton<IMcpToolServerConfigurationService, McpToolServerConfigurationService>();
builder.Services.AddSingleton<IMcpToolRegistrationService, PerplexitySampleAgent.McpToolRegistrationService>();
// **********  END Configure A365 Services **********

// Add AspNet token validation
builder.Services.AddAgentAspNetAuthentication(builder.Configuration);

// Register IStorage for development
builder.Services.AddSingleton<IStorage, MemoryStorage>();

// Add AgentApplicationOptions from config
builder.AddAgentApplicationOptions();

// Add the agent
builder.AddAgent<MyAgent>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}

app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

// Map the /api/messages endpoint
app.MapPost("/api/messages", async (HttpRequest request, HttpResponse response, IAgentHttpAdapter adapter, IAgent agent, CancellationToken cancellationToken) =>
{
    await AgentMetrics.InvokeObservedHttpOperation("agent.process_message", async () =>
    {
        await adapter.ProcessAsync(request, response, agent, cancellationToken);
    }).ConfigureAwait(false);
});

// Health check endpoint
app.MapGet("/api/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

if (app.Environment.IsDevelopment() || app.Environment.EnvironmentName == "Playground")
{
    app.MapGet("/", () => "Perplexity Sample Agent");
    app.UseDeveloperExceptionPage();
    app.MapControllers().AllowAnonymous();
    app.Urls.Add("http://localhost:3978");
}
else
{
    app.MapControllers();
}

app.Run();
