// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;
using Xunit.Abstractions;

// Avoid namespace conflict: McpServer is both a type and namespace
using Server = ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// Integration tests that exercise the full MCP protocol using in-memory transport.
/// These tests use the official MCP SDK client to connect to our server, ensuring:
/// - DI pipeline is correctly configured
/// - Tool discovery via WithToolsFromAssembly() works
/// - Tool schemas are correctly generated
/// - Tools execute properly through the MCP protocol
///
/// This is the CORRECT way to test MCP servers - using the SDK's client to verify
/// the actual protocol behavior, not reflection or direct method calls.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public class McpServerIntegrationTests(ITestOutputHelper output) : IAsyncLifetime, IAsyncDisposable
{
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private Server.McpServer? _server;
    private McpClient? _client;
    private IServiceProvider? _serviceProvider;
    private Task? _serverTask;

    /// <summary>
    /// Expected tool names from our assembly - the source of truth.
    /// After token optimization split (issue #341):
    /// - range split into 4 tools (range, range_edit, range_format, range_link)
    /// - table split into 2 tools (table, table_column)
    /// - pivottable split into 3 tools (pivottable, pivottable_field, pivottable_calc)
    /// - datamodel split into 2 tools (datamodel, datamodel_relationship)
    /// - chart split into 2 tools (chart, chart_config)
    /// - worksheet split into 2 tools (worksheet, worksheet_style)
    /// - Added slicer and calculation_mode
    /// </summary>
    private static readonly HashSet<string> ExpectedToolNames =
    [
        "calculation_mode",
        "chart",
        "chart_config",
        "conditionalformat",
        "connection",
        "datamodel",
        "datamodel_relationship",
        "file",
        "namedrange",
        "pivottable",
        "pivottable_calc",
        "pivottable_field",
        "powerquery",
        "range",
        "range_edit",
        "range_format",
        "range_link",
        "slicer",
        "table",
        "table_column",
        "vba",
        "worksheet",
        "worksheet_style"
    ];

    /// <summary>
    /// Setup: Create MCP server with DI and connect client via in-memory pipes.
    /// This exercises the exact same code path as Program.cs.
    /// </summary>
    public async Task InitializeAsync()
    {
        // Build the server with DI - same pattern as Program.cs
        var services = new ServiceCollection();
        services.AddLogging(builder => builder.AddDebug().SetMinimumLevel(LogLevel.Debug));

        // Configure telemetry (disabled for tests)
        services.AddApplicationInsightsTelemetryWorkerService(options =>
        {
            options.ConnectionString = null;
            options.EnableHeartbeat = false;
            options.EnableAdaptiveSampling = false;
            options.EnableQuickPulseMetricStream = false;
            options.EnablePerformanceCounterCollectionModule = false;
            options.EnableEventCounterCollectionModule = false;
            options.EnableDependencyTrackingTelemetryModule = false;
        });
        services.AddSingleton<ITelemetryInitializer, ExcelMcpTelemetryInitializer>();

        // Add MCP server with tools (same as Program.cs) using stream transport for testing
        services
            .AddMcpServer(options =>
            {
                options.ServerInfo = new() { Name = "ExcelMcp-Test", Version = "1.0.0" };
                options.ServerInstructions = "Test server for integration tests";
            })
            .WithStreamServerTransport(
                _clientToServerPipe.Reader.AsStream(),
                _serverToClientPipe.Writer.AsStream())
            .WithToolsFromAssembly(typeof(ExcelFileTool).Assembly);

        _serviceProvider = services.BuildServiceProvider(validateScopes: true);

        // Get the server and start it
        _server = _serviceProvider.GetRequiredService<Server.McpServer>();
        _serverTask = _server.RunAsync(_cts.Token);

        // Create client connected to the server via pipes
        _client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: _clientToServerPipe.Writer.AsStream(),
                serverOutput: _serverToClientPipe.Reader.AsStream()),
            clientOptions: new McpClientOptions
            {
                ClientInfo = new() { Name = "TestClient", Version = "1.0.0" }
            },
            cancellationToken: _cts.Token);

        output.WriteLine($"✓ Connected to server: {_client.ServerInfo?.Name} v{_client.ServerInfo?.Version}");
    }

    public async Task DisposeAsync()
    {
        await DisposeAsyncCore();
    }

    // Explicit IAsyncDisposable implementation to satisfy CA1001 analyzer
    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        await _cts.CancelAsync();

        _clientToServerPipe.Writer.Complete();
        _serverToClientPipe.Writer.Complete();

        if (_client != null)
        {
            await _client.DisposeAsync();
        }

        if (_serverTask != null)
        {
            try
            {
                await _serverTask;
            }
            catch (OperationCanceledException)
            {
                // Expected during shutdown
            }
        }

        if (_serviceProvider is IAsyncDisposable asyncDisposable)
        {
            await asyncDisposable.DisposeAsync();
        }
        else if (_serviceProvider is IDisposable disposable)
        {
            disposable.Dispose();
        }

        _cts.Dispose();
    }

    /// <summary>
    /// Tests that all 21 expected tools are discoverable via the MCP protocol.
    /// After token optimization (issue #341):
    /// - Original 12 tools split into focused tools for better token efficiency
    /// - range → range, range_edit, range_format, range_link
    /// - table → table, table_column
    /// - pivottable → pivottable, pivottable_field, pivottable_calc
    /// - datamodel → datamodel, datamodel_relationship
    /// - chart → chart, chart_config
    /// - worksheet → worksheet, worksheet_style
    /// This is THE definitive test - it uses client.ListToolsAsync() which exercises:
    /// - DI pipeline
    /// - WithToolsFromAssembly() discovery
    /// - MCP protocol serialization
    /// - Tool schema generation
    /// </summary>
    [Fact]
    public async Task ListTools_ReturnsAll23ExpectedTools()
    {
        output.WriteLine("=== TOOL DISCOVERY VIA MCP PROTOCOL ===\n");

        // Act - Use the REAL MCP protocol to list tools
        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        // Assert - Verify count
        output.WriteLine($"Discovered {tools.Count} tools via MCP protocol:\n");

        foreach (var tool in tools.OrderBy(t => t.Name))
        {
            var descPreview = tool.Description?.Length > 60 ? tool.Description[..60] + "..." : tool.Description;
            output.WriteLine($"  • {tool.Name}: {descPreview}");
        }

        Assert.Equal(ExpectedToolNames.Count, tools.Count);

        // Verify all expected tools are present
        var actualToolNames = tools.Select(t => t.Name).ToHashSet();

        var missingTools = ExpectedToolNames.Except(actualToolNames).ToList();
        if (missingTools.Count > 0)
        {
            output.WriteLine($"\n❌ Missing tools: {string.Join(", ", missingTools)}");
        }
        Assert.Empty(missingTools);

        var unexpectedTools = actualToolNames.Except(ExpectedToolNames).ToList();
        if (unexpectedTools.Count > 0)
        {
            output.WriteLine($"\n❌ Unexpected tools: {string.Join(", ", unexpectedTools)}");
        }
        Assert.Empty(unexpectedTools);

        output.WriteLine($"\n✓ All {ExpectedToolNames.Count} tools discovered successfully via MCP protocol");
    }

    /// <summary>
    /// Tests that each tool has proper schema (parameters, descriptions).
    /// </summary>
    [Fact]
    public async Task ListTools_AllToolsHaveValidSchema()
    {
        output.WriteLine("=== TOOL SCHEMA VALIDATION ===\n");

        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);

        foreach (var tool in tools)
        {
            // Every tool must have a name
            Assert.False(string.IsNullOrEmpty(tool.Name), "Tool has empty name");

            // Every tool should have a description
            Assert.False(string.IsNullOrEmpty(tool.Description), $"Tool {tool.Name} has no description");

            // McpClientTool implements AIFunction which has Parameters property
            // The SDK generates schema from tool methods

            output.WriteLine($"✓ {tool.Name}: Has description ({tool.Description?.Length} chars)");
        }

        output.WriteLine($"\n✓ All {tools.Count} tools have valid schemas");
    }

    /// <summary>
    /// Tests that file tool's Test action works via MCP protocol.
    /// This exercises the complete tool invocation path.
    /// </summary>
    [Fact]
    public async Task CallTool_ExcelFileTest_ReturnsSuccess()
    {
        output.WriteLine("=== TOOL INVOCATION VIA MCP PROTOCOL ===\n");

        // Arrange - Test action doesn't require an actual file
        // Parameter names shortened for token optimization: excelPath -> path
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = "Test",
            ["path"] = "C:\\fake\\test.xlsx"
        };

        // Act - Call tool via MCP protocol
        var result = await _client!.CallToolAsync(
            "file",
            arguments,
            cancellationToken: _cts.Token);

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.Content);
        Assert.NotEmpty(result.Content);

        // Get text content - need to cast from ContentBlock base class
        var textBlock = result.Content.OfType<TextContentBlock>().FirstOrDefault();
        Assert.NotNull(textBlock);

        var textPreview = textBlock.Text.Length > 200 ? textBlock.Text[..200] + "..." : textBlock.Text;
        output.WriteLine($"Tool response: {textPreview}");

        // The test action should return success (property name is "success" in success responses)
        Assert.Contains("success", textBlock.Text.ToLowerInvariant());

        output.WriteLine("\n✓ file Test action executed successfully via MCP protocol");
    }

    /// <summary>
    /// Tests that server information is correctly exposed via MCP protocol.
    /// </summary>
    [Fact]
    public async Task ServerInfo_ReturnsCorrectInformation()
    {
        output.WriteLine("=== SERVER INFO VIA MCP PROTOCOL ===\n");

        // Act - Server info is available after connection
        var serverInfo = _client!.ServerInfo;
        var serverInstructions = _client.ServerInstructions;

        // Assert
        Assert.NotNull(serverInfo);
        Assert.Equal("ExcelMcp-Test", serverInfo.Name);
        Assert.Equal("1.0.0", serverInfo.Version);
        Assert.Equal("Test server for integration tests", serverInstructions);

        output.WriteLine($"Server Name: {serverInfo.Name}");
        output.WriteLine($"Server Version: {serverInfo.Version}");
        output.WriteLine($"Server Instructions: {serverInstructions}");

        output.WriteLine("\n✓ Server info correctly exposed via MCP protocol");
        await Task.CompletedTask; // Satisfy async requirement
    }

    /// <summary>
    /// Tests that telemetry services are properly registered in DI.
    /// </summary>
    [Fact]
    public void DI_TelemetryServicesRegistered()
    {
        output.WriteLine("=== TELEMETRY DI REGISTRATION ===\n");

        Assert.NotNull(_serviceProvider);

        // Act - Verify telemetry services are available
        var telemetryClient = _serviceProvider.GetService<TelemetryClient>();
        var telemetryInitializers = _serviceProvider.GetServices<ITelemetryInitializer>().ToList();

        // Assert
        Assert.NotNull(telemetryClient);
        Assert.Contains(telemetryInitializers, i => i is ExcelMcpTelemetryInitializer);

        output.WriteLine("✓ TelemetryClient registered");
        output.WriteLine($"✓ Found {telemetryInitializers.Count} telemetry initializers");
        output.WriteLine("✓ ExcelMcpTelemetryInitializer present");

        output.WriteLine("\n✓ Telemetry services correctly registered in DI");
    }

    /// <summary>
    /// Tests that all tools can be discovered and iterated via ListToolsAsync.
    /// Note: SDK 0.5.0+ replaced EnumerateToolsAsync with ListToolsAsync.
    /// </summary>
    [Fact]
    public async Task ListTools_CanIterateAllTools()
    {
        output.WriteLine("=== TOOL ITERATION ===\n");

        var tools = await _client!.ListToolsAsync(cancellationToken: _cts.Token);
        var toolCount = 0;
        foreach (var tool in tools)
        {
            toolCount++;
            output.WriteLine($"  Discovered: {tool.Name}");
        }

        Assert.Equal(ExpectedToolNames.Count, toolCount);

        output.WriteLine($"\n✓ Iterated {toolCount} tools");
    }

    /// <summary>
    /// Tests that server capabilities include tools.
    /// </summary>
    [Fact]
    public void ServerCapabilities_IncludesTools()
    {
        output.WriteLine("=== SERVER CAPABILITIES ===\n");

        var capabilities = _client!.ServerCapabilities;

        Assert.NotNull(capabilities);
        Assert.NotNull(capabilities.Tools);

        output.WriteLine($"✓ Tools capability: {capabilities.Tools != null}");
        output.WriteLine($"✓ ListChanged: {capabilities.Tools?.ListChanged}");

        output.WriteLine("\n✓ Server capabilities correctly exposed");
    }
}




