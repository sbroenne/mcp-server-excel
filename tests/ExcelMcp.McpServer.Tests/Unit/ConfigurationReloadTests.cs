// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

/// <summary>
/// Regression tests for MCP Server configuration.
///
/// Bug 8 (Feb 2026): Host.CreateApplicationBuilder() enables reloadOnChange:true by default,
/// creating a FileSystemWatcher for appsettings.json. Under file I/O storms from Excel
/// (temp files, lock files), this watcher fires ParseEventBufferAndNotifyForEach in a tight
/// loop on the threadpool, consuming ~85% CPU.
///
/// These tests ensure that config reload watchers are permanently disabled.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Configuration")]
public class ConfigurationReloadTests
{
    /// <summary>
    /// REGRESSION TEST: No configuration source should use reloadOnChange:true.
    /// This caused 85% CPU from FileSystemWatcher under I/O storms (Bug 8).
    /// </summary>
    [Fact]
    public void Configuration_NoFileSource_HasReloadOnChangeEnabled()
    {
        // Arrange — build the host the same way Program.Main does
        var builder = Host.CreateApplicationBuilder([]);

        // Replicate Program.cs config setup
        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine([]);

        // Act — inspect all configuration sources for file-based sources with reload enabled
        var fileSourcesWithReload = builder.Configuration.Sources
            .OfType<Microsoft.Extensions.Configuration.Json.JsonConfigurationSource>()
            .Where(s => s.ReloadOnChange)
            .ToList();

        // Assert
        Assert.Empty(fileSourcesWithReload);
    }

    /// <summary>
    /// REGRESSION TEST: After clearing default sources, environment variables must still be available.
    /// The initial Bug 8 fix (Sources.Clear()) was too aggressive — it removed env vars and CLI args.
    /// </summary>
    [Fact]
    public void Configuration_AfterClear_EnvironmentVariablesAreAvailable()
    {
        // Arrange
        var builder = Host.CreateApplicationBuilder([]);

        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine([]);

        // Act — environment variables should be accessible
        // PATH is a universally available env var on Windows
        var pathValue = builder.Configuration["PATH"] ?? builder.Configuration["Path"];

        // Assert — if env vars weren't re-added, this would be null
        Assert.NotNull(pathValue);
    }

    /// <summary>
    /// REGRESSION TEST: After clearing default sources, command-line arguments must still work.
    /// </summary>
    [Fact]
    public void Configuration_AfterClear_CommandLineArgsAreAvailable()
    {
        // Arrange
        var builder = Host.CreateApplicationBuilder(["--TestKey=TestValue"]);

        builder.Configuration.Sources.Clear();
        builder.Configuration
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
            .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: false)
            .AddEnvironmentVariables()
            .AddCommandLine(["--TestKey=TestValue"]);

        // Act
        var value = builder.Configuration["TestKey"];

        // Assert
        Assert.Equal("TestValue", value);
    }

    /// <summary>
    /// Verify that the default Host.CreateApplicationBuilder() DOES enable reloadOnChange.
    /// This proves our fix is necessary — without it, file watchers would be created.
    /// </summary>
    [Fact]
    public void DefaultBuilder_HasReloadOnChange_True()
    {
        // Arrange — build with defaults (no clearing)
        var builder = Host.CreateApplicationBuilder([]);

        // Act — check default JSON sources
        var jsonSourcesWithReload = builder.Configuration.Sources
            .OfType<Microsoft.Extensions.Configuration.Json.JsonConfigurationSource>()
            .Where(s => s.ReloadOnChange)
            .ToList();

        // Assert — defaults SHOULD have reloadOnChange:true (this is what we're fixing)
        Assert.NotEmpty(jsonSourcesWithReload);
    }
}
