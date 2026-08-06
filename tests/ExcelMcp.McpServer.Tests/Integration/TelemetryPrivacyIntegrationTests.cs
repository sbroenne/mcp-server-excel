// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration;

/// <summary>
/// Verifies the telemetry payload that leaves the MCP server process.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Telemetry")]
[Trait("RequiresExcel", "false")]
public sealed class TelemetryPrivacyIntegrationTests
{
    [Fact]
    public void ConfigureTelemetry_OptOutPreventsApplicationInsightsRegistration()
    {
        const string optOutVariable = "EXCELMCP_TELEMETRY_OPTOUT";
        var previousValue = Environment.GetEnvironmentVariable(optOutVariable);

        try
        {
            Environment.SetEnvironmentVariable(optOutVariable, "true");
            var builder = Host.CreateApplicationBuilder();

            Program.ConfigureTelemetry(builder, "InstrumentationKey=00000000-0000-0000-0000-000000000000");

            using var host = builder.Build();
            Assert.Null(host.Services.GetService<Microsoft.ApplicationInsights.TelemetryClient>());
        }
        finally
        {
            Environment.SetEnvironmentVariable(optOutVariable, previousValue);
        }
    }

    [Fact]
    public void TrackUnhandledException_EmitsOnlyRedactedExceptionDetails()
    {
        const string path = @"C:\Users\Ada\Finance\Q4.xlsx";
        const string email = "ada@example.com";
        const string password = "SuperSecretPassword";
        const string connectionSecret = "ConnectionSecretValue";

        var telemetry = ExcelMcpTelemetry.CreateUnhandledExceptionTelemetry(
            new InvalidOperationException(
                $"Could not open {path} for {email}; Password={password}; AccountKey={connectionSecret}"),
            "TelemetryPrivacyIntegrationTests");

        var payload = $"{telemetry.Exception.Message}\n{telemetry.Exception.StackTrace}\n{string.Join("\n", telemetry.Properties.Values)}";

        Assert.DoesNotContain(path, payload, StringComparison.Ordinal);
        Assert.DoesNotContain(email, payload, StringComparison.Ordinal);
        Assert.DoesNotContain(password, payload, StringComparison.Ordinal);
        Assert.DoesNotContain(connectionSecret, payload, StringComparison.Ordinal);
        Assert.Contains("[REDACTED_PATH]", payload, StringComparison.Ordinal);
        Assert.Contains("[REDACTED_EMAIL]", payload, StringComparison.Ordinal);
        Assert.Contains("[REDACTED]", payload, StringComparison.Ordinal);
    }
}
