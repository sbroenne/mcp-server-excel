// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Reflection;

using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.Extensibility;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Telemetry initializer that sets User.Id, Session.Id, and Component.Version for Application Insights.
/// This enables the Users and Sessions blades in the Azure Portal and ensures correct version reporting.
/// </summary>
public sealed class ExcelMcpTelemetryInitializer : ITelemetryInitializer
{
    private readonly string _userId;
    private readonly string _sessionId;
    private readonly string _version;
    private readonly string _roleInstance;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelMcpTelemetryInitializer"/> class.
    /// </summary>
    public ExcelMcpTelemetryInitializer()
    {
        _userId = ExcelMcpTelemetry.UserId;
        _sessionId = ExcelMcpTelemetry.SessionId;
        _version = GetVersion();
        _roleInstance = GenerateAnonymousRoleInstance();
    }

    /// <summary>
    /// Initializes the telemetry item with user, session, and version context.
    /// </summary>
    /// <param name="telemetry">The telemetry item to initialize.</param>
    public void Initialize(ITelemetry telemetry)
    {
        // Set user context for Users blade
        if (string.IsNullOrEmpty(telemetry.Context.User.Id))
        {
            telemetry.Context.User.Id = _userId;
        }

        // Set session context for Sessions blade
        if (string.IsNullOrEmpty(telemetry.Context.Session.Id))
        {
            telemetry.Context.Session.Id = _sessionId;
        }

        // Set cloud role for better grouping in Application Map
        if (string.IsNullOrEmpty(telemetry.Context.Cloud.RoleName))
        {
            telemetry.Context.Cloud.RoleName = "ExcelMcp.McpServer";
        }

        // Set role instance to anonymized value (instead of machine name)
        telemetry.Context.Cloud.RoleInstance = _roleInstance;

        // Set version explicitly - ALWAYS override SDK auto-detection
        // SDK picks up Excel COM version (15.0.0.0) instead of our assembly version
        telemetry.Context.Component.Version = _version;
    }

    /// <summary>
    /// Generates an anonymous role instance identifier based on machine identity.
    /// Uses the same hash as UserId but with a different prefix for clarity.
    /// </summary>
    private static string GenerateAnonymousRoleInstance()
    {
        // Reuse the anonymous user ID (already a hash of machine identity)
        return $"instance-{ExcelMcpTelemetry.UserId[..8]}";
    }

    /// <summary>
    /// Gets the application version from assembly metadata.
    /// </summary>
    private static string GetVersion()
    {
        return Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString()
            ?? "1.0.0";
    }
}


