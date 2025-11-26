// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.Extensibility;

namespace Sbroenne.ExcelMcp.McpServer.Telemetry;

/// <summary>
/// Telemetry initializer that sets User.Id and Session.Id for Application Insights.
/// This enables the Users and Sessions blades in the Azure Portal.
/// </summary>
public sealed class ExcelMcpTelemetryInitializer : ITelemetryInitializer
{
    private readonly string _userId;
    private readonly string _sessionId;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelMcpTelemetryInitializer"/> class.
    /// </summary>
    public ExcelMcpTelemetryInitializer()
    {
        _userId = ExcelMcpTelemetry.UserId;
        _sessionId = ExcelMcpTelemetry.SessionId;
    }

    /// <summary>
    /// Initializes the telemetry item with user and session context.
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
    }
}
