namespace Sbroenne.ExcelMcp.CLI.Telemetry;

/// <summary>
/// Who is reporting: the role name, role instance and application version that
/// Application Insights shows for every record. One value is shared by the sink,
/// which sends it as resource attributes, and by the telemetry items themselves.
/// </summary>
internal sealed record TelemetryIdentity(string RoleName, string RoleInstance, string Version);
