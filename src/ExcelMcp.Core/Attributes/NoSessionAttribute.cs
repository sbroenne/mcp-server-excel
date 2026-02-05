namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Marks an interface as not requiring a session.
/// Used for session management interfaces (file/session commands)
/// where operations create/destroy sessions rather than use them.
/// </summary>
/// <remarks>
/// When applied:
/// - MCP tool gets [McpMeta("requiresSession", false)]
/// - CLI commands don't require --session parameter
/// - Service handlers don't validate sessionId
/// </remarks>
[AttributeUsage(AttributeTargets.Interface, AllowMultiple = false, Inherited = false)]
public sealed class NoSessionAttribute : Attribute
{
}
