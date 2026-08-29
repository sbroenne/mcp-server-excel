namespace Sbroenne.ExcelMcp.Core.Attributes;

/// <summary>
/// Excludes a compatibility-only interface method from generated Service, CLI, and MCP surfaces.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
public sealed class ServiceIgnoreAttribute : Attribute;
