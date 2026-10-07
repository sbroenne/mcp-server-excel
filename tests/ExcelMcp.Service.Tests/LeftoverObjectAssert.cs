using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Checks that a failure after Excel created an object tells the caller which object remains and where.
/// </summary>
internal static class LeftoverObjectAssert
{
    public static void Reported(string message, string objectKind, string objectName, string sheetName)
    {
        Assert.Contains($"{objectKind} '{objectName}'", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains($"sheet '{sheetName}'", message, StringComparison.Ordinal);
        Assert.Contains("remains", message, StringComparison.OrdinalIgnoreCase);
    }
}
