using Microsoft.Win32;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands - Core data layer (no console output)
/// </summary>
public partial class VbaCommands : IVbaCommands
{
    // CA1861: Use static readonly for constant array arguments to avoid repeated allocations
    private static readonly string[] RegistryPaths =
    [
        @"Software\Microsoft\Office\16.0\Excel\Security",  // Office 2019/2021/365
        @"Software\Microsoft\Office\15.0\Excel\Security",  // Office 2013
        @"Software\Microsoft\Office\14.0\Excel\Security"   // Office 2010
    ];

    // CA1861: Use static readonly for constant array arguments to avoid repeated allocations
    private static readonly char[] ProcedureSeparators = [' ', '('];

    /// <summary>
    /// Check if VBA trust is enabled by reading registry
    /// </summary>
    private static bool IsVbaTrustEnabled()
    {
        try
        {
            // Try different Office versions
            foreach (string path in RegistryPaths)
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(path);
                    var value = key?.GetValue("AccessVBOM");
                    if (value != null && (int)value == 1)
                    {
                        return true;
                    }
                }
                catch (System.Security.SecurityException)
                {
                    // Registry access denied for this path, try next Office version
                }
            }

            return false; // Assume not enabled if cannot read registry
        }
        catch (System.Security.SecurityException)
        {
            // Registry access completely denied, assume VBA trust not enabled
            return false;
        }
    }

    private const string VbaTrustErrorMessage = "VBA trust access is not enabled. Enable 'Trust access to the VBA project object model' in Excel Trust Center settings.";

    /// <summary>
    /// Validate that file is macro-enabled (.xlsm) for VBA operations
    /// </summary>
    private static (bool IsValid, string? ErrorMessage) ValidateVbaFile(string filePath)
    {
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsm")
        {
            return (false, $"VBA operations require macro-enabled workbooks (.xlsm). Current file has extension: {extension}");
        }
        return (true, null);
    }

    private static string ExtractProcedureName(string codeLine)
    {
        var parts = codeLine.Trim().Split(ProcedureSeparators, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < parts.Length; i++)
        {
            if ((parts[i] is "Sub" or "Function") && i + 1 < parts.Length)
            {
                return parts[i + 1];
            }
        }
        return string.Empty;
    }

    /// <summary>
    /// Converts VBA component type constant to display name
    /// </summary>
    private static string GetVbaModuleTypeName(int componentType)
    {
        return componentType switch
        {
            1 => "Module",
            2 => "Class",
            3 => "Form",
            100 => "Document",
            _ => $"Type{componentType}"
        };
    }
}
