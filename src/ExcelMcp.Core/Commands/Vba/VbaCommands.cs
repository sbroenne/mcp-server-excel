using System.Runtime.InteropServices;
using Microsoft.Win32;
using Sbroenne.ExcelMcp.ComInterop;

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
    internal static bool IsVbaTrustEnabled()
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
    /// Determines whether a <see cref="COMException"/> genuinely represents the
    /// "programmatic access to the VBA project object model is not trusted" failure.
    /// </summary>
    /// <remarks>
    /// HRESULT <c>0x800A03EC</c> is the <em>generic</em> Office automation error
    /// ("Exception occurred") and is reused for many unrelated failures. Treating every
    /// <c>0x800A03EC</c> as a trust error masks real problems and misdirects users to
    /// re-check Trust Center settings that are already correct (see issue #671). We only
    /// classify it as a trust error when the registry confirms trust is actually disabled,
    /// which is locale-independent and therefore also correct on non-English Office builds
    /// where the COM message text differs. The English message check remains as a fast path.
    /// </remarks>
    internal static bool IsVbaTrustError(COMException comEx)
    {
        if (comEx.Message.Contains("programmatic access", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return comEx.ErrorCode == unchecked((int)0x800A03EC) && !IsVbaTrustEnabled();
    }

    /// <summary>
    /// HRESULT for the generic Office automation error (DISP_E_EXCEPTION, "Exception occurred").
    /// </summary>
    internal const int GenericOfficeAutomationError = unchecked((int)0x800A03EC);

    /// <summary>
    /// Builds an enriched error message for a genuine (non-trust) generic Office automation
    /// failure (HRESULT <c>0x800A03EC</c>).
    /// </summary>
    /// <remarks>
    /// After issue #671, a <c>0x800A03EC</c> raised while VBA trust is enabled is no longer
    /// masked as a trust error - the real failure is surfaced. The raw <see cref="COMException"/>
    /// message is often opaque (sometimes just the HRESULT), so this attaches the COM
    /// environment diagnostics (Office channel/version, bitness, CLSID, registered PIA) that the
    /// project already uses for other <c>0x800A03EC</c>-class failures, giving the first re-run
    /// maximum triage information for the still-unknown underlying cause.
    /// </remarks>
    internal static string BuildGenericComErrorMessage(COMException comEx)
    {
        string hr = $"0x{unchecked((uint)comEx.ErrorCode):X8}";
        string detail = string.IsNullOrWhiteSpace(comEx.Message) ? "(no description)" : comEx.Message;
        string diagnostics = ComDiagnostics.FormatForErrorMessage(ComDiagnostics.Collect());

        return $"VBA operation failed with a generic Office automation error (HRESULT {hr}: {detail}). " +
               "This is NOT a VBA trust problem - 'Trust access to the VBA project object model' is enabled. " +
               $"The underlying cause is environment-specific (see issue #671).{Environment.NewLine}{diagnostics}";
    }

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


