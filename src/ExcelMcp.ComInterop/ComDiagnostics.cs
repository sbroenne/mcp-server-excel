using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Gathers COM environment diagnostics for troubleshooting Excel interop failures.
/// All methods are safe to call without creating an Excel instance.
/// </summary>
public static class ComDiagnostics
{
    /// <summary>
    /// Collects diagnostic information about the Excel COM registration and runtime environment.
    /// </summary>
    public static ComDiagnosticReport Collect()
    {
        var report = new ComDiagnosticReport
        {
            ProcessArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
            OsArchitecture = RuntimeInformation.OSArchitecture.ToString(),
            RuntimeVersion = RuntimeInformation.FrameworkDescription,
            CollectedAtUtc = DateTime.UtcNow
        };

        // Resolve ProgID → CLSID
        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        report.ProgIdResolved = excelType != null;
        report.ResolvedClsid = excelType?.GUID.ToString("B");

        // Check PIA interface GUID (what the cast targets)
        Type piaType = typeof(Microsoft.Office.Interop.Excel.Application);
        report.PiaInterfaceGuid = piaType.GUID.ToString("B");
        report.PiaAssemblyName = piaType.Assembly.GetName().Name;
        report.PiaAssemblyVersion = piaType.Assembly.GetName().Version?.ToString();

        // Probe registry for Office installation details
        report.OfficeRegistration = ProbeOfficeRegistration();

        return report;
    }

    /// <summary>
    /// Formats a diagnostic report as a concise string for inclusion in error messages.
    /// </summary>
    public static string FormatForErrorMessage(ComDiagnosticReport report)
    {
        var sb = new StringBuilder();
        sb.AppendLine("COM Diagnostics:");
        sb.Append("  ProgID resolved: ").AppendLine(report.ProgIdResolved ? "yes" : "NO");
        sb.Append("  CLSID: ").AppendLine(report.ResolvedClsid ?? "(null)");
        sb.Append("  PIA interface: ").AppendLine(report.PiaInterfaceGuid);
        sb.Append("  PIA assembly: ").Append(report.PiaAssemblyName ?? "(null)").Append(' ').AppendLine(report.PiaAssemblyVersion ?? "");
        sb.Append("  Process arch: ").Append(report.ProcessArchitecture)
          .Append(", OS arch: ").AppendLine(report.OsArchitecture);

        if (report.OfficeRegistration != null)
        {
            sb.Append("  Office: ").AppendLine(report.OfficeRegistration);
        }

        return sb.ToString();
    }

    private static string? ProbeOfficeRegistration()
    {
        // Check Click-to-Run registration (most common modern Office install)
        string[] registryPaths =
        [
            @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            @"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
        ];

        foreach (var path in registryPaths)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(path);
                if (key != null)
                {
                    var platform = key.GetValue("Platform")?.ToString();
                    var version = key.GetValue("VersionToReport")?.ToString();
                    var channel = key.GetValue("CDNBaseUrl")?.ToString();

                    if (version != null)
                    {
                        // Extract channel name from CDN URL
                        string? channelName = channel switch
                        {
                            string s when s.Contains("Monthly", StringComparison.OrdinalIgnoreCase) => "Monthly",
                            string s when s.Contains("SemiAnnual", StringComparison.OrdinalIgnoreCase) => "Semi-Annual",
                            string s when s.Contains("Current", StringComparison.OrdinalIgnoreCase) => "Current",
                            string s when s.Contains("Insiders", StringComparison.OrdinalIgnoreCase) => "Insiders",
                            _ => null
                        };

                        return $"Click-to-Run {version} ({platform ?? "unknown"} arch)" +
                               (channelName != null ? $" [{channelName} channel]" : "");
                    }
                }
            }
            catch
            {
                // Registry access may fail — continue probing
            }
        }

        return null;
    }
}

/// <summary>
/// Diagnostic information about the Excel COM environment.
/// </summary>
public sealed class ComDiagnosticReport
{
    /// <summary>Whether Type.GetTypeFromProgID("Excel.Application") resolved.</summary>
    public bool ProgIdResolved { get; set; }

    /// <summary>The CLSID that "Excel.Application" resolved to.</summary>
    public string? ResolvedClsid { get; set; }

    /// <summary>The interface GUID that the PIA cast targets.</summary>
    public string? PiaInterfaceGuid { get; set; }

    /// <summary>PIA assembly name.</summary>
    public string? PiaAssemblyName { get; set; }

    /// <summary>PIA assembly version.</summary>
    public string? PiaAssemblyVersion { get; set; }

    /// <summary>Process architecture (x64, x86, Arm64).</summary>
    public string? ProcessArchitecture { get; set; }

    /// <summary>OS architecture.</summary>
    public string? OsArchitecture { get; set; }

    /// <summary>.NET runtime version.</summary>
    public string? RuntimeVersion { get; set; }

    /// <summary>Office installation details from registry.</summary>
    public string? OfficeRegistration { get; set; }

    /// <summary>When the report was collected.</summary>
    public DateTime CollectedAtUtc { get; set; }
}
