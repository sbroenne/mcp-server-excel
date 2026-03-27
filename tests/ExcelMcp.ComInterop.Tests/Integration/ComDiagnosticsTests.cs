using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

/// <summary>
/// Integration tests for ComDiagnostics — verifies the diagnostic helper
/// produces correct environment information on a machine with Excel installed.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Diagnostics")]
[Trait("Layer", "ComInterop")]
[Trait("RequiresExcel", "true")]
public sealed class ComDiagnosticsTests
{
    [Fact]
    public void Collect_OnMachineWithExcel_ReturnsValidReport()
    {
        var report = ComDiagnostics.Collect();

        Assert.True(report.ProgIdResolved, "Excel.Application ProgID should resolve on a machine with Excel");
        Assert.NotNull(report.ResolvedClsid);
        Assert.NotNull(report.PiaInterfaceGuid);
        Assert.NotNull(report.ProcessArchitecture);
        Assert.NotNull(report.OsArchitecture);
        Assert.NotNull(report.RuntimeVersion);
    }

    [Fact]
    public void Collect_ReturnsClickToRunDetails_WhenPresent()
    {
        var report = ComDiagnostics.Collect();

        // On CI/dev machines with Click-to-Run Office, this should be populated
        // On MSI installs it may be null — that's OK, just verify the field exists
        if (report.OfficeRegistration != null)
        {
            Assert.Contains("Click-to-Run", report.OfficeRegistration);
        }
    }

    [Fact]
    public void FormatForErrorMessage_ProducesReadableOutput()
    {
        var report = ComDiagnostics.Collect();
        var formatted = ComDiagnostics.FormatForErrorMessage(report);

        Assert.Contains("COM Diagnostics:", formatted);
        Assert.Contains("ProgID resolved:", formatted);
        Assert.Contains("CLSID:", formatted);
        Assert.Contains("PIA interface:", formatted);
        Assert.Contains("PIA assembly:", formatted);
        Assert.Contains("Process arch:", formatted);
    }

    [Fact]
    public void Collect_PiaAssemblyInfo_IsPopulated()
    {
        var report = ComDiagnostics.Collect();

        Assert.NotNull(report.PiaAssemblyName);
        Assert.NotNull(report.PiaAssemblyVersion);
        Assert.Contains("Interop", report.PiaAssemblyName);
    }
}
