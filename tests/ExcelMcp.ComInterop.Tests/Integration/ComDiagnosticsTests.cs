using System.Runtime.InteropServices;
using Microsoft.Win32;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

/// <summary>
/// Integration tests for ComDiagnostics — verifies the diagnostic helper
/// produces correct environment information on a machine with Excel installed.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Diagnostics")]
[Trait("Layer", "ComInterop")]
[Collection("Sequential")]
[Trait("RequiresExcel", "true")]
public sealed class ComDiagnosticsTests
{
    [Fact]
    public void Collect_OnMachineWithExcel_ReturnsValidReport()
    {
        var before = DateTime.UtcNow;
        var report = ComDiagnostics.Collect();

        Assert.True(report.ProgIdResolved, "Excel.Application ProgID should resolve on a machine with Excel");
        using var clsid = Registry.ClassesRoot.OpenSubKey(@"Excel.Application\CLSID");
        var expectedClsid = Assert.IsType<string>(clsid?.GetValue(null));
        Assert.Equal(Guid.Parse(expectedClsid), Guid.Parse(report.ResolvedClsid!));
        Assert.Equal("{000208d5-0000-0000-c000-000000000046}", report.PiaInterfaceGuid);
        Assert.Equal(RuntimeInformation.ProcessArchitecture.ToString(), report.ProcessArchitecture);
        Assert.Equal(RuntimeInformation.OSArchitecture.ToString(), report.OsArchitecture);
        Assert.Equal(RuntimeInformation.FrameworkDescription, report.RuntimeVersion);
        Assert.InRange(report.CollectedAtUtc, before, DateTime.UtcNow);
    }

    [Fact]
    public void Collect_ReturnsClickToRunDetails_WhenPresent()
    {
        var report = ComDiagnostics.Collect();

        using var native = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
        using var redirected = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration");
        var registration = native?.GetValue("VersionToReport") is not null ? native : redirected;
        if (registration?.GetValue("VersionToReport") is string version)
        {
            var platform = registration.GetValue("Platform")?.ToString() ?? "unknown";
            Assert.Contains($"Click-to-Run {version} ({platform} arch)", report.OfficeRegistration);
        }
        else
            Assert.Null(report.OfficeRegistration);
    }

    [Fact]
    public void FormatForErrorMessage_ProducesReadableOutput()
    {
        var report = ComDiagnostics.Collect();
        var formatted = ComDiagnostics.FormatForErrorMessage(report);

        Assert.Contains("COM Diagnostics:", formatted);
        Assert.Contains($"ProgID resolved: {(report.ProgIdResolved ? "yes" : "NO")}", formatted);
        Assert.Contains($"CLSID: {report.ResolvedClsid}", formatted);
        Assert.Contains($"PIA interface: {report.PiaInterfaceGuid}", formatted);
        Assert.Contains($"PIA assembly: {report.PiaAssemblyName} {report.PiaAssemblyVersion}", formatted);
        Assert.Contains($"Process arch: {report.ProcessArchitecture}, OS arch: {report.OsArchitecture}", formatted);
    }

    [Fact]
    public void Collect_PiaAssemblyInfo_IsPopulated()
    {
        var report = ComDiagnostics.Collect();

        // Excel interfaces are embedded into the production assembly, not loaded from a runtime PIA.
        var embeddedAssembly = typeof(ComDiagnostics).Assembly.GetName();
        Assert.Equal(embeddedAssembly.Name, report.PiaAssemblyName);
        Assert.Equal(embeddedAssembly.Version?.ToString(), report.PiaAssemblyVersion);
    }

    [Fact]
    public void Collect_RegisteredPiaAssemblyName_IsWellFormed_WhenRegistered()
    {
        var report = ComDiagnostics.Collect();

        using var interfaceKey = Registry.ClassesRoot.OpenSubKey(
            @"Interface\{000208D5-0000-0000-C000-000000000046}\TypeLib");
        var id = interfaceKey?.GetValue(null)?.ToString();
        var version = interfaceKey?.GetValue("Version")?.ToString();
        Assert.Equal(id, report.ExcelTypeLibId);
        Assert.Equal(version, report.ExcelTypeLibVersion);
        using var library = id is not null && version is not null
            ? Registry.ClassesRoot.OpenSubKey($@"TypeLib\{id}\{version}") : null;
        Assert.Equal(library?.GetValue("PrimaryInteropAssemblyName")?.ToString(),
            report.ExcelTypeLibPrimaryInteropAssemblyName);
    }
}
