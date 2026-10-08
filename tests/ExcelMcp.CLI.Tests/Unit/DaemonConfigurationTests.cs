using System.Xml.Linq;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "CLI")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DaemonConfigurationTests
{
    [Fact]
    public void MutexNames_UseDisjointHashedNamespaces()
    {
        const string daemonPrefix = "ExcelMcpCli_Daemon_";
        const string startupPrefix = "ExcelMcpCli_Startup_";
        const string trackerPrefix = "ExcelMcpCli_Tracker_";
        var pipeNames = GetAdversarialPipeNames();
        var names = pipeNames
            .SelectMany(pipeName => new[]
            {
                DaemonAutoStart.GetDaemonMutexName(pipeName),
                DaemonAutoStart.GetDaemonStartupLockName(pipeName),
                DaemonProcessTracker.GetTrackingMutexName(pipeName)
            })
            .ToList();

        Assert.Equal(
            pipeNames.Distinct(StringComparer.OrdinalIgnoreCase).Count() * 3,
            names.Distinct(StringComparer.Ordinal).Count());
        foreach (var pipeName in pipeNames)
        {
            var daemonName = DaemonAutoStart.GetDaemonMutexName(pipeName);
            var startupName = DaemonAutoStart.GetDaemonStartupLockName(pipeName);
            var trackerName = DaemonProcessTracker.GetTrackingMutexName(pipeName);
            Assert.StartsWith(daemonPrefix, daemonName, StringComparison.Ordinal);
            Assert.StartsWith(startupPrefix, startupName, StringComparison.Ordinal);
            Assert.StartsWith(trackerPrefix, trackerName, StringComparison.Ordinal);
            Assert.Equal(64, daemonName[daemonPrefix.Length..].Length);
            Assert.Equal(64, startupName[startupPrefix.Length..].Length);
            Assert.Equal(64, trackerName[trackerPrefix.Length..].Length);
            Assert.All(daemonName[daemonPrefix.Length..], character =>
                Assert.True(char.IsAsciiHexDigit(character)));
            Assert.All(startupName[startupPrefix.Length..], character =>
                Assert.True(char.IsAsciiHexDigit(character)));
            Assert.All(trackerName[trackerPrefix.Length..], character =>
                Assert.True(char.IsAsciiHexDigit(character)));
        }

        Assert.Equal(
            DaemonAutoStart.GetDaemonMutexName("foo"),
            DaemonAutoStart.GetDaemonMutexName("FOO"));
        Assert.Equal(
            DaemonAutoStart.GetDaemonStartupLockName("foo"),
            DaemonAutoStart.GetDaemonStartupLockName("FOO"));
        Assert.Equal(
            DaemonProcessTracker.GetTrackingMutexName("foo"),
            DaemonProcessTracker.GetTrackingMutexName("FOO"));
        Assert.Equal(
            DaemonProcessTracker.GetTrackingFilePath("foo"),
            DaemonProcessTracker.GetTrackingFilePath("FOO"));
    }

    [Fact]
    public void BuildServiceStop_RunsOnlyForCliProject()
    {
        var buildProperties = XDocument.Load(
            Path.Combine(GetRepositoryRoot(), "Directory.Build.props"));
        var cleanupTarget = buildProperties
            .Descendants("Target")
            .Single(element => string.Equals(
                element.Attribute("Name")?.Value,
                "StopExcelCliService",
                StringComparison.Ordinal));

        Assert.Contains(
            "'$(MSBuildProjectName)' == 'ExcelMcp.CLI'",
            cleanupTarget.Attribute("Condition")?.Value,
            StringComparison.Ordinal);
        Assert.Equal("BeforeBuild", cleanupTarget.Attribute("BeforeTargets")?.Value);
    }

    [Fact]
    public void DevelopmentStop_HasNoBootstrapOrProductShutdownPath()
    {
        var cleanupScript = File.ReadAllText(
            Path.Combine(GetRepositoryRoot(), "scripts", "Stop-ExcelCliService.ps1"));
        var forbiddenFragments = new[]
        {
            "dotnet",
            "service stop",
            "service.shutdown",
            "EXCEL.EXE",
            "DaemonProcessTracker",
            "ExcelMcp.Cleanup",
            "LastWriteTime"
        };
        Assert.All(forbiddenFragments, fragment =>
            Assert.DoesNotContain(fragment, cleanupScript, StringComparison.OrdinalIgnoreCase));
        Assert.False(File.Exists(Path.Combine(GetRepositoryRoot(), "src", "ExcelMcp.Cleanup", "ExcelMcp.Cleanup.csproj")));
    }

    private static string GetRepositoryRoot() =>
        Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "..",
            ".."));

    private static IReadOnlyList<string> GetAdversarialPipeNames() =>
    [
        "foo",
        "FOO",
        "foo_startup",
        @"pipe/segment\with:separators?and spaces",
        $"excelmcp-{new string('x', 4096)}"
    ];
}
