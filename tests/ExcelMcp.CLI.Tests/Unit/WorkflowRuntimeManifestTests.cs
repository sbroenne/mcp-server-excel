using System.Text.RegularExpressions;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "Workflow")]
[Trait("Speed", "Fast")]
public sealed partial class WorkflowRuntimeManifestTests
{
    [Fact]
    public void CreateWorkflowRuntimeManifest_IdentifiesTheCliHostAndWorkflowSurface()
    {
        var manifest = Sbroenne.ExcelMcp.CLI.Program.CreateWorkflowRuntimeManifest();

        Assert.Equal("excelcli", manifest.ServerName);
        Assert.Equal("cli", manifest.ToolProfile);
        Assert.Equal(["workflow"], manifest.ToolProfileTools);
        Assert.Matches(ProfileManifestHashPattern(), manifest.ToolProfileManifestHash);
    }

    [GeneratedRegex("^[0-9a-f]{64}$", RegexOptions.CultureInvariant)]
    private static partial Regex ProfileManifestHashPattern();
}
