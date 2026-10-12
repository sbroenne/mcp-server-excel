using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("Feature", "PreCommit")]
[Trait("Feature", "AutomationSafety")]
[Trait("RequiresExcel", "false")]
public sealed class TypedSourceGuardTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"ExcelMcp.SourceGuards.{Guid.NewGuid():N}");
    private const string Commands = @"src\ExcelMcp.Core\Commands";

    public TypedSourceGuardTests()
    {
        Directory.CreateDirectory(Path.Combine(_root, Commands));
        Directory.CreateDirectory(Path.Combine(_root, "src", "ExcelMcp.ComInterop"));
    }

    [Theory]
    [InlineData("var count = range.Rows.Count;", "chained COM property access")]
    [InlineData("var row = table.Range.Row;", "chained COM property access")]
    [InlineData("var column = table.Range.Column;", "chained COM property access")]
    [InlineData("namesCollection.Add(\"Example\", \"=Sheet1!A1\");", "discarded Names.Add result")]
    public void ComGuard_ReportsEachOriginalUnsafeFixture(string source, string rule)
    {
        var path = Write(Path.Combine(Commands, "Unsafe.cs"), source);
        var finding = Assert.Single(new SourceGuards(_root).Scan("com-leaks", [path]));
        Assert.Equal(rule, finding.Rule);
        Assert.Equal(1, finding.Line);
        Assert.Equal(Path.Combine(Commands, "Unsafe.cs"), finding.Path);
        Assert.Contains("finally", finding.Guidance, StringComparison.Ordinal);
    }

    [Fact]
    public void ComGuard_AllowsOriginalSafeFixtureAndCommentLines()
    {
        var path = Write(Path.Combine(Commands, "Safe.cs"), """
            dynamic? rows = null;
            try
            {
                rows = range.Rows;
                int count = rows.Count;
            }
            finally
            {
                ComUtilities.Release(ref rows);
            }
            // var count = range.Rows.Count;
            * range.Rows.Count
            """);
        Assert.Empty(new SourceGuards(_root).Scan("com-leaks", [path]));
    }

    [Theory]
    [InlineData("result.Success = true;\nresult.ErrorMessage = \"failure\";", 1)]
    [InlineData("result.Success = true;\nresult.Success = false;\nresult.ErrorMessage = \"failure\";", 0)]
    [InlineData("result.Success = true;\nresult.ErrorMessage = \"\";", 0)]
    [InlineData("result.Success = true;\nresult.ErrorMessage = string.Empty;", 0)]
    [InlineData("result.Success = true;\nresult.ErrorMessage = null;", 0)]
    [InlineData("result.Success = true;\nother.Success = true;\nresult.ErrorMessage = \"failure\";", 1)]
    public void SuccessGuard_PreservesNearbyAssignmentBoundary(string source, int expected)
    {
        Write(Path.Combine(Commands, "Cases.cs"), source);
        var findings = new SourceGuards(_root).Scan("success-flag");
        Assert.Equal(expected, findings.Count);
        Assert.All(findings, finding => Assert.Contains("not control-flow analysis", finding.Guidance, StringComparison.Ordinal));
    }

    [Fact]
    public void SuccessGuard_DoesNotLookBeyondOriginalThirtyLineWindow()
    {
        Write(Path.Combine(Commands, "Cases.cs"),
            "result.Success = true;\n" + new string('\n', 29) + "result.ErrorMessage = \"failure\";");
        Assert.Empty(new SourceGuards(_root).Scan("success-flag"));
    }

    [Theory]
    [InlineData("// PIA gap: unavailable\nvar value = ((dynamic)source).Value;", 0)]
    [InlineData("// TODO: tracked\n\n// context\nvar value = ((dynamic)source).Value;", 0)]
    [InlineData("// REASON: unavailable\nvar value = ((dynamic)source).Value;", 0)]
    [InlineData("// Reason: unavailable\nvar intervening = 1;\nvar value = ((dynamic)source).Value;", 1)]
    [InlineData("// PIA gap: unavailable\n\n\n\n\n\nvar value = ((dynamic)source).Value;", 1)]
    [InlineData("// var value = ((dynamic)source).Value;", 0)]
    public void DynamicGuard_PreservesJustificationWindow(string source, int expected)
    {
        Write(Path.Combine(Commands, "Cases.cs"), source);
        Write(@"src\ExcelMcp.ComInterop\Cases.cs", "class Cases {}");
        Assert.Equal(expected, new SourceGuards(_root).Scan("dynamic-casts").Count);
    }

    [Fact]
    public void DynamicGuard_PreservesNamedInfrastructureExceptions()
    {
        Write(Path.Combine(Commands, "Cases.cs"), "class Cases {}");
        Write(@"src\ExcelMcp.ComInterop\ExcelSession.cs", "var value = ((dynamic)source).Value;");
        Assert.Empty(new SourceGuards(_root).Scan("dynamic-casts"));
    }

    [Theory]
    [InlineData("com-leaks")]
    [InlineData("success-flag")]
    [InlineData("dynamic-casts")]
    public void Guards_RejectEmptyOrGeneratedOnlyInputs(string rule)
    {
        Write(Path.Combine(Commands, "Generated.g.cs"), "var count = range.Rows.Count;");
        Write(@"src\ExcelMcp.ComInterop\obj\Generated.cs", "var value = ((dynamic)source).Value;");
        var guard = new SourceGuards(_root);
        Assert.Throws<InvalidOperationException>(() =>
            guard.Scan(rule, rule == "com-leaks" ? [Path.Combine(_root, Commands)] : null));
    }

    [Fact]
    public void PackageGuard_ReportsProductionProjectDependenciesAndIgnoresTestsAndObj()
    {
        Write(@"src\Runtime\Runtime.csproj", "<PackageReference Include=\"DocumentFormat.OpenXml\" />");
        Write(@"tests\Fixture.cs", "const string Part = \"xl/workbook.xml\";");
        Write(@"src\Runtime\obj\Generated.cs", "const string Part = \"xl/workbook.xml\";");
        var finding = Assert.Single(new SourceGuards(_root).Scan("workbook-package-access"));
        Assert.Equal(@"src\Runtime\Runtime.csproj", finding.Path);
        Assert.Equal("Open XML or package API", finding.Rule);
    }

    private string Write(string relative, string content)
    {
        var path = Path.Combine(_root, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        return path;
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);
}
