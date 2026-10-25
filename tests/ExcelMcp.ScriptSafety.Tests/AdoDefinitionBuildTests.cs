using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class AdoDefinitionBuildTests
{
    [Theory]
    [InlineData("-ImporterPath '.\\missing-ado-importer.exe'", "TlbImp.exe is missing")]
    [InlineData("-TypeLibraryPath '.\\missing-ado-definitions.tlb'", "Windows ADO 2.8 definitions are missing")]
    [InlineData("-TypeLibraryPath '.\\Directory.Build.targets'", "A valid Microsoft signature is required")]
    [InlineData("-ImporterPath (Join-Path $env:CommonProgramFiles 'System\\ado\\msado15.dll')", "Expected Microsoft TlbImp.exe 4.8.9037.0")]
    public async Task Import_RejectsMissingAndUnsignedInputs(string arguments, string expected)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            try {
                & .\scripts\Import-AdoDefinitions.ps1 -ValidateOnly {{arguments}}
                throw 'Invalid import inputs were accepted.'
            } catch {
                if ($_.Exception.Message -notlike '*{{expected}}*') { throw }
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }

    [Fact]
    public async Task RepeatedImports_ProduceIdenticalEmbeddedApplicationOutput()
    {
        string directory = Path.Combine(Path.GetTempPath(), $"ExcelMcp-AdoDefinitions-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            string first = Path.Combine(directory, "first", "ADODB.dll");
            string second = Path.Combine(directory, "second", "ADODB.dll");
            var run = await ValidationSelectionTests.RunAsync($$"""
                & .\scripts\Import-AdoDefinitions.ps1 -OutputPath '{{first.Replace("'", "''", StringComparison.Ordinal)}}'
                & .\scripts\Import-AdoDefinitions.ps1 -OutputPath '{{second.Replace("'", "''", StringComparison.Ordinal)}}'
                """);
            Assert.True(run.ExitCode == 0, run.Output);
            Assert.DoesNotContain("warning TI", run.Output, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(Emit(first), Emit(second));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }

        static byte[] Emit(string definitions)
        {
            var compilation = SafetyAnalyzerTests.CreateCompilation("""
                class C {
                    object Execute(ADODB._Connection connection, string query) =>
                        connection.Execute(query, out _, (int)ADODB.ExecuteOptionEnum.adOptionUnspecified);
                }
                """).WithOptions(new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, deterministic: true))
                .AddReferences(MetadataReference.CreateFromFile(definitions,
                    new MetadataReferenceProperties(embedInteropTypes: true)));
            using var output = new MemoryStream();
            var emitted = compilation.Emit(output);
            Assert.True(emitted.Success, string.Join(Environment.NewLine, emitted.Diagnostics));
            return output.ToArray();
        }
    }

    [Theory]
    [InlineData("Import-AdoDefinitions.ps1")]
    [InlineData("Install-AdoBuildTools.ps1")]
    public async Task BuildInputs_SelectCompilationNativeAndOutputChecks(string file)
    {
        var run = await ValidationSelectionTests.RunAsync($$"""
            $plan = Get-ValidationPlan -Paths 'scripts/{{file}}'
            if (-not $plan.FullSolutionBuild -or
                'ScriptSafety' -notin $plan.ToolingProjects -or
                -not $plan.Excel -or
                'Cli' -notin $plan.PackageComponents -or
                'Mcp' -notin $plan.PackageComponents) {
                throw 'ADO build inputs lost compilation, native, or deployed-output coverage.'
            }
            """);
        Assert.True(run.ExitCode == 0, run.Output);
    }
}
