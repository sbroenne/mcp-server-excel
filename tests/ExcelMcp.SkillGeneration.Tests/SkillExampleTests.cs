using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "SkillGeneration")]
public sealed class SkillExampleTests
{
    private static readonly string RepoRoot = FindRepoRoot();

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    public async Task SkillGeneration_SelectsNativeExamplesWithoutTranslatingContent(string newline)
    {
        var root = NewSandbox();
        try
        {
            var script = Path.Combine(RepoRoot, "scripts", "Build-AgentSkills.ps1");
            var result = await RunAsync(root, $$"""
                $ast = [Management.Automation.Language.Parser]::ParseFile('{{Quote(script)}}', [ref]$null, [ref]$null)
                $function = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Copy-SharedReferences' }, $true)
                if (-not $function) { throw 'Missing shared-reference renderer.' }
                . ([scriptblock]::Create($function.Extent.Text))
                $SharedDir = New-Item -ItemType Directory -Path shared
                $document = @'
                # Workflow
                Shared policy uses sessionId responses.
                ```cli
                excelcli -q range get-values --session $sessionId --sheet Sales --range A1
                ```
                ```mcp
                range(action: 'get-values', workbook_session_id: sessionId, sheet_name: 'Sales', range_address: 'A1')
                ```
                ```json
                {"mCodeFile":"query.m"}
                ```
                '@
                $newline = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{{Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(newline))}}'))
                [IO.File]::WriteAllText((Join-Path $SharedDir 'report-formatting.md'), (($document -replace "`r`n?", "`n") -replace "`n", $newline))
                foreach ($surface in @('cli', 'mcp')) {
                    Copy-SharedReferences -SkillPath "excel-$surface-report-formatting" -Surface $surface
                    $content = Get-Content -LiteralPath "excel-$surface-report-formatting\references\report-formatting.md" -Raw
                    if ($content -notmatch 'Shared policy uses sessionId responses.' -or
                        $content -notmatch '\{"mCodeFile":"query.m"\}' -or
                        $content -match '(?m)^```(?:cli|mcp)$') { throw 'Shared content was changed or fences were not rendered.' }
                    if ($surface -eq 'cli') {
                        if ($content -notmatch '```powershell' -or $content -notmatch 'excelcli -q range' -or
                            $content -match 'range\(action:') { throw 'CLI examples were not selected.' }
                    } else {
                        if ($content -notmatch '```text' -or $content -notmatch "workbook_session_id: sessionId" -or
                            $content -match 'excelcli -q') { throw 'MCP examples were not selected.' }
                    }
                }
                """);
            Assert.True(result.ExitCode == 0, result.Output);
        }
        finally { Directory.Delete(root, true); }
    }

    private static string NewSandbox()
        => Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.Automation.{Guid.NewGuid():N}")).FullName;

    private static string Quote(string value) => value.Replace("'", "''", StringComparison.Ordinal);

    private static async Task<(int ExitCode, string Output)> RunAsync(string root, string body)
    {
        var script = Path.Combine(root, $"test-{Guid.NewGuid():N}.ps1");
        File.WriteAllText(script, "$ErrorActionPreference='Stop'\n$global:LASTEXITCODE=0\n" + body + "\nexit $LASTEXITCODE");
        var info = new ProcessStartInfo("pwsh")
        {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
            WorkingDirectory = root
        };
        foreach (var argument in new[] { "-NoProfile", "-File", script }) { info.ArgumentList.Add(argument); }
        using var process = Process.Start(info)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        try { await process.WaitForExitAsync(timeout.Token); }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw new TimeoutException("Automation script exceeded its 30-second deadline.");
        }
        return (process.ExitCode, await stdout + await stderr);
    }

    private static string FindRepoRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Sbroenne.ExcelMcp.sln"))) { return directory.FullName; }
        }
        throw new DirectoryNotFoundException("Repository root was not found.");
    }
}
