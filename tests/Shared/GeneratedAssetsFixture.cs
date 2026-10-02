using System.Diagnostics;
using Xunit;

namespace Sbroenne.ExcelMcp.Tests.Infrastructure;

[CollectionDefinition("GeneratedAssets", DisableParallelization = true)]
public sealed class GeneratedAssetsCollectionDefinition : ICollectionFixture<GeneratedAssetsFixture>;

public sealed class GeneratedAssetsFixture : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"ExcelMcpAssets-{Guid.NewGuid():N}");
    public static string SkillsDirectory { get; private set; } = "";

    public GeneratedAssetsFixture()
    {
        SkillsDirectory = Path.Combine(_root, "skills");
        var repo = new DirectoryInfo(AppContext.BaseDirectory);
        while (repo != null && !File.Exists(Path.Combine(repo.FullName, "Sbroenne.ExcelMcp.sln"))) { repo = repo.Parent; }
        if (repo == null) { throw new DirectoryNotFoundException("Repository root not found."); }
        try
        {
            Run(repo.FullName, "Build-AgentSkills", "-GenerateOnly", "-OutputDir", SkillsDirectory);
        }
        catch
        {
            Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        if (Directory.Exists(_root)) { Directory.Delete(_root, recursive: true); }
    }

    private static void Run(string root, string script, params string[] arguments)
    {
        var info = new ProcessStartInfo("pwsh")
        {
            WorkingDirectory = root,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var value in new[] { "-NoProfile", "-File", Path.Combine(root, "scripts", $"{script}.ps1") }.Concat(arguments))
        {
            info.ArgumentList.Add(value);
        }
        using var process = Process.Start(info)!;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(120000))
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit();
            throw new TimeoutException($"{script} exceeded two minutes.");
        }
        if (process.ExitCode != 0)
        {
            throw new InvalidOperationException($"{script} failed. Build the Release solution first.\n{stdout.GetAwaiter().GetResult()}\n{stderr.GetAwaiter().GetResult()}");
        }
    }
}
