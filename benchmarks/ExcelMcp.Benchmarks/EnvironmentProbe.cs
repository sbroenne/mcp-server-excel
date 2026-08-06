using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Benchmarks;

internal static class EnvironmentProbe
{
    public static BenchmarkEnvironment Capture(BenchmarkContext context)
    {
        var excelVersion = CaptureExcelVersion(context);
        var gitCommit = RunGit(context.Options.RepoRoot, "rev-parse", "HEAD") ?? "unknown";
        var gitBranch = RunGit(context.Options.RepoRoot, "branch", "--show-current") ?? "unknown";
        var gitDirty = !string.IsNullOrWhiteSpace(RunGit(context.Options.RepoRoot, "status", "--porcelain"));
        var machineHash = Convert.ToHexString(
            SHA256.HashData(Encoding.UTF8.GetBytes(Environment.MachineName)))[..12].ToLowerInvariant();

        return new BenchmarkEnvironment(
            machineHash,
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            excelVersion,
            RuntimeInformation.ProcessArchitecture.ToString(),
            Environment.ProcessorCount,
            GC.GetGCMemoryInfo().TotalAvailableMemoryBytes,
            gitCommit,
            gitBranch,
            gitDirty);
    }

    private static string CaptureExcelVersion(BenchmarkContext context)
    {
        var path = context.CreateWorkingPath("environment-probe");
        BenchmarkContext.CreateEmptyWorkbook(path);
        using var batch = ExcelSession.BeginBatch(path);
        return batch.Execute((excelContext, _) =>
        {
            var version = Convert.ToString(excelContext.App.Version, System.Globalization.CultureInfo.InvariantCulture) ?? "unknown";
            var build = Convert.ToString(excelContext.App.Build, System.Globalization.CultureInfo.InvariantCulture) ?? "unknown";
            return $"{version} build {build}";
        });
    }

    private static string? RunGit(string workingDirectory, params string[] arguments)
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = "git",
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (var argument in arguments)
        {
            process.StartInfo.ArgumentList.Add(argument);
        }

        try
        {
            if (!process.Start())
            {
                return null;
            }

            var output = process.StandardOutput.ReadToEnd();
            if (!process.WaitForExit(10_000) || process.ExitCode != 0)
            {
                return null;
            }

            return output.Trim();
        }
        catch (System.ComponentModel.Win32Exception)
        {
            return null;
        }
    }
}
