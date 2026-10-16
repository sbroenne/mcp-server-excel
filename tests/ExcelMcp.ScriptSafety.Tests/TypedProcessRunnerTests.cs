using System.Diagnostics;
using System.Globalization;
using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class TypedProcessRunnerTests
{
    [Fact]
    public async Task NativeFailure_PreservesExitCodeAndOriginalDiagnostic()
    {
        var runner = new ProcessRunner(TypedValidationPolicyTests.Root);
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            runner.CheckedAsync("pwsh", ["-NoProfile", "-Command", "[Console]::Error.WriteLine('native-root-cause'); exit 23"],
                TimeSpan.FromSeconds(30)));
        Assert.Contains("exit code 23", error.Message, StringComparison.Ordinal);
        Assert.Contains("native-root-cause", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task NativeArguments_AreNotCombinedIntoShellCode()
    {
        var runner = new ProcessRunner(TypedValidationPolicyTests.Root);
        var value = "spaces; 'quotes' & $variables";
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.Arguments.{Guid.NewGuid():N}.ps1");
        try
        {
            File.WriteAllText(file, "param([string]$Value)\n[Console]::Write($Value)");
            var result = await runner.CheckedAsync("pwsh", ["-NoProfile", "-File", file, value], TimeSpan.FromSeconds(30));
            Assert.Equal(value, result.Output);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public async Task HardDeadline_TerminatesTheExactStartedProcess()
    {
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.TypedDeadline.{Guid.NewGuid():N}.txt");
        try
        {
            var runner = new ProcessRunner(TypedValidationPolicyTests.Root);
            var script = $"$p=Get-Process -Id $PID; $child=Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command','Start-Sleep -Seconds 60' -PassThru; @($p,$child) | ForEach-Object {{ \"$($_.Id),$($_.StartTime.ToUniversalTime().Ticks)\" }} | Set-Content -LiteralPath '{file.Replace("'", "''", StringComparison.Ordinal)}'; Start-Sleep -Seconds 60";
            await Assert.ThrowsAsync<TimeoutException>(() =>
                runner.RunAsync("pwsh", ["-NoProfile", "-Command", script], TimeSpan.FromSeconds(10)));
            Assert.True(File.Exists(file), "The native child did not reach its startup marker.");
            var identities = File.ReadAllLines(file);
            Assert.Equal(2, identities.Length);
            foreach (var row in identities)
            {
                var identity = row.Split(',');
                var pid = int.Parse(identity[0], CultureInfo.InvariantCulture);
                var started = long.Parse(identity[1], CultureInfo.InvariantCulture);
                try
                {
                    using var process = Process.GetProcessById(pid);
                    Assert.True(process.HasExited || process.StartTime.ToUniversalTime().Ticks != started,
                        "An exact started process survived its deadline.");
                }
                catch (ArgumentException)
                {
                    // An exited process has no entry in the current process table.
                }
            }
        }
        finally
        {
            try { StopOwnedProcesses(file); }
            finally { File.Delete(file); }
        }
    }

    [Fact]
    public async Task HardDeadline_AlsoBoundsOutputHeldOpenAfterTheParentExits()
    {
        var file = Path.Combine(Path.GetTempPath(), $"ExcelMcp.OutputDeadline.{Guid.NewGuid():N}.txt");
        Task<ProcessResult>? operation = null;
        try
        {
            var script = $"$p=Get-Process -Id $PID; $child=Start-Process pwsh -NoNewWindow -ArgumentList '-NoProfile','-Command','Start-Sleep -Seconds 60' -PassThru; @($p,$child) | ForEach-Object {{ \"$($_.Id),$($_.StartTime.ToUniversalTime().Ticks)\" }} | Set-Content -LiteralPath '{file.Replace("'", "''", StringComparison.Ordinal)}'; [Console]::Write('parent-stdout'); [Console]::Error.Write('parent-stderr'); exit 0";
            operation = new ProcessRunner(TypedValidationPolicyTests.Root).RunAsync(
                "pwsh", ["-NoProfile", "-Command", script], TimeSpan.FromSeconds(3));

            var error = await Assert.ThrowsAsync<TimeoutException>(() => operation.WaitAsync(TimeSpan.FromSeconds(8)));

            Assert.Contains("hard deadline", error.Message, StringComparison.Ordinal);
            Assert.Contains("parent-stdout", error.Message, StringComparison.Ordinal);
            Assert.Contains("parent-stderr", error.Message, StringComparison.Ordinal);
            Assert.Equal(2, File.ReadAllLines(file).Length);
            var parent = File.ReadAllLines(file)[0].Split(',');
            try
            {
                using var process = Process.GetProcessById(int.Parse(parent[0], CultureInfo.InvariantCulture));
                Assert.True(process.HasExited || process.StartTime.ToUniversalTime().Ticks != long.Parse(parent[1], CultureInfo.InvariantCulture));
            }
            catch (ArgumentException) { }
        }
        finally
        {
            try
            {
                StopOwnedProcesses(file);
                if (operation is { IsCompleted: false }) { await operation.WaitAsync(TimeSpan.FromSeconds(10)); }
            }
            finally { File.Delete(file); }
        }
    }

    private static void StopOwnedProcesses(string file)
    {
        if (!File.Exists(file)) { return; }
        foreach (var row in File.ReadAllLines(file))
        {
            var identity = row.Split(',');
            Process process;
            try { process = Process.GetProcessById(int.Parse(identity[0], CultureInfo.InvariantCulture)); }
            catch (ArgumentException) { continue; }
            using (process)
            {
                _ = process.Handle;
                if (process.HasExited || process.StartTime.ToUniversalTime().Ticks != long.Parse(identity[1], CultureInfo.InvariantCulture)) { continue; }
                process.Kill(entireProcessTree: true);
                Assert.True(process.WaitForExit(10_000), "An exact owned process survived test cleanup.");
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HardDeadline_WhenTerminationThrows_OnlyAcceptsAConfirmedExit(bool exits)
    {
        Process? retained = null;
        var expected = new InvalidOperationException("Synthetic termination failure.");
        var runner = new ProcessRunner(TypedValidationPolicyTests.Root, process =>
        {
            retained = Process.GetProcessById(process.Id);
            if (exits)
            {
                process.Kill(entireProcessTree: true);
                Assert.True(process.WaitForExit(10_000), "The exact started process did not exit.");
            }
            throw expected;
        });
        try
        {
            var error = await Record.ExceptionAsync(() => runner.RunAsync("pwsh",
                ["-NoProfile", "-Command", "[Console]::WriteLine('deadline-stdout'); [Console]::Error.WriteLine('deadline-stderr'); Start-Sleep -Seconds 60"],
                TimeSpan.FromSeconds(5)));

            Assert.NotNull(retained);
            if (exits)
            {
                var timeout = Assert.IsType<TimeoutException>(error);
                Assert.IsAssignableFrom<OperationCanceledException>(timeout.InnerException);
                Assert.Contains("deadline-stdout", timeout.Message, StringComparison.Ordinal);
                Assert.Contains("deadline-stderr", timeout.Message, StringComparison.Ordinal);
                Assert.True(retained.HasExited);
            }
            else
            {
                Assert.Same(expected, error);
                Assert.False(retained.HasExited);
            }
        }
        finally
        {
            if (retained is not null)
            {
                using (retained)
                {
                    if (!retained.HasExited) { retained.Kill(entireProcessTree: true); }
                    Assert.True(retained.WaitForExit(10_000), "The exact retained process survived test cleanup.");
                }
            }
        }
    }
}
