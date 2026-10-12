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
            var script = $"$p=Get-Process -Id $PID; $child=Start-Process pwsh -ArgumentList '-NoProfile','-Command','Start-Sleep -Seconds 60' -PassThru; @($p,$child) | ForEach-Object {{ \"$($_.Id),$($_.StartTime.ToUniversalTime().Ticks)\" }} | Set-Content -LiteralPath '{file.Replace("'", "''", StringComparison.Ordinal)}'; Start-Sleep -Seconds 60";
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
        finally { File.Delete(file); }
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
                    await retained.WaitForExitAsync();
                }
            }
        }
    }
}
