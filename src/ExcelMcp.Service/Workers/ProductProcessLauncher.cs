using System.Diagnostics;
using System.Reflection;

namespace Sbroenne.ExcelMcp.Service.Workers;

internal sealed record ProductProcessLauncher(string FileName, IReadOnlyList<string> PrefixArguments)
{
    internal static ProductProcessLauncher Current()
    {
        var executable = Environment.ProcessPath
            ?? throw new InvalidOperationException("Cannot locate the product executable for an Excel worker.");
        var entry = Assembly.GetEntryAssembly()?.GetName().Name
            ?? throw new InvalidOperationException("Cannot identify the product hosting the Excel controller.");
        if (entry is not ("excelcli" or "Sbroenne.ExcelMcp.McpServer"))
            throw new InvalidOperationException("Tests must supply a built product launcher for Excel workers.");
        return Path.GetFileNameWithoutExtension(executable).Equals("dotnet", StringComparison.OrdinalIgnoreCase)
            ? new(executable, [Path.Combine(AppContext.BaseDirectory, $"{entry}.dll")])
            : new(executable, []);
    }

    internal Process Start(IEnumerable<string> arguments)
    {
        var start = new ProcessStartInfo(FileName)
        {
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            WorkingDirectory = Path.GetDirectoryName(FileName) ?? Environment.CurrentDirectory
        };
        foreach (var argument in PrefixArguments.Concat(arguments))
            start.ArgumentList.Add(argument);
        return Process.Start(start)
            ?? throw new InvalidOperationException("The Excel worker process could not be started.");
    }
}
