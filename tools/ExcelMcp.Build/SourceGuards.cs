namespace Sbroenne.ExcelMcp.Build;

/// <summary>Compatibility entry point; the compiler owns source policy.</summary>
public sealed class SourceGuards(string root)
{
    private static readonly string[] Rules = ["com-leaks", "success-flag", "dynamic-casts", "workbook-package-access"];

    public async Task CheckAsync(string rule, IReadOnlyList<string>? inputs, IProcessRunner runner)
    {
        if (!Rules.Contains(rule, StringComparer.Ordinal))
        {
            throw new ArgumentException($"Unknown source guard: {rule}.");
        }
        if (inputs is { Count: 0 }) { throw new ArgumentException("Explicit source inputs cannot be empty."); }
        if (inputs is { Count: > 0 } && rule != "com-leaks")
        {
            throw new ArgumentException("Explicit input paths are supported only by the com-leaks compatibility command.");
        }
        var projects = Projects(inputs);
        Console.Error.WriteLine("Source safety is enforced by compilation (all four policies), not an independent scanner.");
        foreach (var project in projects)
        {
            var policy = await runner.CheckedAsync("dotnet",
                ["msbuild", project, "-nologo", "-getProperty:ExcelMcpProductionSafety"], TimeSpan.FromMinutes(2));
            if (!policy.Output.Trim().Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Production compiler source safety is not active for {project}. No source check was performed.");
            }
            var result = await runner.CheckedAsync("dotnet",
                ["build", project, "-c", "Release", "--no-incremental", "--disable-build-servers", "--verbosity", "minimal"],
                TimeSpan.FromMinutes(20));
            Console.Error.WriteLine(result.Output);
            if (!string.IsNullOrWhiteSpace(result.Error)) { Console.Error.WriteLine(result.Error); }
        }
    }

    internal string[] Projects(IReadOnlyList<string>? inputs)
    {
        var sourceRoot = Path.Combine(root, "src");
        if (!Directory.Exists(sourceRoot)) { throw new DirectoryNotFoundException($"Production source root not found: {sourceRoot}."); }
        var projects = Directory.EnumerateFiles(sourceRoot, "*.csproj", SearchOption.AllDirectories)
            .Where(path => !Path.GetFileNameWithoutExtension(path).StartsWith("ExcelMcp.Generators", StringComparison.Ordinal) &&
                Path.GetFileNameWithoutExtension(path) != "ExcelMcp.Analyzers" &&
                !path.Split(Path.DirectorySeparatorChar).Any(segment => segment is "bin" or "obj"))
            .Order(StringComparer.Ordinal).ToArray();
        if (projects.Length == 0) { throw new InvalidOperationException("No production projects found for compiler source safety validation."); }
        if (inputs is not { Count: > 0 }) { return projects; }
        var selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var input in inputs)
        {
            var path = Path.GetFullPath(input, root);
            if (!path.Equals(sourceRoot, StringComparison.OrdinalIgnoreCase) &&
                !path.StartsWith(Path.TrimEndingDirectorySeparator(sourceRoot) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Source input must belong to production src: {input}.");
            }
            if (!File.Exists(path) && !Directory.Exists(path)) { throw new FileNotFoundException($"Source input not found: {path}."); }
            var matches = projects.Where(project =>
                path.StartsWith(Path.GetDirectoryName(project)! + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                project.StartsWith(Path.TrimEndingDirectorySeparator(path) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                path.Equals(project, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length == 0) { throw new ArgumentException($"Source input has no production project: {input}."); }
            selected.UnionWith(matches);
        }
        return selected.Order(StringComparer.Ordinal).ToArray();
    }
}
