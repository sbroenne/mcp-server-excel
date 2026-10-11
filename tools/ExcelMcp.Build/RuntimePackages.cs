using System.Formats.Tar;
using System.IO.Compression;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Build;

public sealed class RuntimePackages(string root, PackageCommands commands)
{
    public async Task<(string Launcher, string Runtime)> BuildNpmAsync(string component, string architecture, string version, string executable, string output)
    {
        var (name, command) = Names(component);
        version = AuthoredPackages.RequiredVersion(version);
        executable = Path.GetFullPath(executable, root);
        output = Path.GetFullPath(output, root);
        PackageFiles.AssertOutput(output, root, [executable]);
        if (!Path.GetExtension(executable).Equals(".exe", StringComparison.OrdinalIgnoreCase)) { throw new ArgumentException($"Runtime executable must be an .exe file: {executable}"); }
        PackageFiles.AssertArchitecture(executable, architecture);
        var stage = AuthoredPackages.NewStage("npm");
        try
        {
            var launcher = Path.Combine(stage, name);
            var runtime = Path.Combine(stage, $"{name}-win32-{architecture}");
            foreach (var (source, destination, entries) in new[]
            {
                (name, launcher, new[] { "package.json", "README.md", "bin" }),
                ($"{name}-win32-{architecture}", runtime, new[] { "package.json", "README.md" })
            })
            {
                Directory.CreateDirectory(destination);
                foreach (var entry in entries) { PackageFiles.Copy(Path.Combine(root, "npm-packages", source, entry), Path.Combine(destination, entry)); }
                PackageFiles.Copy(Path.Combine(root, "LICENSE"), Path.Combine(destination, "LICENSE"));
            }
            PackageFiles.Copy(Path.Combine(root, "npm-packages", "shared", "launcher.js"), Path.Combine(launcher, "lib", "launcher.js"));
            PackageFiles.Copy(executable, Path.Combine(runtime, $"{command}.exe"));
            foreach (var directory in new[] { runtime, launcher })
            {
                var path = Path.Combine(directory, "package.json");
                var manifest = AuthoredPackages.ReadJson(path);
                manifest["version"] = version;
                if (directory == launcher)
                {
                    var dependencies = AuthoredPackages.Map(manifest, "optionalDependencies");
                    foreach (var arch in new[] { "x64", "arm64" }) { dependencies[$"@sbroenne/{name}-win32-{arch}"] = version; }
                }
                AuthoredPackages.WriteJson(path, manifest);
            }
            Directory.CreateDirectory(output);
            var runtimeArchive = await PackAsync(runtime, output, stage, [$"{command}.exe", "package.json", "LICENSE"]);
            var launcherArchive = await PackAsync(launcher, output, stage, [$"bin/{command}.js", "lib/launcher.js", "package.json", "LICENSE"]);
            return (launcherArchive, runtimeArchive);
        }
        finally { PackageFiles.RemoveStaging(stage); }
    }

    public async Task VerifyNpmAsync(string component, string architecture, string launcher, string runtime, bool archiveOnly)
    {
        var (name, command) = Names(component);
        var stage = AuthoredPackages.NewStage("npmtest");
        try
        {
            string? version = null;
            foreach (var (archive, kind, packageName) in new[] { (runtime, "runtime", $"{name}-win32-{architecture}"), (launcher, "launcher", name) })
            {
                var directory = Path.Combine(stage, kind);
                ExtractTar(Path.GetFullPath(archive, root), directory);
                var package = Path.Combine(directory, "package");
                var manifest = AuthoredPackages.ReadJson(Path.Combine(package, "package.json"));
                if (AuthoredPackages.Text(manifest, "name") != $"@sbroenne/{packageName}") { throw new InvalidOperationException($"Unexpected npm package name: {manifest["name"]}"); }
                if (!File.Exists(Path.Combine(package, "LICENSE"))) { throw new InvalidOperationException($"Missing license in {kind} npm archive."); }
                if (kind == "runtime")
                {
                    version = AuthoredPackages.Text(manifest, "version");
                    if (AuthoredPackages.Text(manifest, "main") != $"{command}.exe" || !SingleValue(manifest, "os", "win32") || !SingleValue(manifest, "cpu", architecture))
                    {
                        throw new InvalidOperationException("npm runtime metadata does not match the requested architecture.");
                    }
                    PackageFiles.AssertArchitecture(Path.Combine(package, $"{command}.exe"), architecture);
                }
                else
                {
                    if (AuthoredPackages.Text(manifest, "version") != version || AuthoredPackages.Text(AuthoredPackages.Map(manifest, "bin"), command) != $"bin/{command}.js")
                    {
                        throw new InvalidOperationException("npm launcher metadata does not match the runtime.");
                    }
                    foreach (var arch in new[] { "x64", "arm64" })
                    {
                        if (AuthoredPackages.Text(AuthoredPackages.Map(manifest, "optionalDependencies"), $"@sbroenne/{name}-win32-{arch}") != version)
                        {
                            throw new InvalidOperationException($"npm launcher must depend on the matching {arch} release version.");
                        }
                    }
                    foreach (var file in new[] { $"bin/{command}.js", "lib/launcher.js" })
                    {
                        if (!File.Exists(Path.Combine(package, file))) { throw new InvalidOperationException($"Missing launcher file: {file}"); }
                    }
                }
            }
            Console.WriteLine($"{component} {architecture} npm archives validated.");
            if (archiveOnly) { Console.WriteLine("Archive-only validation requested; native execution is a separate check."); return; }
            if (!OperatingSystem.IsWindows()) { throw new InvalidOperationException("npm runtime smoke tests require Windows."); }
            var node = await commands.RunAsync(stage, "node", ["-p", "process.arch"]);
            if (node.Output.Trim() != architecture)
            {
                Console.Error.WriteLine($"{component} {architecture} execution NOT RUN: Node.js is {node.Output.Trim()}. Archive validation passed.");
                return;
            }
            await commands.RunAsync(stage, "npm", ["install", "--prefix", stage, "--ignore-scripts", "--no-audit", "--no-fund", Path.GetFullPath(runtime, root), Path.GetFullPath(launcher, root)]);
            var installed = Path.Combine(stage, "node_modules", "@sbroenne", name, "bin", $"{command}.js");
            Console.WriteLine((await commands.RunAsync(stage, "node", [installed, "--version"])).Output.Trim());
            Console.WriteLine((await commands.RunAsync(stage, "node", [Path.Combine(root, "npm-packages", name, "scripts", "verify-runtime.mjs"), installed])).Output.Trim());
        }
        finally { PackageFiles.RemoveStaging(stage, TimeSpan.FromSeconds(5), TimeSpan.FromMilliseconds(250)); }
    }

    private async Task<string> PackAsync(string directory, string output, string stage, string[] required)
    {
        var manifest = AuthoredPackages.ReadJson(Path.Combine(directory, "package.json"));
        var name = AuthoredPackages.Text(manifest, "name").TrimStart('@').Replace('/', '-');
        var archive = Path.Combine(output, $"{name}-{AuthoredPackages.Text(manifest, "version")}.tgz");
        if (File.Exists(archive)) { PackageFiles.Delete(archive); }
        await commands.RunAsync(root, "npm", ["pack", directory, "--pack-destination", output, "--silent"]);
        if (!File.Exists(archive)) { throw new InvalidOperationException($"npm pack did not create the expected archive '{archive}'."); }
        var inspection = Path.Combine(stage, $"inspect-{Guid.NewGuid():N}");
        ExtractTar(archive, inspection);
        foreach (var file in required)
        {
            if (!File.Exists(Path.Combine(inspection, "package", file))) { throw new InvalidOperationException($"Packed npm package '{Path.GetFileName(archive)}' is missing required file '{file}'."); }
        }
        return archive;
    }

    public static void ExtractTar(string archive, string destination)
    {
        Directory.CreateDirectory(destination);
        using var file = File.OpenRead(archive);
        using var compressed = new GZipStream(file, CompressionMode.Decompress);
        using var reader = new TarReader(compressed);
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (reader.GetNextEntry() is { } entry)
        {
            var path = Path.GetFullPath(entry.Name.Replace('/', Path.DirectorySeparatorChar), destination);
            if (!PackageFiles.Contains(destination, path) || !paths.Add(path)) { throw new InvalidOperationException($"Unsafe or duplicate npm archive entry: {entry.Name}"); }
            switch (entry.EntryType)
            {
                case TarEntryType.Directory: Directory.CreateDirectory(path); break;
                case TarEntryType.RegularFile:
                case TarEntryType.V7RegularFile:
                    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                    entry.ExtractToFile(path, overwrite: false);
                    break;
                default: throw new InvalidOperationException($"Unsupported npm archive entry: {entry.Name} ({entry.EntryType}).");
            }
        }
    }

    public static (string Package, string Command) Names(string component) => component switch
    {
        "Cli" => ("excelcli", "excelcli"),
        "Mcp" or "McpServer" => ("mcp-server-excel", "mcp-excel"),
        _ => throw new ArgumentException($"Unknown package component: {component}")
    };
    private static bool SingleValue(JsonObject manifest, string key, string expected) =>
        manifest[key] is JsonArray { Count: 1 } value && value[0]?.GetValue<string>() == expected;
}

public sealed record PackageCommand(string WorkingDirectory, string Executable, string[] Arguments, TimeSpan Deadline);

public sealed class PackageCommands(Func<PackageCommand, Task<ProcessResult>>? run = null)
{
    public async Task<ProcessResult> RunAsync(string directory, string executable, string[] arguments)
    {
        var command = new PackageCommand(directory, executable, arguments, TimeSpan.FromMinutes(20));
        var result = await (run ?? RunNativeAsync)(command);
        if (result.ExitCode != 0) { throw new InvalidOperationException($"{executable} failed with exit code {result.ExitCode}.\n{result.Output}\n{result.Error}"); }
        return result;
    }

    private static Task<ProcessResult> RunNativeAsync(PackageCommand command)
    {
        var executable = command.Executable;
        var arguments = command.Arguments;
        if (executable == "npm" && OperatingSystem.IsWindows())
        {
            var node = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator)
                .Select(directory => Path.Combine(directory.Trim('"'), "node.exe")).FirstOrDefault(File.Exists)
                ?? throw new FileNotFoundException("Node.js is required for npm package operations.");
            var script = Path.Combine(Path.GetDirectoryName(node)!, "node_modules", "npm", "bin", "npm-cli.js");
            if (!File.Exists(script)) { throw new FileNotFoundException("The Node.js npm CLI is missing.", script); }
            executable = node;
            arguments = [script, .. arguments];
        }
        return new ProcessRunner(command.WorkingDirectory).RunAsync(executable, arguments, command.Deadline);
    }
}
