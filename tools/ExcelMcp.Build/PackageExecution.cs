using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Build;

public sealed class PackageOptions
{
    public string Operation { get; set; } = "Release";
    public string[] Components { get; set; } = ["Cli", "Mcp", "Extension", "Mcpb", "Skills", "Plugins"];
    public string? Version { get; set; }
    public string? OutputDirectory { get; set; }
    public string? SkillsDirectory { get; set; }
    public string? McpRuntimeExecutable { get; set; }
    public string? CliRuntimeExecutable { get; set; }
    public string Component { get; set; } = "McpServer";
    public string Architecture { get; set; } = "x64";
    public string? RuntimeExecutable { get; set; }
    public string? LauncherPackage { get; set; }
    public string? RuntimePackage { get; set; }
    public bool GenerateOnly { get; set; }
    public bool ArchiveOnly { get; set; }
    public bool SkipExtensionTests { get; set; }
    public string? BaseRef { get; set; }
    public string HeadRef { get; set; } = "HEAD";
    public string? Source { get; set; }
    public string? Destination { get; set; }
    public string[] Inputs { get; set; } = [];
    public double TimeoutMilliseconds { get; set; } = 120000;
    public double RetryMilliseconds { get; set; } = 500;
}

public sealed partial class PackageExecution(string root, PackageCommands? native = null)
{
    private readonly PackageCommands _commands = native ?? new();
    private static readonly HashSet<string> Components = new(StringComparer.Ordinal) { "Cli", "Mcp", "Extension", "Mcpb", "Skills", "Plugins" };

    public async Task ExecuteAsync(PackageOptions options)
    {
        var authored = new AuthoredPackages(root);
        switch (options.Operation)
        {
            case "Skills": Console.WriteLine(authored.Skills(options.Version, options.OutputDirectory, options.GenerateOnly, options.SkillsDirectory)); break;
            case "Plugins":
                Console.WriteLine(authored.Plugins(options.Version, options.OutputDirectory, options.SkillsDirectory));
                Console.WriteLine("[ok] excel-mcp - npx config and skill");
                Console.WriteLine("[ok] excel-cli - argument-safe npx wrapper and skill");
                break;
            case "Mcpb": Console.WriteLine(authored.Mcpb(options.Version, options.OutputDirectory)); break;
            case "Npm":
                var packages = await new RuntimePackages(root, _commands).BuildNpmAsync(options.Component, options.Architecture, Required(options.Version), Required(options.RuntimeExecutable), Required(options.OutputDirectory));
                Console.WriteLine(new JsonObject { ["LauncherPackage"] = packages.Launcher, ["RuntimePackage"] = packages.Runtime }.ToJsonString());
                break;
            case "VerifyNpm":
                await new RuntimePackages(root, _commands).VerifyNpmAsync(options.Component, options.Architecture, Required(options.LauncherPackage), Required(options.RuntimePackage), options.ArchiveOnly); break;
            case "AssertOutput": PackageFiles.AssertOutput(Required(options.OutputDirectory), root, options.Inputs); break;
            case "RuntimeArchitecture": PackageFiles.AssertArchitecture(Required(options.RuntimeExecutable), options.Architecture); break;
            case "InstallOutput": PackageFiles.Install(Required(options.Source), Required(options.Destination)); break;
            case "RemoveStaging": PackageFiles.RemoveStaging(Required(options.Source), TimeSpan.FromMilliseconds(options.TimeoutMilliseconds), TimeSpan.FromMilliseconds(options.RetryMilliseconds)); break;
            case "PublishRuntime": await PublishRuntimeAsync(options.Component, options.Architecture, AuthoredPackages.RequiredVersion(options.Version), Required(options.OutputDirectory)); break;
            case "Release": await ReleaseAsync(options); break;
            default: throw new ArgumentException($"Unknown package operation: {options.Operation}");
        }
    }

    public async Task ReleaseAsync(PackageOptions options)
    {
        if (options.Components.Any(component => !Components.Contains(component))) { throw new ArgumentException("Unknown release package component."); }
        var selected = options.Components.Distinct(StringComparer.Ordinal).ToArray();
        if (options.BaseRef is not null)
        {
            var result = await _commands.RunAsync(root, "git", ["-c", "core.quotepath=false", "diff", "--name-only", "--no-renames", $"{options.BaseRef}...{options.HeadRef}"]);
            var plan = new ValidationPolicy(root).Select(result.Output.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries));
            selected = selected.Where(component => component switch
            {
                "Cli" => plan.Cli,
                "Mcp" => plan.Mcp,
                "Extension" => plan.Extension,
                "Mcpb" => plan.Mcpb,
                "Skills" => plan.Skills,
                "Plugins" => plan.Plugins,
                _ => false
            }).ToArray();
            options.SkipExtensionTests |= !plan.ExtensionTests;
            foreach (var reason in plan.Reasons) { Console.WriteLine(reason); }
        }
        if (selected.Length == 0) { Console.WriteLine("No distributable package inputs changed."); return; }
        if (!OperatingSystem.IsWindows()) { throw new InvalidOperationException("Package installation checks require Windows (not Excel)."); }
        var authored = new AuthoredPackages(root);
        var version = AuthoredPackages.RequiredVersion(options.Version ?? authored.PackageVersion());
        var skills = options.SkillsDirectory is null ? null : Path.GetFullPath(options.SkillsDirectory, root);
        var needsSkills = selected.Any(component => component is "Skills" or "Plugins" or "Extension");
        if (skills is not null && needsSkills)
        {
            var requiredSkills = selected.Any(component => component is "Skills" or "Plugins")
                ? new[] { "excel-cli", "excel-cli-report-formatting", "excel-mcp-report-formatting" }
                : ["excel-mcp-report-formatting"];
            foreach (var name in requiredSkills)
            {
                var stamp = Path.Combine(skills, name, "VERSION");
                if (!File.Exists(stamp) || File.ReadAllText(stamp).Trim() != version) { throw new InvalidOperationException($"Prepared {name} skill must match package version {version}."); }
            }
        }
        var output = Path.GetFullPath(options.OutputDirectory ?? Path.Combine(root, "artifacts", "packages", Guid.NewGuid().ToString("N")), root);
        PackageFiles.AssertOutput(output, root, [skills, options.McpRuntimeExecutable, options.CliRuntimeExecutable]);
        if (Directory.Exists(output) || File.Exists(output)) { throw new ArgumentException($"Use a new package output directory: {output}"); }
        Directory.CreateDirectory(output);
        var prepared = new Dictionary<string, string>(StringComparer.Ordinal);
        var runtimes = new List<string>();
        if (selected.Contains("Cli")) { runtimes.Add("Cli"); }
        if (selected.Any(component => component is "Mcp" or "Extension")) { runtimes.Add("Mcp"); }
        foreach (var component in runtimes)
        {
            if (selected.Contains(component)) { await NuGetAsync(component, version, output); }
            var supplied = component == "Cli" ? options.CliRuntimeExecutable : options.McpRuntimeExecutable;
            var executable = supplied is null ? await PublishRuntimeAsync(component, "x64", version, Path.Combine(output, "runtimes", component)) : Path.GetFullPath(supplied, root);
            VerifyRuntime(executable, "x64", version);
            prepared[component] = executable;
            await _commands.RunAsync(root, executable, ["--version"]);
            if (!selected.Contains(component)) { continue; }
            var npm = new RuntimePackages(root, _commands);
            foreach (var architecture in new[] { "x64", "arm64" })
            {
                var runtime = executable;
                if (architecture == "arm64")
                {
                    runtime = await PublishRuntimeAsync(component, architecture, version, Path.Combine(output, "runtimes", $"{component}-arm64"));
                    VerifyRuntime(runtime, architecture, version);
                    prepared[$"{component}-arm64"] = runtime;
                }
                var archives = await npm.BuildNpmAsync(component, architecture, version, runtime, Path.Combine(output, "npm"));
                await npm.VerifyNpmAsync(component, architecture, archives.Launcher, archives.Runtime, archiveOnly: architecture == "arm64");
            }
            var stage = Path.Combine(output, $"zip-{component}");
            PackageFiles.Copy(executable, Path.Combine(stage, component == "Cli" ? "excelcli.exe" : "mcp-excel.exe"));
            foreach (var file in new[] { "README.md", "LICENSE", "CHANGELOG.md" }) { PackageFiles.Copy(Path.Combine(root, file), Path.Combine(stage, file)); }
            ZipFile.CreateFromDirectory(stage, Path.Combine(output, component == "Cli" ? $"ExcelMcp-CLI-{version}-windows.zip" : $"ExcelMcp-MCP-Server-{version}-windows.zip"));
        }
        if (selected.Contains("Mcpb")) { authored.Mcpb(version, Path.Combine(output, "mcpb")); }
        if (needsSkills && skills is null) { skills = authored.Skills(version, Path.Combine(output, "generated-skills"), generateOnly: true, prepared: null); }
        if (selected.Contains("Skills")) { authored.Skills(version, Path.Combine(output, "skills"), generateOnly: false, skills); }
        if (selected.Contains("Plugins"))
        {
            var plugins = authored.Plugins(version, Path.Combine(output, "plugins"), skills);
            ZipFile.CreateFromDirectory(plugins, Path.Combine(output, $"excel-plugins-v{version}.zip"));
        }
        if (selected.Contains("Extension")) { await ExtensionAsync(version, Required(skills), output, prepared, options.SkipExtensionTests); }
        Console.WriteLine($"Verified packages: {string.Join(", ", selected)}. Output: {output}");
    }

    public async Task ExtensionAsync(string version, string skills, string output, IDictionary<string, string> prepared, bool skipTests)
    {
        if (!prepared.TryGetValue("Mcp-arm64", out var arm))
        {
            arm = await PublishRuntimeAsync("Mcp", "arm64", version, Path.Combine(output, "runtimes", "Mcp-arm64"));
            VerifyRuntime(arm, "arm64", version);
            prepared["Mcp-arm64"] = arm;
        }
        var stage = AuthoredPackages.NewStage("extension");
        try
        {
            foreach (var entry in Directory.EnumerateFileSystemEntries(Path.Combine(root, "vscode-extension")).Where(entry =>
                Path.GetFileName(entry) is not ("node_modules" or "bin" or "out" or "skills") && Path.GetExtension(entry) != ".vsix"))
            {
                PackageFiles.Copy(entry, Path.Combine(stage, Path.GetFileName(entry)));
            }
            var bundled = Path.Combine(stage, "bin", "Sbroenne.ExcelMcp.McpServer.exe");
            PackageFiles.Copy(prepared["Mcp"], bundled);
            var skill = Path.Combine(stage, "skills", "excel-mcp-report-formatting");
            PackageFiles.Copy(Path.Combine(skills, "excel-mcp-report-formatting"), skill);
            File.WriteAllText(Path.Combine(skill, "VERSION"), version);
            PackageFiles.Copy(Path.Combine(root, "CHANGELOG.md"), Path.Combine(stage, "CHANGELOG.md"));
            var manifestPath = Path.Combine(stage, "package.json");
            var manifest = AuthoredPackages.ReadJson(manifestPath);
            manifest["version"] = version;
            AuthoredPackages.Map(manifest, "scripts")["vscode:prepublish"] = "npm run compile";
            AuthoredPackages.WriteJson(manifestPath, manifest);
            await _commands.RunAsync(stage, "npm", ["ci", "--ignore-scripts"]);
            await _commands.RunAsync(stage, "npm", ["run", "compile"]);
            await _commands.RunAsync(stage, "npm", ["run", "lint"]);
            if (!skipTests)
            {
                await _commands.RunAsync(stage, "npm", ["run", "typecheck:tests"]);
                await _commands.RunAsync(stage, "npm", ["test"]);
            }
            foreach (var (architecture, runtime, filename) in new[] {
                ("x64", prepared["Mcp"], $"excel-mcp-{version}.vsix"),
                ("arm64", arm, $"excel-mcp-{version}-win32-arm64.vsix")
            })
            {
                PackageFiles.Copy(runtime, bundled);
                var target = $"win32-{architecture}";
                var path = Path.Combine(output, filename);
                await _commands.RunAsync(stage, "npm", ["exec", "--", "vsce", "package", "--no-dependencies", "--target", target, "--out", path]);
                VerifyVsix(path, stage, skill, target, architecture, version);
            }
            PackageFiles.Copy(RuntimeInformation.OSArchitecture == Architecture.Arm64 ? arm : prepared["Mcp"], bundled);
            var debug = Path.Combine(output, "extension");
            Directory.CreateDirectory(debug);
            foreach (var entry in Directory.EnumerateFileSystemEntries(stage).Where(entry => Path.GetFileName(entry) != "node_modules"))
            {
                PackageFiles.Copy(entry, Path.Combine(debug, Path.GetFileName(entry)));
            }
        }
        finally { PackageFiles.RemoveStaging(stage); }
    }

    private async Task NuGetAsync(string component, string version, string output)
    {
        var projectName = component == "Cli" ? "CLI" : "McpServer";
        await _commands.RunAsync(root, "dotnet", ["pack", Path.Combine(root, "src", $"ExcelMcp.{projectName}", $"ExcelMcp.{projectName}.csproj"), "-c", "Release", $"-p:Version={version}", "-p:NuGetAudit=false", "-o", Path.Combine(output, "nuget")]);
        var config = Path.Combine(output, "nuget.config");
        new XDocument(new XElement("configuration", new XElement("packageSources", new XElement("clear"),
            new XElement("add", new XAttribute("key", "built"), new XAttribute("value", Path.Combine(output, "nuget")))))).Save(config);
        var installed = Path.Combine(output, $"installed-{component}");
        await _commands.RunAsync(root, "dotnet", ["tool", "install", $"Sbroenne.ExcelMcp.{projectName}", "--version", version, "--tool-path", installed, "--configfile", config, "--no-cache"]);
        await _commands.RunAsync(root, Path.Combine(installed, component == "Cli" ? "excelcli.exe" : "mcp-excel.exe"), ["--version"]);
    }

    private async Task<string> PublishRuntimeAsync(string component, string architecture, string version, string output)
    {
        if (component is not ("Cli" or "Mcp") || architecture is not ("x64" or "arm64")) { throw new ArgumentException("Unknown runtime component or architecture."); }
        var projectName = component == "Cli" ? "CLI" : "McpServer";
        await _commands.RunAsync(root, "dotnet", ["publish", Path.Combine(root, "src", $"ExcelMcp.{projectName}", $"ExcelMcp.{projectName}.csproj"),
            "-c", "Release", "-r", $"win-{architecture}", "--self-contained", "true", "-p:PublishSingleFile=true",
            "-p:IncludeNativeLibrariesForSelfExtract=true", "-p:PublishTrimmed=false", "-p:PublishReadyToRun=false",
            "-p:NuGetAudit=false", $"-p:Version={version}", "-o", output, "--verbosity", "minimal"]);
        return Path.Combine(output, component == "Cli" ? "excelcli.exe" : "Sbroenne.ExcelMcp.McpServer.exe");
    }

    private static void VerifyRuntime(string path, string architecture, string version)
    {
        PackageFiles.AssertArchitecture(path, architecture);
        var product = FileVersionInfo.GetVersionInfo(path).ProductVersion;
        if (product?.Split('+')[0] != version) { throw new InvalidOperationException($"Runtime version {product} does not match {version}."); }
    }

    private static void VerifyVsix(string path, string stage, string skill, string target, string architecture, string version)
    {
        using var zip = ZipFile.OpenRead(path);
        foreach (var entry in new[] { "extension/bin/Sbroenne.ExcelMcp.McpServer.exe", "extension/out/extension.js", "extension/out/prerequisites.js" })
        {
            if (zip.GetEntry(entry) is null) { throw new InvalidOperationException($"VSIX is missing {entry}."); }
        }
        var inspection = Path.Combine(Path.GetDirectoryName(path)!, $"{target}-server-inspection.exe");
        try
        {
            zip.GetEntry("extension/bin/Sbroenne.ExcelMcp.McpServer.exe")!.ExtractToFile(inspection, overwrite: false);
            PackageFiles.AssertArchitecture(inspection, architecture);
        }
        finally { if (File.Exists(inspection)) { PackageFiles.Delete(inspection); } }
        foreach (var file in Directory.EnumerateFiles(skill, "*", SearchOption.AllDirectories))
        {
            var relative = Path.GetRelativePath(stage, file).Replace('\\', '/');
            if (zip.GetEntry($"extension/{relative}") is null) { throw new InvalidOperationException($"VSIX is missing {relative}."); }
        }
        if (ReadEntry(zip, "extension/skills/excel-mcp-report-formatting/VERSION").Trim() != version) { throw new InvalidOperationException("VSIX skill version does not match the package."); }
        var manifest = JsonNode.Parse(ReadEntry(zip, "extension/package.json")) as JsonObject ?? throw new InvalidOperationException("Invalid VSIX manifest.");
        if (AuthoredPackages.Text(manifest, "version") != version ||
            string.Join(',', AuthoredPackages.Array(manifest, "extensionKind").Select(value => value!.GetValue<string>())) != "ui" ||
            string.Join(',', AuthoredPackages.Array(manifest, "os").Select(value => value!.GetValue<string>())) != "win32")
        {
            throw new InvalidOperationException("VSIX version or local Windows host metadata is incorrect.");
        }
        var metadata = XDocument.Parse(ReadEntry(zip, "extension.vsixmanifest"));
        if (metadata.Descendants().FirstOrDefault(element => element.Name.LocalName == "Identity")?.Attribute("TargetPlatform")?.Value != target) { throw new InvalidOperationException($"VSIX target does not match {target}."); }
        foreach (var entry in zip.Entries)
        {
            if (DevelopmentEntry().IsMatch(entry.FullName)) { throw new InvalidOperationException($"VSIX contains development files or the CLI: {entry.FullName}"); }
        }
        Console.WriteLine($"Verified VSIX target: {target}");
    }

    private static string ReadEntry(ZipArchive zip, string name)
    {
        using var reader = new StreamReader((zip.GetEntry(name) ?? throw new InvalidOperationException($"VSIX is missing {name}.")).Open());
        return reader.ReadToEnd();
    }
    private static string Required(string? value) => !string.IsNullOrWhiteSpace(value) ? value : throw new ArgumentException("A required package argument is missing.");
    [GeneratedRegex(@"^extension/(?:node_modules|tests|scripts|\.vitest|coverage|out/tests)/|^extension/(?:vitest\.config\.|tsconfig(?:\.test)?\.json$|bin/excelcli)|^extension/(?:.*/)?(?:AGENTS|CLAUDE)\.md$", RegexOptions.CultureInvariant)]
    private static partial Regex DevelopmentEntry();
}
