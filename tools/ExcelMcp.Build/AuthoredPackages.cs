using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Build;

public sealed partial class AuthoredPackages(string root, Action<string, string>? createArchive = null)
{
    private readonly Action<string, string> _archive = createArchive ?? ((source, destination) => ZipFile.CreateFromDirectory(source, destination));
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private static readonly string[] SkillNames = ["excel-cli", "excel-cli-report-formatting", "excel-mcp-report-formatting"];
    private static readonly string[] McpbEntries = ["manifest.json", "icon-512.png", "README.md", "LICENSE", "CHANGELOG.md"];
    private static readonly HashSet<string> SkillFields = new(StringComparer.Ordinal) {
        "name", "description", "license", "compatibility", "metadata", "allowed-tools"
    };
    private static readonly HashSet<string> PluginFields = new(StringComparer.Ordinal) {
        "$schema", "name", "version", "description", "author", "homepage", "repository", "license", "keywords", "extensions"
    };
    private static readonly HashSet<string> ServerFields = new(StringComparer.Ordinal) { "type", "command", "args", "env", "cwd" };
    private static readonly HashSet<string> RuntimeExtensions = new(StringComparer.OrdinalIgnoreCase) { ".exe", ".dll", ".pdb" };
    private static readonly string[] PluginNames = ["excel-cli", "excel-mcp"];
    private static readonly string[] RetiredSkills = ["excel-mcp"];

    public string Skills(string? version, string? output, bool generateOnly, string? prepared)
    {
        version = RequiredVersion(version ?? (generateOnly ? PackageVersion() : null));
        output = Path.GetFullPath(output ?? (generateOnly ? @"artifacts\generated-skills" : @"artifacts\skills"), root);
        prepared = Path.GetFullPath(prepared ?? @"artifacts\generated-skills", root);
        PackageFiles.AssertOutput(output, root, generateOnly ? [] : [prepared]);
        var stage = NewStage("skills");
        try
        {
            var stagedSkills = Path.Combine(stage, "skills");
            foreach (var name in SkillNames)
            {
                var destination = Path.Combine(stagedSkills, name);
                PackageFiles.Copy(Path.Combine(generateOnly ? Path.Combine(root, "skills") : prepared, name), destination);
                var formatting = name.EndsWith("-report-formatting", StringComparison.Ordinal);
                if (generateOnly && formatting)
                {
                    var surface = name.StartsWith("excel-cli-", StringComparison.Ordinal) ? "cli" : "mcp";
                    var content = File.ReadAllText(Path.Combine(root, "docs", "reference", "report-formatting.md")).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
                    content = SurfaceFences().Replace(content, match => match.Groups["surface"].Value == surface
                        ? $"```{(surface == "cli" ? "powershell" : "text")}\n{match.Groups["body"].Value}```\n" : "");
                    content = ReferenceLinks().Replace(content, match =>
                        $"[{match.Groups[1].Value}](https://excelmcpserver.dev/reference/{match.Groups[2].Value}/{match.Groups[3].Value})");
                    Directory.CreateDirectory(Path.Combine(destination, "references"));
                    File.WriteAllText(Path.Combine(destination, "references", "report-formatting.md"), content);
                }
                foreach (var file in formatting ? new[] { "SKILL.md", @"references\report-formatting.md" } : ["SKILL.md"])
                {
                    if (!File.Exists(Path.Combine(destination, file))) { throw new InvalidOperationException($"{name} is missing {file}."); }
                }
                File.WriteAllText(Path.Combine(destination, "VERSION"), version);
            }
            Directory.CreateDirectory(output);
            if (generateOnly)
            {
                foreach (var name in SkillNames) { PackageFiles.Install(Path.Combine(stagedSkills, name), Path.Combine(output, name)); }
                foreach (var retired in RetiredSkills)
                {
                    var legacy = Path.Combine(output, retired);
                    if (!Directory.Exists(legacy)) { continue; }
                    PackageFiles.AssertOutput(legacy, root);
                    if (!File.Exists(Path.Combine(legacy, "SKILL.md")) || !File.Exists(Path.Combine(legacy, "VERSION")))
                    {
                        throw new InvalidOperationException($"Refusing to remove unrecognized retired skill output: {legacy}");
                    }
                    PackageFiles.Delete(legacy);
                }
            }
            else
            {
                PackageFiles.Copy(Path.Combine(root, "docs", "AGENT-SKILLS.md"), Path.Combine(stage, "README.md"));
                var archive = Path.Combine(Path.GetDirectoryName(stage)!, $"excel-skills-v{version}-{Guid.NewGuid():N}.zip");
                try
                {
                    _archive(stage, archive);
                    PackageFiles.Install(archive, Path.Combine(output, $"excel-skills-v{version}.zip"));
                }
                finally { PackageFiles.Delete(archive); }
            }
            return output;
        }
        finally { PackageFiles.RemoveStaging(stage); }
    }

    public string Plugins(string? version, string? output, string? prepared)
    {
        version = RequiredVersion(version);
        output = Path.GetFullPath(output ?? "plugins", root);
        prepared = Path.GetFullPath(prepared ?? @"artifacts\generated-skills", root);
        PackageFiles.AssertOutput(output, root, [prepared]);
        if (PackageFiles.Contains(root, output) && output != Path.Combine(root, "plugins") && !PackageFiles.Contains(Path.Combine(root, "artifacts"), output))
        {
            throw new ArgumentException($"Unsafe plugin output directory: {output}");
        }
        var stage = NewStage("plugins");
        try
        {
            foreach (var component in PluginNames)
            {
                var destination = Path.Combine(stage, component);
                PackageFiles.Copy(Path.Combine(root, ".github", "plugins", component), destination);
                var bin = Path.Combine(destination, "bin");
                if (Directory.Exists(bin))
                {
                    foreach (var file in Directory.EnumerateFiles(bin, "*", SearchOption.AllDirectories).Where(file =>
                        RuntimeExtensions.Contains(Path.GetExtension(file)) || file.EndsWith(".deps.json", StringComparison.OrdinalIgnoreCase) ||
                        file.EndsWith(".runtimeconfig.json", StringComparison.OrdinalIgnoreCase)))
                    {
                        PackageFiles.Delete(file);
                    }
                }
                var manifest = ReadJson(Path.Combine(destination, "plugin.json"));
                manifest["version"] = version;
                WriteJson(Path.Combine(destination, "plugin.json"), manifest);
                File.WriteAllText(Path.Combine(destination, "version.txt"), version);
                foreach (var name in component == "excel-cli" ? new[] { "excel-cli", "excel-cli-report-formatting" } : ["excel-mcp-report-formatting"])
                {
                    var skill = Path.Combine(destination, "skills", name);
                    if (Directory.Exists(skill)) { PackageFiles.Delete(skill); }
                    PackageFiles.Copy(Path.Combine(prepared, name), skill);
                    File.WriteAllText(Path.Combine(skill, "VERSION"), version);
                }
                ValidatePlugin(component, destination, version);
            }
            Directory.CreateDirectory(output);
            foreach (var component in PluginNames) { PackageFiles.AssertNoLinks(Path.Combine(output, component)); }
            foreach (var component in PluginNames) { PackageFiles.Install(Path.Combine(stage, component), Path.Combine(output, component)); }
            return output;
        }
        finally { PackageFiles.RemoveStaging(stage); }
    }

    public string Mcpb(string? version, string? output)
    {
        version ??= XDocument.Load(Path.Combine(root, "Directory.Build.props")).Descendants("Version").First().Value;
        version = RequiredVersion(version);
        output = Path.GetFullPath(output ?? Path.Combine(root, "mcpb", "artifacts"), root);
        PackageFiles.AssertOutput(output, root);
        if (output == Path.Combine(root, "mcpb")) { throw new ArgumentException($"Unsafe package output directory: {output}"); }
        var stage = NewStage("mcpb");
        try
        {
            var manifest = ReadJson(Path.Combine(root, "mcpb", "manifest.json"));
            manifest["version"] = version;
            WriteJson(Path.Combine(stage, "manifest.json"), manifest);
            foreach (var name in McpbEntries.Skip(1))
            {
                PackageFiles.Copy(Path.Combine(name is "LICENSE" or "CHANGELOG.md" ? root : Path.Combine(root, "mcpb"), name), Path.Combine(stage, name));
            }
            var archive = Path.Combine(Path.GetDirectoryName(stage)!, $"ExcelMcpMcpb-{Guid.NewGuid():N}.zip");
            try
            {
                _archive(stage, archive);
                using (var zip = ZipFile.OpenRead(archive))
                {
                    foreach (var name in McpbEntries)
                    {
                        if (zip.GetEntry(name) is null) { throw new InvalidOperationException($"MCPB is missing required metadata: {name}"); }
                    }
                }
                Directory.CreateDirectory(output);
                var destination = Path.Combine(output, $"excel-mcp-{version}.mcpb");
                PackageFiles.Install(archive, destination);
                PackageFiles.Install(Path.Combine(stage, "manifest.json"), Path.Combine(output, "manifest.json"));
                return destination;
            }
            finally { PackageFiles.Delete(archive); }
        }
        finally { PackageFiles.RemoveStaging(stage); }
    }

    public static void ValidateSkill(string directory, string version)
    {
        var path = Path.Combine(directory, "SKILL.md");
        if (!File.Exists(path)) { throw new InvalidOperationException($"Agent Skill is missing SKILL.md: {directory}"); }
        var stamp = Path.Combine(directory, "VERSION");
        if (!File.Exists(stamp)) { throw new InvalidOperationException($"Agent Skill is missing VERSION: {directory}"); }
        if (File.ReadAllText(stamp).Trim() != version) { throw new InvalidOperationException($"{stamp} has an unexpected version; expected '{version}'."); }
        var lines = File.ReadAllLines(path);
        if (lines.Length < 3 || lines[0].Trim() != "---") { throw new InvalidOperationException($"{path} must begin with YAML frontmatter."); }
        var end = System.Array.FindIndex(lines, 1, line => line.Trim() == "---");
        if (end < 2) { throw new InvalidOperationException($"{path} is missing the closing YAML frontmatter delimiter."); }
        var fields = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 1; index < end; index++)
        {
            var match = FrontmatterField().Match(lines[index]);
            if (!match.Success) { continue; }
            var field = match.Groups[1].Value;
            if (!SkillFields.Contains(field)) { throw new InvalidOperationException($"{path} contains unsupported Agent Skills frontmatter field '{field}'."); }
            var value = match.Groups[2].Value;
            if (field == "description" && value is ">" or "|")
            {
                value = string.Join(' ', lines[(index + 1)..end].TakeWhile(line => line.Length > 0 && char.IsWhiteSpace(line[0])).Select(line => line.Trim()));
            }
            fields[field] = value;
        }
        var name = fields.GetValueOrDefault("name") ?? "";
        if (name != Path.GetFileName(directory) || !SkillName().IsMatch(name) || name.Length > 64)
        {
            throw new InvalidOperationException($"{path} has invalid Agent Skill name '{name}'; it must match directory '{Path.GetFileName(directory)}'.");
        }
        if (!fields.TryGetValue("description", out var description)) { throw new InvalidOperationException($"{path} must declare an Agent Skill description."); }
        if (string.IsNullOrWhiteSpace(description) || description.Length > 1024) { throw new InvalidOperationException($"{path} has an invalid Agent Skill description length."); }
    }

    public static void ValidatePlugin(string name, string directory, string version)
    {
        var path = Path.Combine(directory, "plugin.json");
        var manifest = ReadJson(path);
        foreach (var property in manifest)
        {
            if (!PluginFields.Contains(property.Key)) { throw new InvalidOperationException($"{path} contains unsupported Agent Plugins 1.0 field '{property.Key}'."); }
        }
        const string schema = "https://agent-plugins.org/schemas/1.0.0/plugin.schema.json";
        if (Text(manifest, "$schema") != schema) { throw new InvalidOperationException($"{path} must target {schema}."); }
        if (Text(manifest, "name") != name || Text(manifest, "version") != version) { throw new InvalidOperationException($"{path} has unexpected name or version; expected '{name}' '{version}'."); }
        if (manifest["repository"] is not JsonValue repository || !repository.TryGetValue<string>(out _)) { throw new InvalidOperationException($"{path} repository must be a string."); }
        if (Directory.EnumerateFiles(directory, "install-global.ps1", SearchOption.AllDirectories).Any()) { throw new InvalidOperationException("Global installation helpers are retired; use npx instead."); }
        if (File.Exists(Path.Combine(directory, ".mcp.json"))) { throw new InvalidOperationException($"Legacy MCP configuration is not permitted in Agent Plugins 1.0 packages: {directory}"); }
        var skills = Path.Combine(directory, "skills");
        if (Directory.Exists(skills)) { foreach (var skill in Directory.EnumerateDirectories(skills)) { ValidateSkill(skill, version); } }
        var mcpPath = Path.Combine(directory, "mcp.json");
        if (!File.Exists(mcpPath)) { return; }
        var mcp = ReadJson(mcpPath);
        if (mcp.Count != 2 || !mcp.ContainsKey("$schema") || !mcp.ContainsKey("mcpServers")) { throw new InvalidOperationException($"{mcpPath} must contain only '$schema' and 'mcpServers'."); }
        if (Text(mcp, "$schema") != "https://agent-plugins.org/schemas/1.0.0/mcp.schema.json") { throw new InvalidOperationException($"{mcpPath} has an unexpected MCP schema."); }
        foreach (var server in Map(mcp, "mcpServers"))
        {
            var config = server.Value as JsonObject ?? throw new InvalidOperationException($"Malformed server '{server.Key}'.");
            if (config.Any(field => !ServerFields.Contains(field.Key))) { throw new InvalidOperationException($"{mcpPath} server '{server.Key}' contains unsupported fields."); }
            if (Text(config, "type") != "stdio" || Text(config, "command").Any(char.IsWhiteSpace)) { throw new InvalidOperationException($"{mcpPath} server '{server.Key}' must use type 'stdio' and a single executable command token."); }
            if (Array(config, "args").Any(argument => argument?.GetValue<string>().Contains("{pluginDir}", StringComparison.Ordinal) == true)) { throw new InvalidOperationException($"{mcpPath} server '{server.Key}' still uses the legacy '{{pluginDir}}' placeholder."); }
        }
    }

    public string PackageVersion() => Text(ReadJson(Path.Combine(root, "package.json")), "version");
    public static string RequiredVersion(string? version)
    {
        if (string.IsNullOrWhiteSpace(version) || !VersionPattern().IsMatch(version.Trim())) { throw new ArgumentException("Version is required. Pass a valid package version."); }
        return version.Trim();
    }
    public static JsonObject ReadJson(string path) => JsonNode.Parse(File.ReadAllText(path)) as JsonObject ?? throw new InvalidOperationException($"Expected a JSON object: {path}");
    public static void WriteJson(string path, JsonObject value) => File.WriteAllText(path, value.ToJsonString(JsonOptions));
    public static string Text(JsonObject value, string key) => value[key]?.GetValue<string>() ?? throw new InvalidOperationException($"Missing {key}.");
    public static JsonObject Map(JsonObject value, string key) => value[key] as JsonObject ?? throw new InvalidOperationException($"Missing object {key}.");
    public static JsonArray Array(JsonObject value, string key) => value[key] as JsonArray ?? throw new InvalidOperationException($"Missing array {key}.");
    public static string NewStage(string name)
    {
        var path = Path.Combine(Path.GetTempPath(), $"ExcelMcp{name}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        return path;
    }
    [GeneratedRegex(@"^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$", RegexOptions.CultureInvariant)] private static partial Regex VersionPattern();
    [GeneratedRegex(@"(?ms)^```(?<surface>cli|mcp)\n(?<body>.*?)^```[ \t]*(?:\n|$)", RegexOptions.CultureInvariant)] private static partial Regex SurfaceFences();
    [GeneratedRegex(@"\[([^\]]+)\]\(([a-z-]+)\.md(#[^)]*)?\)", RegexOptions.CultureInvariant)] private static partial Regex ReferenceLinks();
    [GeneratedRegex(@"^([a-z][a-z-]*):(?:\s*(.*))?$", RegexOptions.CultureInvariant)] private static partial Regex FrontmatterField();
    [GeneratedRegex(@"^(?!.*--)[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$", RegexOptions.CultureInvariant)] private static partial Regex SkillName();
}
