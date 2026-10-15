using System.Xml.Linq;

namespace Sbroenne.ExcelMcp.Build;

public sealed record PackageChanges(bool VersionOnly, string[] PackageIds, string[] Projects)
{
    private const string PackagesPath = "Directory.Packages.props";
    private const string ThisFileDirectory = "$(MSBuildThisFileDirectory)";

    public static async Task<PackageChanges> ReadAsync(string root, IProcessRunner runner,
        string? baseRef = null, string? headRef = null, bool staged = false, bool directComparison = false)
    {
        async Task<string> GitAsync(params string[] arguments) =>
            (await runner.CheckedAsync("git", arguments, TimeSpan.FromSeconds(30), preserveGitContext: staged)).Output.Trim();

        string comparisonBase;
        string? comparisonHead;
        if (staged)
        {
            var merge = await runner.RunAsync("git", ["rev-parse", "--verify", "--quiet", "MERGE_HEAD"],
                TimeSpan.FromSeconds(30), preserveGitContext: true);
            comparisonBase = merge.ExitCode == 0 ? merge.Output.Trim() : "HEAD";
            comparisonHead = "";
        }
        else if (baseRef is not null)
        {
            comparisonHead = await GitAsync("rev-parse", "--verify", $"{headRef ?? "HEAD"}^{{commit}}");
            comparisonBase = directComparison
                ? await GitAsync("rev-parse", "--verify", $"{baseRef}^{{commit}}")
                : await GitAsync("merge-base", "--", baseRef, comparisonHead);
        }
        else
        {
            comparisonBase = "HEAD";
            comparisonHead = null;
        }
        var before = await ReadSnapshotAsync(comparisonBase);
        var after = await ReadSnapshotAsync(comparisonHead);
        if (!before.TryGetValue(PackagesPath, out var basePackages) || !after.TryGetValue(PackagesPath, out var headPackages))
        {
            return new PackageChanges(false, [], []);
        }
        var baseDocument = XDocument.Parse(basePackages);
        var headDocument = XDocument.Parse(headPackages);
        var baseVersions = RemoveVersions(baseDocument);
        var headVersions = RemoveVersions(headDocument);
        if (baseVersions is null || headVersions is null || !XNode.DeepEquals(baseDocument, headDocument))
        {
            return new PackageChanges(false, [], []);
        }
        var changed = baseVersions.Keys.Concat(headVersions.Keys).Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(id => baseVersions.GetValueOrDefault(id) != headVersions.GetValueOrDefault(id))
            .Order(StringComparer.OrdinalIgnoreCase).ToArray();
        if (changed.Length == 0) { return new PackageChanges(true, [], []); }

        var affected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var dependencies = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var snapshot in new[] { before, after })
        {
            foreach (var project in snapshot.Keys.Where(path => path.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase)))
            {
                if (!dependencies.TryGetValue(project, out var references)) { dependencies[project] = references = new(StringComparer.OrdinalIgnoreCase); }
                foreach (var (path, document) in ProjectInputs(snapshot, project))
                {
                    foreach (var reference in document.Descendants().Where(element => element.Name.LocalName == "PackageReference"))
                    {
                        var ids = ReferenceIds(reference, path);
                        if (ids.Intersect(changed, StringComparer.OrdinalIgnoreCase).Any()) { affected.Add(project); }
                    }
                    foreach (var reference in document.Descendants().Where(element => element.Name.LocalName == "ProjectReference"))
                    {
                        foreach (var include in ReferenceIds(reference, path))
                        {
                            var declaringDirectory = include.StartsWith(ThisFileDirectory, StringComparison.OrdinalIgnoreCase);
                            references.Add(ResolvePath(declaringDirectory ? path : project,
                                declaringDirectory ? include[ThisFileDirectory.Length..].TrimStart('\\', '/') : include));
                        }
                    }
                }
            }
        }
        bool added;
        do
        {
            added = false;
            foreach (var (project, references) in dependencies)
            {
                if (references.Overlaps(affected)) { added |= affected.Add(project); }
            }
        } while (added);
        return new PackageChanges(true, changed, affected.Order(StringComparer.Ordinal).ToArray());

        async Task<Dictionary<string, string>> ReadSnapshotAsync(string? revision)
        {
            var names = revision is null
                ? await GitAsync("-c", "core.quotepath=false", "ls-files", "--cached", "--others", "--exclude-standard")
                : revision.Length == 0
                    ? await GitAsync("-c", "core.quotepath=false", "ls-files", "--cached")
                    : await GitAsync("-c", "core.quotepath=false", "ls-tree", "-r", "--name-only", revision);
            var snapshot = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in names.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries).Distinct(StringComparer.Ordinal))
            {
                if (path != PackagesPath && !path.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase) &&
                    !path.EndsWith(".props", StringComparison.OrdinalIgnoreCase) && !path.EndsWith(".targets", StringComparison.OrdinalIgnoreCase)) { continue; }
                if (path.Split('/').Any(segment => segment is "bin" or "obj")) { continue; }
                if (revision is null)
                {
                    var full = Path.Combine(root, path.Replace('/', Path.DirectorySeparatorChar));
                    if (File.Exists(full)) { snapshot.Add(path, await File.ReadAllTextAsync(full)); }
                }
                else
                {
                    snapshot.Add(path, await GitAsync("show", $"{revision}:{path}"));
                }
            }
            return snapshot;
        }
    }

    private static Dictionary<string, string>? RemoveVersions(XDocument document)
    {
        foreach (var comment in document.DescendantNodes().OfType<XComment>().ToArray()) { comment.Remove(); }
        if (document.Root?.Name != "Project") { return null; }
        var versions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in document.Descendants("PackageVersion").ToArray())
        {
            var id = item.Attribute("Include")?.Value;
            var version = item.Attribute("Version")?.Value;
            if (!IsLiteral(id) || !IsLiteral(version) || item.HasElements || item.Value.Length > 0 ||
                item.Attributes().Any(attribute => attribute.Name != "Include" && attribute.Name != "Version") ||
                item.Parent?.Name != "ItemGroup" || item.Parent.HasAttributes || item.Parent.Parent != document.Root ||
                !versions.TryAdd(id!, version!))
            {
                return null;
            }
            item.Remove();
        }
        foreach (var group in document.Root.Elements("ItemGroup").Where(group => !group.HasAttributes && !group.Nodes().Any()).ToArray())
        {
            group.Remove();
        }
        return versions;
    }

    private static IEnumerable<(string Path, XDocument Document)> ProjectInputs(Dictionary<string, string> snapshot, string project)
    {
        var pending = new Queue<string>();
        pending.Enqueue(project);
        for (var directory = Path.GetDirectoryName(project.Replace('/', Path.DirectorySeparatorChar)); ; directory = Path.GetDirectoryName(directory))
        {
            foreach (var name in new[] { "Directory.Build.props", "Directory.Build.targets" })
            {
                var path = string.IsNullOrEmpty(directory) ? name : $"{directory.Replace('\\', '/')}/{name}";
                if (snapshot.ContainsKey(path)) { pending.Enqueue(path); }
            }
            if (string.IsNullOrEmpty(directory)) { break; }
        }
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (pending.TryDequeue(out var path))
        {
            if (!visited.Add(path)) { continue; }
            var document = XDocument.Parse(snapshot[path]);
            if (document.Root?.Name.LocalName != "Project") { throw new InvalidOperationException($"Invalid project input: {path}."); }
            yield return (path, document);
            foreach (var import in document.Descendants().Where(element => element.Name.LocalName == "Import"))
            {
                var include = import.Attribute("Project")?.Value ?? throw new InvalidOperationException($"Import has no project: {path}.");
                // These repository imports are already covered by the ancestor inputs, or are untracked local settings.
                if (include == "$([MSBuild]::GetDirectoryNameOfFileAbove($(MSBuildThisFileDirectory).., 'Directory.Build.props'))\\Directory.Build.props" ||
                    include == "Directory.Build.props.user") { continue; }
                if (!IsLiteral(include)) { throw new InvalidOperationException($"Cannot resolve project import '{include}' in {path}."); }
                var imported = ResolvePath(path, include);
                if (!snapshot.ContainsKey(imported)) { throw new InvalidOperationException($"Cannot read project import {imported} from {path}."); }
                pending.Enqueue(imported);
            }
        }
    }

    private static string[] ReferenceIds(XElement reference, string path)
    {
        var value = reference.Attribute("Include")?.Value ?? reference.Attribute("Update")?.Value;
        if (string.IsNullOrWhiteSpace(value)) { throw new InvalidOperationException($"Cannot resolve {reference.Name} in {path}."); }
        var ids = value.Split(';', StringSplitOptions.TrimEntries);
        if (ids.Any(id => !IsLiteral(reference.Name.LocalName == "ProjectReference" &&
            id.StartsWith(ThisFileDirectory, StringComparison.OrdinalIgnoreCase) ? id[ThisFileDirectory.Length..] : id)))
        {
            throw new InvalidOperationException($"Cannot resolve {reference.Name} '{value}' in {path}.");
        }
        return ids;
    }

    private static bool IsLiteral(string? value) =>
        !string.IsNullOrWhiteSpace(value) && value.IndexOfAny(['$', '%', '@', '*', '?', ';']) < 0;

    private static string ResolvePath(string source, string include)
    {
        if (Path.IsPathRooted(include)) { throw new InvalidOperationException($"Project input must be repository-relative: {include}."); }
        var anchor = Path.GetFullPath(".");
        var full = Path.GetFullPath(include.Replace('\\', Path.DirectorySeparatorChar), Path.Combine(anchor, Path.GetDirectoryName(source)!));
        var relative = Path.GetRelativePath(anchor, full).Replace('\\', '/');
        if (relative == ".." || relative.StartsWith("../", StringComparison.Ordinal))
        {
            throw new InvalidOperationException($"Project input is outside the repository: {include} in {source}.");
        }
        return relative;
    }
}
