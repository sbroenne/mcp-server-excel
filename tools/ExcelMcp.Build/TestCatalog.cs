using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Sbroenne.ExcelMcp.Build;

public sealed class TestCatalog
{
    private readonly List<TestClass> _classes = [];
    private readonly List<SourceType> _types = [];
    private readonly string _root;
    private readonly Dictionary<string, string?> _previous = new(StringComparer.Ordinal);

    public TestCatalog(string root)
    {
        _root = root;
        foreach (var project in Directory.EnumerateFiles(Path.Combine(root, "tests"), "*.csproj", SearchOption.AllDirectories))
        {
            var projectName = Path.GetFileNameWithoutExtension(project);
            if (!projectName.StartsWith("ExcelMcp.", StringComparison.Ordinal) || !projectName.EndsWith(".Tests", StringComparison.Ordinal)) { continue; }
            var owner = projectName["ExcelMcp.".Length..^".Tests".Length];
            ReadProject(owner, project);
        }
    }

    private void ReadProject(string owner, string project)
    {
        var directory = Path.GetDirectoryName(project)!;
        var document = XDocument.Load(project);
        var files = Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories).Where(IsSource).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var include in document.Descendants("Compile").Attributes("Include"))
        {
            var path = Path.GetFullPath(include.Value.Replace('\\', Path.DirectorySeparatorChar), directory);
            if (File.Exists(path)) { files.Add(path); }
        }
        var options = CSharpParseOptions.Default.WithPreprocessorSymbols(ProjectSymbols(project));
        if (document.Descendants("Compile").Any(item => item.Attribute("Remove") is not null || item.Attribute("Condition") is not null ||
            item.Parent?.Attribute("Condition") is not null || item.Attribute("Include")?.Value.Contains('*', StringComparison.Ordinal) == true))
        {
            var result = new ProcessRunner(directory).CheckedAsync("dotnet",
                ["msbuild", project, "-p:Configuration=Release", "-getItem:Compile", "-getProperty:DefineConstants"], TimeSpan.FromMinutes(2)).GetAwaiter().GetResult();
            using var evaluated = JsonDocument.Parse(result.Output);
            files = evaluated.RootElement.GetProperty("Items").GetProperty("Compile").EnumerateArray()
                .Select(item => item.GetProperty("FullPath").GetString() ?? throw new InvalidOperationException($"Compile input has no path: {project}"))
                .Where(IsSource).ToHashSet(StringComparer.OrdinalIgnoreCase);
            options = options.WithPreprocessorSymbols(evaluated.RootElement.GetProperty("Properties").GetProperty("DefineConstants").GetString()!
                .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries));
        }
        var groups = files.SelectMany(file => Parse(File.ReadAllText(file), file, options))
            .GroupBy(part => FullName(part.Node, part.Name), StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var names = groups.Values.ToLookup(parts => parts[0].Name, StringComparer.Ordinal);
        var collections = groups.Values.Where(parts => parts.SelectMany(part => Attributes(part.Node)).Any(attribute => AttributeName(attribute) == "CollectionDefinition"))
            .SelectMany(parts => parts.SelectMany(part => Attributes(part.Node))
                .Where(attribute => AttributeName(attribute) == "CollectionDefinition")
                .Select(attribute => (Name: ArgumentText(attribute, 0), Type: parts[0].Name)))
            .Where(collection => collection.Name is not null).ToArray();
        foreach (var parts in groups.Values)
        {
            var first = parts[0];
            var ancestors = Ancestors(parts, names).ToArray();
            var references = References(parts.Concat(ancestors.SelectMany(part => part))).ToHashSet(StringComparer.Ordinal);
            foreach (var collection in parts.SelectMany(part => Attributes(part.Node)).Where(attribute => AttributeName(attribute) == "Collection"))
            {
                foreach (var definition in collections.Where(definition => definition.Name == ArgumentText(collection, 0))) { references.Add(definition.Type); }
            }
            var sources = parts.Concat(ancestors.SelectMany(part => part)).Select(part => part.File).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (!first.Node.Ancestors().OfType<TypeDeclarationSyntax>().Any())
            {
                _types.Add(new SourceType(owner, first.Name, sources, references));
            }
            if (first.Node is not TypeDeclarationSyntax || parts.Any(part => part.Node is TypeDeclarationSyntax type && type.Modifiers.Any(SyntaxKind.AbstractKeyword))) { continue; }
            var classTraits = parts.Concat(ancestors.SelectMany(part => part)).SelectMany(part => ReadTraits(Attributes(part.Node))).ToArray();
            var methods = parts.Concat(ancestors.SelectMany(part => part))
                .SelectMany(part => (part.Node as TypeDeclarationSyntax)?.Members.OfType<MethodDeclarationSyntax>() ?? [])
                .Where(method => IsTestMethod(method, names)).ToArray();
            if (methods.Length == 0) { continue; }
            var traits = classTraits.Concat(methods.SelectMany(method => ReadTraits(method.AttributeLists.SelectMany(list => list.Attributes)))).ToArray();
            var features = traits.Where(trait => trait.Name == "Feature").Select(trait => trait.Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var normal = methods.Where(method => !MethodTraits(method).Any(trait => trait.Name == "RunType" && trait.Value == "OnDemand")).ToArray();
            var classifications = classTraits.Concat(normal.SelectMany(method => ReadTraits(method.AttributeLists.SelectMany(list => list.Attributes))))
                .Where(trait => trait.Name == "RequiresExcel").Select(trait => trait.Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (normal.Length > 0 && classifications.Count == 0) { throw new InvalidOperationException($"Missing Excel classification for {FullName(first.Node, first.Name)}."); }
            _classes.Add(new TestClass(owner, first.Name, FullName(first.Node, first.Name), features,
                normal.Length > 0 && classifications.Contains("false"), normal.Length > 0 && classifications.Contains("true"),
                traits.Any(trait => trait.Name == "AdapterTestKind" && trait.Value == "System"), sources, references,
                normal.Length > 0 && normal.All(method => MethodTraits(method).Any(trait => trait.Name == "Acceptance" && trait.Value == "Required"))));

            IEnumerable<(string Name, string Value)> MethodTraits(MethodDeclarationSyntax method) =>
                classTraits.Concat(ReadTraits(method.AttributeLists.SelectMany(list => list.Attributes)));
        }
    }

    public IEnumerable<TestClass> ForFeatures(string owner, IEnumerable<string> features)
    {
        var selected = features.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return _classes.Where(type => type.Owner == owner && (type.Features.Overlaps(selected) ||
            selected.Any(feature => type.Name.StartsWith(feature, StringComparison.Ordinal)) ||
            selected.Any(feature => type.References.Any(reference => reference.StartsWith(feature, StringComparison.Ordinal) && reference.EndsWith("Commands", StringComparison.Ordinal)))));
    }

    public IEnumerable<TestClass> ForOwner(string owner) => _classes.Where(type => type.Owner == owner);

    public IEnumerable<TestClass> ForFile(string path)
    {
        var full = Path.GetFullPath(path.Replace('/', Path.DirectorySeparatorChar), _root);
        var direct = _classes.Where(type => type.Files.Contains(full)).ToArray();
        var changed = _types.Where(type => type.Files.Contains(full)).ToArray();
        if (direct.Length > 0 && changed.All(type => direct.Any(test => test.Owner == type.Owner && test.Name == type.Name))) { return direct; }
        if (changed.Length == 0 && !File.Exists(full))
        {
            var source = PreviousSource(path);
            if (source is null) { return []; }
            var parts = Parse(source, full, CSharpParseOptions.Default.WithPreprocessorSymbols("TRACE")).ToArray();
            var names = parts.Where(part => !part.Node.Ancestors().OfType<TypeDeclarationSyntax>().Any()).Select(part => part.Name).ToHashSet(StringComparer.Ordinal);
            var owner = path.Split('/').ElementAtOrDefault(1)?.Replace("ExcelMcp.", "", StringComparison.Ordinal).Replace(".Tests", "", StringComparison.Ordinal);
            if (owner is null) { return []; }
            var survivors = ForReferences(owner, names).ToArray();
            if (survivors.Length > 0) { return survivors; }
            var features = parts.SelectMany(part => ReadTraits(Attributes(part.Node))).Where(trait => trait.Name == "Feature").Select(trait => trait.Value);
            return ForFeatures(owner, features).Concat(ForOwner(owner).Where(type => type.Name == "TestClassificationArchitectureTests")).Distinct();
        }
        return ForDependencies(changed.Select(type => (type.Owner, type.Name)));
    }

    public IEnumerable<TestClass> ForReferences(string owner, IEnumerable<string> names)
    {
        var selected = names.ToHashSet(StringComparer.Ordinal);
        var direct = _types.Where(type => type.Owner == owner && (selected.Contains(type.Name) || type.References.Overlaps(selected)));
        return ForDependencies(direct.Select(type => (type.Owner, type.Name)));
    }

    public string? ReadSource(string path)
    {
        var full = Path.GetFullPath(path.Replace('/', Path.DirectorySeparatorChar), _root);
        return File.Exists(full) ? File.ReadAllText(full) : PreviousSource(path);
    }

    public static HashSet<string> DeclaredNames(string source) => Parse(source, "", CSharpParseOptions.Default.WithPreprocessorSymbols("TRACE"))
        .Where(part => !part.Node.Ancestors().OfType<TypeDeclarationSyntax>().Any()).Select(part => part.Name).ToHashSet(StringComparer.Ordinal);

    public static HashSet<string> ReferencedNames(string source) => CSharpSyntaxTree.ParseText(source, CSharpParseOptions.Default.WithPreprocessorSymbols("TRACE"))
        .GetRoot().DescendantNodes().OfType<SimpleNameSyntax>().Select(name => name.Identifier.ValueText).ToHashSet(StringComparer.Ordinal);

    private IEnumerable<TestClass> ForDependencies(IEnumerable<(string Owner, string Name)> initial)
    {
        var dependencies = initial.ToHashSet();
        bool added;
        do
        {
            added = false;
            foreach (var type in _types)
            {
                if (dependencies.Contains((type.Owner, type.Name))) { continue; }
                if (dependencies.Any(dependency => dependency.Owner == type.Owner && type.References.Contains(dependency.Name))) { added |= dependencies.Add((type.Owner, type.Name)); }
            }
        } while (added);
        return _classes.Where(type => dependencies.Contains((type.Owner, type.Name)));
    }

    private string? PreviousSource(string path)
    {
        if (_previous.TryGetValue(path, out var source)) { return source; }
        var git = new ProcessRunner(_root);
        var result = git.RunAsync("git", ["show", $"HEAD:{path}"], TimeSpan.FromSeconds(30)).GetAwaiter().GetResult();
        if (result.ExitCode == 0) { return _previous[path] = result.Output; }
        var history = git.CheckedAsync("git", ["log", "-1", "--format=%H", "--", path], TimeSpan.FromSeconds(30)).GetAwaiter().GetResult();
        if (string.IsNullOrWhiteSpace(history.Output)) { return _previous[path] = null; }
        result = git.RunAsync("git", ["show", $"{history.Output.Trim()}^:{path}"], TimeSpan.FromSeconds(30)).GetAwaiter().GetResult();
        if (result.ExitCode != 0) { throw new InvalidOperationException($"Cannot read previous ownership for {path}: {result.Error}"); }
        return _previous[path] = result.Output;
    }

    private static IEnumerable<Part[]> Ancestors(Part[] parts, ILookup<string, Part[]> names)
    {
        var visited = new HashSet<string>(StringComparer.Ordinal) { parts[0].Name };
        var pending = new Queue<string>(BaseNames(parts));
        while (pending.TryDequeue(out var name))
        {
            if (!visited.Add(name)) { continue; }
            foreach (var ancestor in names[name])
            {
                yield return ancestor;
                foreach (var parent in BaseNames(ancestor)) { pending.Enqueue(parent); }
            }
        }
    }
    private static IEnumerable<string> BaseNames(IEnumerable<Part> parts) =>
        parts.SelectMany(part => (part.Node as TypeDeclarationSyntax)?.BaseList?.Types ?? [])
            .Select(type => type.Type is GenericNameSyntax generic ? generic.Identifier.ValueText : type.Type.ToString().Split('.').Last());

    private static bool IsTestMethod(MethodDeclarationSyntax method, ILookup<string, Part[]> names) =>
        method.AttributeLists.SelectMany(list => list.Attributes).Any(attribute =>
        {
            var name = AttributeName(attribute);
            if (name is "Fact" or "Theory") { return true; }
            var declaration = names[name + "Attribute"].Concat(names[name]);
            return declaration.Any(parts => Ancestors(parts, names).SelectMany(BaseNames).Concat(BaseNames(parts)).Any(parent => parent is "FactAttribute" or "TheoryAttribute"));
        });

    private static IEnumerable<string> References(IEnumerable<Part> parts) =>
        parts.SelectMany(part => part.Node.DescendantNodes().OfType<SimpleNameSyntax>()
            .Concat(part.Node.SyntaxTree.GetRoot().DescendantNodes().OfType<UsingDirectiveSyntax>()
                .Where(directive => directive.StaticKeyword.IsKind(SyntaxKind.StaticKeyword) || directive.Alias is not null)
                .SelectMany(directive => directive.DescendantNodes().OfType<SimpleNameSyntax>())))
            .Select(identifier => identifier.Identifier.ValueText);

    private static IEnumerable<Part> Parse(string source, string file, CSharpParseOptions options)
    {
        foreach (var node in CSharpSyntaxTree.ParseText(source, options).GetRoot().DescendantNodes())
        {
            var name = node switch
            {
                BaseTypeDeclarationSyntax type => type.Identifier.ValueText,
                DelegateDeclarationSyntax declaration => declaration.Identifier.ValueText,
                _ => null
            };
            if (name is not null) { yield return new Part(file, node, name); }
        }
    }
    private static IEnumerable<AttributeSyntax> Attributes(SyntaxNode node) => node switch
    {
        BaseTypeDeclarationSyntax type => type.AttributeLists.SelectMany(list => list.Attributes),
        DelegateDeclarationSyntax declaration => declaration.AttributeLists.SelectMany(list => list.Attributes),
        _ => []
    };
    private static string FullName(SyntaxNode node, string name) =>
        string.Join('.', node.Ancestors().Reverse().Select(parent => parent switch
        {
            BaseNamespaceDeclarationSyntax space => space.Name.ToString(),
            TypeDeclarationSyntax type => type.Identifier.ValueText,
            _ => null
        }).Where(part => part is not null).Append(name));

    private static string AttributeName(AttributeSyntax attribute)
    {
        var name = attribute.Name.ToString().Split('.').Last();
        var alias = attribute.SyntaxTree.GetRoot().DescendantNodes().OfType<UsingDirectiveSyntax>()
            .FirstOrDefault(directive => directive.Alias?.Name.Identifier.ValueText == name);
        if (alias?.Name is { } target) { name = target.ToString().Split('.').Last(); }
        return name.EndsWith("Attribute", StringComparison.Ordinal) ? name[..^"Attribute".Length] : name;
    }
    private static string? ArgumentText(AttributeSyntax attribute, int index) =>
        attribute.ArgumentList?.Arguments.ElementAtOrDefault(index)?.Expression is LiteralExpressionSyntax value ? value.Token.ValueText : null;
    private static IEnumerable<(string Name, string Value)> ReadTraits(IEnumerable<AttributeSyntax> attributes)
    {
        foreach (var attribute in attributes)
        {
            if (AttributeName(attribute) == "Trait" && ArgumentText(attribute, 0) is { } name && ArgumentText(attribute, 1) is { } value) { yield return (name, value); }
        }
    }
    private static bool IsSource(string path) => !path.Split(Path.DirectorySeparatorChar).Any(segment => segment is "bin" or "obj");
    private static HashSet<string> ProjectSymbols(string project)
    {
        var symbols = new HashSet<string>(StringComparer.Ordinal) { "TRACE", "NET", "NETCOREAPP", "NET10_0", "NET10_0_OR_GREATER" };
        var inputs = new List<string>();
        for (var directory = new DirectoryInfo(Path.GetDirectoryName(project)!); directory is not null; directory = directory.Parent)
        {
            var props = Path.Combine(directory.FullName, "Directory.Build.props");
            if (File.Exists(props)) { inputs.Insert(0, props); }
        }
        inputs.Add(project);
        var framework = "net10.0-windows";
        foreach (var input in inputs)
        {
            foreach (var value in XDocument.Load(input).Descendants().Where(element => element.Name.LocalName is "TargetFramework" or "DefineConstants"))
            {
                var condition = value.Attribute("Condition")?.Value ?? value.Parent?.Attribute("Condition")?.Value;
                if (condition is not null)
                {
                    if (condition.Replace(" ", "", StringComparison.Ordinal) is "'$(TargetFramework)'==''" && value.Name.LocalName == "TargetFramework") { continue; }
                    var release = Regex.Replace(condition, @"\$\((Configuration|OS)\)", match => match.Groups[1].Value == "Configuration" ? "Release" : OperatingSystem.IsWindows() ? "Windows_NT" : "Unix");
                    var comparison = Regex.Match(release, @"^\s*'([^']*)'\s*(==|!=)\s*'([^']*)'\s*$");
                    if (!comparison.Success) { throw new InvalidOperationException($"Unsupported compile-symbol condition in {input}: {condition}"); }
                    if ((comparison.Groups[1].Value == comparison.Groups[3].Value) != (comparison.Groups[2].Value == "==")) { continue; }
                }
                if (value.Name.LocalName == "TargetFramework") { framework = value.Value; }
                else { foreach (var symbol in value.Value.Replace("$(DefineConstants)", "", StringComparison.Ordinal).Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)) { symbols.Add(symbol); } }
            }
        }
        if (framework.Contains("-windows", StringComparison.OrdinalIgnoreCase)) { symbols.Add("WINDOWS"); }
        return symbols;
    }
    private sealed record Part(string File, SyntaxNode Node, string Name);
    private sealed record SourceType(string Owner, string Name, HashSet<string> Files, HashSet<string> References);
}

public sealed record TestClass(
    string Owner, string Name, string FullName, HashSet<string> Features, bool ExcelFree, bool Excel, bool System,
    HashSet<string> Files, HashSet<string> References, bool RequiredOnly);
