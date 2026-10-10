using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace Sbroenne.ExcelMcp.Build;

public sealed class SourceDependencies(string root, TestCatalog catalog)
{
    public IEnumerable<TestClass> AdapterConsumers(string owner, string path)
    {
        if (!path.EndsWith(".cs", StringComparison.Ordinal) || Path.GetFileName(path) is "Program.cs" or "GlobalUsings.cs")
        {
            return catalog.ForOwner(owner);
        }
        var source = catalog.ReadSource(path) ?? throw new InvalidOperationException($"No source ownership is available for {path}.");
        var changed = TestCatalog.DeclaredNames(source);
        var directory = Path.Combine(root, "src", $"ExcelMcp.{owner}");
        var sources = Directory.EnumerateFiles(directory, "*.cs", SearchOption.AllDirectories)
            .Where(file => !file.Split(Path.DirectorySeparatorChar).Any(segment => segment is "bin" or "obj"))
            .Where(file => Path.GetFileName(file) != "Program.cs")
            .Select(file =>
            {
                var content = File.ReadAllText(file);
                return (Names: TestCatalog.DeclaredNames(content), References: TestCatalog.ReferencedNames(content));
            }).ToArray();
        bool added;
        do
        {
            added = false;
            foreach (var item in sources.Where(item => item.References.Overlaps(changed)))
            {
                foreach (var name in item.Names) { added |= changed.Add(name); }
            }
        } while (added);
        var tools = CSharpSyntaxTree.ParseText(source).GetRoot().DescendantNodes().OfType<AttributeSyntax>()
            .Where(attribute => attribute.Name.ToString().EndsWith("McpServerTool", StringComparison.Ordinal))
            .SelectMany(attribute => attribute.ArgumentList?.Arguments ?? [])
            .Where(argument => argument.NameEquals?.Name.Identifier.ValueText == "Name")
            .Select(argument => argument.Expression as LiteralExpressionSyntax)
            .Where(value => value is not null).Select(value => value!.Token.ValueText).ToArray();
        var consumers = catalog.ForReferences(owner, changed).Concat(catalog.ForFeatures(owner, tools)).Distinct().ToArray();
        return consumers.Length > 0 ? consumers : throw new InvalidOperationException($"No validation mapping for adapter source {path}; add an actual consumer.");
    }

    public (string[] Areas, HashSet<string> Names) CommandConsumers(string path)
    {
        var source = catalog.ReadSource(path) ?? throw new InvalidOperationException($"No source ownership is available for {path}.");
        var changed = TestCatalog.DeclaredNames(source);
        var sources = Directory.EnumerateFiles(Path.Combine(root, "src", "ExcelMcp.Core"), "*.cs", SearchOption.AllDirectories)
            .Where(file => !file.Split(Path.DirectorySeparatorChar).Any(segment => segment is "bin" or "obj"))
            .Select(file =>
            {
                var content = File.ReadAllText(file);
                return (Path: Path.GetRelativePath(root, file).Replace('\\', '/'), Names: TestCatalog.DeclaredNames(content), References: TestCatalog.ReferencedNames(content));
            }).ToArray();
        var selected = new HashSet<string>(StringComparer.Ordinal);
        bool added;
        do
        {
            added = false;
            foreach (var item in sources.Where(item => item.Names.Overlaps(changed) || item.References.Overlaps(changed)))
            {
                selected.Add(item.Path);
                foreach (var name in item.Names) { added |= changed.Add(name); }
            }
        } while (added);
        var areas = selected.Where(file => file.StartsWith("src/ExcelMcp.Core/Commands/", StringComparison.Ordinal) && file.Split('/').Length > 4)
            .Select(file => file.Split('/')[3]).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        return (areas, changed);
    }
}
