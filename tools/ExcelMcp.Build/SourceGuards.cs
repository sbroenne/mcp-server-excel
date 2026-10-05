using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Build;

public sealed record SourceFinding(string Path, int Line, string Rule, string Guidance, string Text);

public sealed class SourceGuards(string root)
{
    private static readonly string[] ComInputs = [
        @"src\ExcelMcp.Core\Commands\Connection",
        @"src\ExcelMcp.Core\Commands\Table\TableCommands.Data.cs",
        @"src\ExcelMcp.Core\Commands\Table\TableCommands.Lifecycle.cs",
        @"src\ExcelMcp.Core\Commands\Table\TableCommands.Filters.cs",
        @"src\ExcelMcp.Core\Commands\Table\TableCommands.StructuredReferences.cs",
        @"src\ExcelMcp.Core\Commands\PivotTable\PivotTableCommands.Create.cs",
        @"src\ExcelMcp.Core\Commands\Chart\RegularChartStrategy.cs",
        @"src\ExcelMcp.Core\Commands\Chart\ChartCommands.Lifecycle.cs",
        @"src\ExcelMcp.Core\Commands\Chart\ChartCommands.DataSource.cs",
        @"src\ExcelMcp.Core\Commands\NamedRange\NamedRangeCommands.Operations.cs"
    ];
    private static readonly HashSet<string> DynamicExceptions = new(StringComparer.OrdinalIgnoreCase) {
        "ExcelBatch.cs", "ExcelSession.cs", "ExcelShutdownService.cs", "ExcelShutdownHelper.cs"
    };
    private static readonly HashSet<string> PackageExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".cs", ".csproj", ".props", ".targets"
    };
    private static readonly Regex ChainedCom = Pattern(
        @"\.(Rows|Columns|ListColumns|ListRows|Range|TableRange1|TableRange2|ChartArea|Parent|Application|Cells)\s*(?:\[[^\r\n]*\])?\s*\.(Count|Address|AutoFilter|Item|Range|Row|Rows|Column|Columns|Parent|Application|SeriesCollection|Formula|Name)\b");
    private static readonly Regex DiscardedName = Pattern(@"^\s*namesCollection\.Add\s*\(");
    private static readonly Regex SuccessTrue = Pattern(@"\.Success\s*=\s*true");
    private static readonly Regex SuccessAssignment = Pattern(@"\.Success\s*=");
    private static readonly Regex ErrorAssignment = Pattern("""\.ErrorMessage\s*=\s*["\$]""");
    private static readonly Regex DynamicCast = Pattern(@"\(\(dynamic\)");
    private static readonly Regex WorkbookPart = Pattern(
        """(?:\[Content_Types\]\.xml|(?:^|["'\\/])xl[\\/](?:workbook|worksheets|_rels|sharedStrings|styles|theme|connections|pivot|charts?)[^"'\r\n]*\.xml)""");
    private static readonly Regex OpenXml = Pattern(
        @"(?:DocumentFormat\.OpenXml|SpreadsheetDocument|OpenXmlPackage|WorkbookPart|WorksheetPart|System\.IO\.Packaging)");
    private static readonly Regex Zip = Pattern(@"(?:System\.IO\.Compression|ZipArchive|ZipFile)");
    private static readonly Regex Xml = Pattern(@"(?:System\.Xml|XDocument|XmlDocument|XmlReader|XElement)");
    private static readonly string[] Justifications = ["// PIA gap:", "// TODO:", "// Reason:", "// REASON:"];

    public IReadOnlyList<SourceFinding> Scan(string rule, IReadOnlyList<string>? inputs = null)
    {
        var files = Files(rule, inputs).Distinct(StringComparer.OrdinalIgnoreCase).Order(StringComparer.Ordinal).ToArray();
        if (files.Length == 0) { throw new InvalidOperationException($"No eligible source files found for the {rule} guard."); }
        var findings = new List<SourceFinding>();
        foreach (var file in files)
        {
            var lines = File.ReadAllLines(file);
            switch (rule)
            {
                case "com-leaks": ScanCom(file, lines, findings); break;
                case "success-flag": ScanSuccess(file, lines, findings); break;
                case "dynamic-casts": ScanDynamic(file, lines, findings); break;
                case "workbook-package-access": ScanPackages(file, string.Join('\n', lines), findings); break;
                default: throw new ArgumentException($"Unknown source guard: {rule}.");
            }
        }
        return findings;
    }

    public void Check(string rule, IReadOnlyList<string>? inputs = null)
    {
        var findings = Scan(rule, inputs);
        foreach (var finding in findings)
        {
            Console.Error.WriteLine($"{finding.Path}:{finding.Line} - {finding.Rule}");
            Console.Error.WriteLine($"  {finding.Text}");
            Console.Error.WriteLine($"  {finding.Guidance}");
        }
        if (findings.Count > 0)
        {
            throw new InvalidOperationException($"{findings.Count} {rule} violation(s).");
        }
        Console.WriteLine(rule switch
        {
            "com-leaks" => "No high-risk COM access patterns detected.",
            "success-flag" => "No nearby conflicting Success/ErrorMessage assignments found; this pattern check is not control-flow analysis.",
            "dynamic-casts" => "All ((dynamic)) casts are documented.",
            _ => "No production Excel workbook package XML access found."
        });
    }

    private IEnumerable<string> Files(string rule, IReadOnlyList<string>? inputs)
    {
        if (inputs is { Count: > 0 } && rule != "com-leaks")
        {
            throw new ArgumentException("Explicit input paths are supported only by the com-leaks guard.");
        }
        var paths = rule switch
        {
            "com-leaks" => inputs is { Count: > 0 } ? inputs : ComInputs,
            "success-flag" => [@"src\ExcelMcp.Core\Commands"],
            "dynamic-casts" => [@"src\ExcelMcp.Core", @"src\ExcelMcp.ComInterop"],
            "workbook-package-access" => ["src"],
            _ => throw new ArgumentException($"Unknown source guard: {rule}.")
        };
        foreach (var input in paths)
        {
            var path = Path.GetFullPath(input.Replace('\\', Path.DirectorySeparatorChar), root);
            IEnumerable<string> discovered;
            if (Directory.Exists(path))
            {
                discovered = Directory.EnumerateFiles(path, rule == "workbook-package-access" ? "*" : "*.cs", SearchOption.AllDirectories);
            }
            else if (File.Exists(path)) { discovered = [path]; }
            else if (rule == "com-leaks" && inputs is not { Count: > 0 }) { continue; }
            else { throw new FileNotFoundException($"Source input not found: {path}."); }
            var eligible = discovered.Where(file =>
                !file.Split(Path.DirectorySeparatorChar).Any(segment => segment.Equals("bin", StringComparison.OrdinalIgnoreCase) ||
                    segment.Equals("obj", StringComparison.OrdinalIgnoreCase)) &&
                (rule == "workbook-package-access" ? PackageExtensions.Contains(Path.GetExtension(file)) :
                    !file.EndsWith(".g.cs", StringComparison.OrdinalIgnoreCase))).ToArray();
            if (rule == "dynamic-casts" && eligible.Length == 0)
            {
                throw new InvalidOperationException($"No source files found for the dynamic cast guard: {path}.");
            }
            foreach (var file in eligible)
            {
                if (rule == "dynamic-casts" && DynamicExceptions.Contains(Path.GetFileName(file))) { continue; }
                yield return file;
            }
        }
    }

    private void ScanCom(string file, string[] lines, List<SourceFinding> findings)
    {
        for (var index = 0; index < lines.Length; index++)
        {
            var line = lines[index].TrimStart();
            if (line.StartsWith("//", StringComparison.Ordinal) || line.StartsWith('*')) { continue; }
            if (ChainedCom.IsMatch(line))
            {
                Add(file, index + 1, "chained COM property access", "Capture each COM object in a local variable and release it in finally.", line, findings);
            }
            if (DiscardedName.IsMatch(line))
            {
                Add(file, index + 1, "discarded Names.Add result", "Capture the returned Excel.Name and release it in finally.", line, findings);
            }
        }
    }

    private void ScanSuccess(string file, string[] lines, List<SourceFinding> findings)
    {
        for (var index = 0; index < lines.Length; index++)
        {
            if (!SuccessTrue.IsMatch(lines[index])) { continue; }
            for (var next = index + 1; next < Math.Min(index + 30, lines.Length); next++)
            {
                if (SuccessAssignment.IsMatch(lines[next])) { break; }
                if (!ErrorAssignment.IsMatch(lines[next]) ||
                    lines[next].Contains("= null", StringComparison.OrdinalIgnoreCase) ||
                    lines[next].Contains("= string.Empty", StringComparison.OrdinalIgnoreCase) ||
                    lines[next].Contains("= \"\"", StringComparison.OrdinalIgnoreCase)) { continue; }
                Add(file, next + 1, "Success/ErrorMessage conflict",
                    $"Success was set true at line {index + 1}; set it false before returning an error. This is a nearby-assignment check, not control-flow analysis.",
                    lines[next].Trim(), findings);
                break;
            }
        }
    }

    private void ScanDynamic(string file, string[] lines, List<SourceFinding> findings)
    {
        for (var index = 0; index < lines.Length; index++)
        {
            if (!DynamicCast.IsMatch(lines[index]) || lines[index].TrimStart().StartsWith("//", StringComparison.Ordinal)) { continue; }
            var justified = false;
            for (var previous = index - 1; previous >= 0 && previous >= index - 5; previous--)
            {
                var line = lines[previous].TrimStart();
                if (string.IsNullOrWhiteSpace(line)) { continue; }
                if (!line.StartsWith("//", StringComparison.Ordinal)) { break; }
                if (Justifications.Any(prefix => line.StartsWith(prefix, StringComparison.Ordinal)))
                {
                    justified = true;
                    break;
                }
            }
            if (!justified)
            {
                Add(file, index + 1, "undocumented ((dynamic)) cast",
                    "Document the PIA gap, TODO, or Reason in the preceding comment block.", lines[index].Trim(), findings);
            }
        }
    }

    private void ScanPackages(string file, string content, List<SourceFinding> findings)
    {
        const string guidance = "Use Excel COM in production. ZIP/OOXML access belongs only in tests.";
        foreach (var (pattern, rule) in new[] { (OpenXml, "Open XML or package API"), (WorkbookPart, "Excel workbook package part path") })
        {
            var match = pattern.Match(content);
            if (match.Success)
            {
                Add(file, content.AsSpan(0, match.Index).Count('\n') + 1, rule, guidance, match.Value, findings);
            }
        }
        if (Path.GetExtension(file).Equals(".cs", StringComparison.OrdinalIgnoreCase) && Zip.IsMatch(content) && Xml.IsMatch(content))
        {
            var match = Zip.Match(content);
            Add(file, content.AsSpan(0, match.Index).Count('\n') + 1, "combined ZIP and XML processing", guidance, match.Value, findings);
        }
    }

    private void Add(string file, int line, string rule, string guidance, string text, List<SourceFinding> findings) =>
        findings.Add(new SourceFinding(Path.GetRelativePath(root, file), line, rule, guidance, text));

    private static Regex Pattern(string pattern) =>
        new(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
}
