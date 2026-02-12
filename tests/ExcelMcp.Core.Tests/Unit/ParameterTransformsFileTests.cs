using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Unit tests for ParameterTransforms.ResolveValuesOrFile and ResolveFormulasOrFile.
/// These are pure utility methods (file I/O + parsing) — no COM interop required.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "ParameterTransforms")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ParameterTransformsFileTests : IDisposable
{
    private readonly string _tempDir;

    public ParameterTransformsFileTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_PT_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            Directory.Delete(_tempDir, recursive: true);
        }
    }

    private string CreateTempFile(string name, string content)
    {
        var path = Path.Combine(_tempDir, name);
        File.WriteAllText(path, content);
        return path;
    }

    // === ResolveValuesOrFile: Inline values take priority ===

    [Fact]
    public void ResolveValuesOrFile_InlineValues_ReturnsInlineDirectly()
    {
        var inline = new List<List<object?>> { new() { 1, 2 }, new() { 3, 4 } };
        var result = ParameterTransforms.ResolveValuesOrFile(inline, null);

        Assert.Same(inline, result);
    }

    [Fact]
    public void ResolveValuesOrFile_InlineValues_IgnoresFile()
    {
        var inline = new List<List<object?>> { new() { "A" } };
        // File doesn't even exist — should not matter
        var result = ParameterTransforms.ResolveValuesOrFile(inline, @"C:\nonexistent\file.json");

        Assert.Same(inline, result);
    }

    // === ResolveValuesOrFile: Neither provided ===

    [Fact]
    public void ResolveValuesOrFile_NeitherProvided_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveValuesOrFile(null, null));

        Assert.Contains("values", ex.Message);
        Assert.Contains("valuesFile", ex.Message);
    }

    [Fact]
    public void ResolveValuesOrFile_EmptyListAndNoFile_ThrowsArgumentException()
    {
        var empty = new List<List<object?>>();
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveValuesOrFile(empty, null));

        Assert.Contains("values", ex.Message);
    }

    // === ResolveValuesOrFile: File not found ===

    [Fact]
    public void ResolveValuesOrFile_FileNotFound_ThrowsFileNotFoundException()
    {
        var missingPath = Path.Combine(_tempDir, "does_not_exist.json");

        var ex = Assert.Throws<FileNotFoundException>(
            () => ParameterTransforms.ResolveValuesOrFile(null, missingPath));

        Assert.Contains(missingPath, ex.Message);
    }

    // === ResolveValuesOrFile: JSON file ===

    [Fact]
    public void ResolveValuesOrFile_JsonFile_Parses2DArray()
    {
        var path = CreateTempFile("data.json", "[[1,2,3],[4,5,6]]");

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal(3, result[0].Count);
        Assert.Equal(3, result[1].Count);
    }

    [Fact]
    public void ResolveValuesOrFile_JsonFile_PreservesStringValues()
    {
        var path = CreateTempFile("strings.json", "[[\"Hello\",\"World\"],[\"Foo\",\"Bar\"]]");

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal("Hello", result[0][0]?.ToString());
        Assert.Equal("Bar", result[1][1]?.ToString());
    }

    [Fact]
    public void ResolveValuesOrFile_JsonFile_PreservesNulls()
    {
        var path = CreateTempFile("nulls.json", "[[1,null,3],[null,5,null]]");

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Null(result[0][1]);
        Assert.Null(result[1][0]);
        Assert.Null(result[1][2]);
    }

    [Fact]
    public void ResolveValuesOrFile_JsonFile_MixedTypes()
    {
        var path = CreateTempFile("mixed.json", "[[\"Name\",\"Age\"],[\"Alice\",30],[\"Bob\",25]]");

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(3, result.Count);
        Assert.Equal("Name", result[0][0]?.ToString());
        Assert.Equal("Age", result[0][1]?.ToString());
    }

    [Fact]
    public void ResolveValuesOrFile_InvalidJson_ThrowsArgumentException()
    {
        var path = CreateTempFile("bad.json", "{ this is not valid }");

        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveValuesOrFile(null, path));

        Assert.Contains("Invalid JSON", ex.Message);
        Assert.Contains("bad.json", ex.Message);
    }

    // === ResolveValuesOrFile: CSV file ===

    [Fact]
    public void ResolveValuesOrFile_CsvFile_ParsesRowsAndColumns()
    {
        var csv = "Alice,30,Engineering\nBob,25,Marketing\nCharlie,35,Sales";
        var path = CreateTempFile("data.csv", csv);

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(3, result.Count);
        Assert.Equal(3, result[0].Count);
        Assert.Equal("Alice", result[0][0]);
        Assert.Equal("30", result[0][1]);
        Assert.Equal("Engineering", result[0][2]);
        Assert.Equal("Bob", result[1][0]);
    }

    [Fact]
    public void ResolveValuesOrFile_CsvFile_HandlesQuotedValues()
    {
        var csv = "\"Alice\",\"30\"\n\"Bob\",\"25\"";
        var path = CreateTempFile("quoted.csv", csv);

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal("Alice", result[0][0]);
        Assert.Equal("30", result[0][1]);
    }

    [Fact]
    public void ResolveValuesOrFile_CsvFile_SkipsEmptyLines()
    {
        var csv = "A,B\n\nC,D\n\n";
        var path = CreateTempFile("gaps.csv", csv);

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal("A", result[0][0]);
        Assert.Equal("C", result[1][0]);
    }

    [Fact]
    public void ResolveValuesOrFile_NonJsonExtension_TreatedAsCsv()
    {
        var csv = "1,2,3\n4,5,6";
        var path = CreateTempFile("data.txt", csv);

        var result = ParameterTransforms.ResolveValuesOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal("1", result[0][0]);
    }

    [Fact]
    public void ResolveValuesOrFile_EmptyCsvFile_ThrowsArgumentException()
    {
        var path = CreateTempFile("empty.csv", "   ");

        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveValuesOrFile(null, path));

        Assert.Contains("empty or contains no parseable data", ex.Message);
    }

    // === ResolveValuesOrFile: Custom parameterName ===

    [Fact]
    public void ResolveValuesOrFile_CustomParameterName_UsedInErrorMessage()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveValuesOrFile(null, null, "rows"));

        Assert.Contains("rows", ex.Message);
        Assert.Contains("rowsFile", ex.Message);
    }

    // === ResolveFormulasOrFile: Inline formulas take priority ===

    [Fact]
    public void ResolveFormulasOrFile_InlineFormulas_ReturnsInlineDirectly()
    {
        var inline = new List<List<string>> { new() { "=A1+B1", "=SUM(A:A)" } };
        var result = ParameterTransforms.ResolveFormulasOrFile(inline, null);

        Assert.Same(inline, result);
    }

    // === ResolveFormulasOrFile: Neither provided ===

    [Fact]
    public void ResolveFormulasOrFile_NeitherProvided_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveFormulasOrFile(null, null));

        Assert.Contains("formulas", ex.Message);
        Assert.Contains("formulasFile", ex.Message);
    }

    [Fact]
    public void ResolveFormulasOrFile_EmptyListAndNoFile_ThrowsArgumentException()
    {
        var empty = new List<List<string>>();
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveFormulasOrFile(empty, null));

        Assert.Contains("formulas", ex.Message);
    }

    // === ResolveFormulasOrFile: File not found ===

    [Fact]
    public void ResolveFormulasOrFile_FileNotFound_ThrowsFileNotFoundException()
    {
        var missingPath = Path.Combine(_tempDir, "no_such_file.json");

        var ex = Assert.Throws<FileNotFoundException>(
            () => ParameterTransforms.ResolveFormulasOrFile(null, missingPath));

        Assert.Contains(missingPath, ex.Message);
    }

    // === ResolveFormulasOrFile: JSON file ===

    [Fact]
    public void ResolveFormulasOrFile_JsonFile_Parses2DStringArray()
    {
        var json = "[[\"=A1+B1\",\"=C1*2\"],[\"=SUM(A:A)\",\"=AVERAGE(B:B)\"]]";
        var path = CreateTempFile("formulas.json", json);

        var result = ParameterTransforms.ResolveFormulasOrFile(null, path);

        Assert.Equal(2, result.Count);
        Assert.Equal("=A1+B1", result[0][0]);
        Assert.Equal("=C1*2", result[0][1]);
        Assert.Equal("=SUM(A:A)", result[1][0]);
        Assert.Equal("=AVERAGE(B:B)", result[1][1]);
    }

    [Fact]
    public void ResolveFormulasOrFile_InvalidJson_ThrowsArgumentException()
    {
        var path = CreateTempFile("bad_formulas.json", "not json at all");

        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveFormulasOrFile(null, path));

        Assert.Contains("Invalid JSON", ex.Message);
    }

    // === ResolveFormulasOrFile: Custom parameterName ===

    [Fact]
    public void ResolveFormulasOrFile_CustomParameterName_UsedInErrorMessage()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => ParameterTransforms.ResolveFormulasOrFile(null, null, "formats"));

        Assert.Contains("formats", ex.Message);
        Assert.Contains("formatsFile", ex.Message);
    }

    // === ParseCsvToRows ===

    [Fact]
    public void ParseCsvToRows_NullInput_ReturnsNull()
    {
        var result = ParameterTransforms.ParseCsvToRows(null);
        Assert.Null(result);
    }

    [Fact]
    public void ParseCsvToRows_WhitespaceOnly_ReturnsNull()
    {
        var result = ParameterTransforms.ParseCsvToRows("   ");
        Assert.Null(result);
    }

    [Fact]
    public void ParseCsvToRows_SingleRow_ReturnsOneRow()
    {
        var result = ParameterTransforms.ParseCsvToRows("A,B,C");

        Assert.NotNull(result);
        Assert.Single(result);
        Assert.Equal(3, result[0].Count);
        Assert.Equal("A", result[0][0]);
        Assert.Equal("B", result[0][1]);
        Assert.Equal("C", result[0][2]);
    }

    [Fact]
    public void ParseCsvToRows_MultipleRows_ParsesCorrectly()
    {
        var result = ParameterTransforms.ParseCsvToRows("1,2\n3,4\n5,6");

        Assert.NotNull(result);
        Assert.Equal(3, result.Count);
        Assert.Equal("1", result[0][0]);
        Assert.Equal("6", result[2][1]);
    }

    [Fact]
    public void ParseCsvToRows_EmptyCells_TreatedAsNull()
    {
        var result = ParameterTransforms.ParseCsvToRows("A,,C\n,B,");

        Assert.NotNull(result);
        Assert.Null(result[0][1]);
        Assert.Null(result[1][0]);
        Assert.Null(result[1][2]);
    }
}
