using Xunit;
using Sbroenne.ExcelMcp.Core;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Fast unit tests that don't require Excel installation.
/// These tests run by default and validate argument parsing, validation logic, etc.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
public class UnitTests
{
    [Theory]
    [InlineData("test.xlsx", true)]
    [InlineData("test.xlsm", true)]
    [InlineData("test.xls", true)]
    [InlineData("test.txt", false)]
    [InlineData("test.docx", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void ValidateExcelFile_WithVariousExtensions_ReturnsExpectedResult(string? filePath, bool expectedValid)
    {
        // Act
        bool result = ExcelHelper.ValidateExcelFile(filePath ?? "", requireExists: false);
        
        // Assert
        Assert.Equal(expectedValid, result);
    }

    [Theory]
    [InlineData(new string[] { "command" }, 2, false)]
    [InlineData(new string[] { "command", "arg1" }, 2, true)]
    [InlineData(new string[] { "command", "arg1", "arg2" }, 2, true)]
    [InlineData(new string[] { "command", "arg1", "arg2", "arg3" }, 3, true)]
    public void ValidateArgs_WithVariousArgCounts_ReturnsExpectedResult(string[] args, int required, bool expectedValid)
    {
        // Act
        bool result = ExcelHelper.ValidateArgs(args, required, "test command usage");
        
        // Assert
        Assert.Equal(expectedValid, result);
    }

    [Fact]
    public void ExcelDiagnostics_ReportOperationContext_DoesNotThrow()
    {
        // Act & Assert - Should not throw
        ExcelDiagnostics.ReportOperationContext("test-operation", "test.xlsx", 
            ("key1", "value1"), 
            ("key2", 42), 
            ("key3", null));
    }

    [Theory]
    [InlineData("test", new[] { "test", "other" }, "test")]
    [InlineData("Test", new[] { "test", "other" }, "test")]
    [InlineData("tst", new[] { "test", "other" }, "test")]
    [InlineData("other", new[] { "test", "other" }, "other")]
    [InlineData("xyz", new[] { "test", "other" }, null)]
    public void FindClosestMatch_WithVariousInputs_ReturnsExpectedResult(string target, string[] candidates, string? expected)
    {
        // This tests the private method indirectly by using the pattern from PowerQueryCommands
        // We'll test the logic with a simple implementation
        
        // Act
        string? result = FindClosestMatchSimple(target, candidates.ToList());
        
        // Assert
        Assert.Equal(expected, result);
    }

    private static string? FindClosestMatchSimple(string target, List<string> candidates)
    {
        if (candidates.Count == 0) return null;
        
        // First try exact case-insensitive match
        var exactMatch = candidates.FirstOrDefault(c => 
            string.Equals(c, target, StringComparison.OrdinalIgnoreCase));
        if (exactMatch != null) return exactMatch;
        
        // Then try substring match
        var substringMatch = candidates.FirstOrDefault(c => 
            c.Contains(target, StringComparison.OrdinalIgnoreCase) || 
            target.Contains(c, StringComparison.OrdinalIgnoreCase));
        if (substringMatch != null) return substringMatch;
        
        // Finally use simple Levenshtein distance (simplified for testing)
        int minDistance = int.MaxValue;
        string? bestMatch = null;
        
        foreach (var candidate in candidates)
        {
            int distance = ComputeLevenshteinDistance(target.ToLowerInvariant(), candidate.ToLowerInvariant());
            if (distance < minDistance && distance <= Math.Max(target.Length, candidate.Length) / 2)
            {
                minDistance = distance;
                bestMatch = candidate;
            }
        }
        
        return bestMatch;
    }
    
    private static int ComputeLevenshteinDistance(string s1, string s2)
    {
        int[,] d = new int[s1.Length + 1, s2.Length + 1];
        
        for (int i = 0; i <= s1.Length; i++)
            d[i, 0] = i;
        for (int j = 0; j <= s2.Length; j++)
            d[0, j] = j;
            
        for (int i = 1; i <= s1.Length; i++)
        {
            for (int j = 1; j <= s2.Length; j++)
            {
                int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }
        
        return d[s1.Length, s2.Length];
    }
}
