using Sbroenne.ExcelMcp.Build;
using Xunit;

namespace Sbroenne.ExcelMcp.SkillGeneration.Tests;

[Collection("Sequential")]
[Trait("RequiresExcel", "false")]
[Trait("Category", "Integration")]
[Trait("Feature", "SkillGeneration")]
public sealed class SkillExampleTests
{
    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    public void SkillGeneration_SelectsNativeExamplesWithoutTranslatingContent(string newline)
    {
        var root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"ExcelMcp.SkillExamples.{Guid.NewGuid():N}")).FullName;
        try
        {
            var shared = Directory.CreateDirectory(Path.Combine(root, "docs", "reference")).FullName;
            const string document = """
                # Workflow
                Shared policy uses sessionId responses.
                ```cli
                excelcli -q range get-values --session $sessionId --sheet Sales --range A1
                ```
                ```mcp
                range(action: 'get-values', workbook_session_id: sessionId, sheet_name: 'Sales', range_address: 'A1')
                ```
                ```json
                {"mCodeFile":"query.m"}
                ```
                """;
            File.WriteAllText(Path.Combine(shared, "report-formatting.md"), document.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\n", newline, StringComparison.Ordinal));
            var launcher = Directory.CreateDirectory(Path.Combine(root, "skills", "excel-cli")).FullName;
            File.WriteAllText(Path.Combine(launcher, "SKILL.md"), "Authored CLI discovery");
            foreach (var surface in new[] { "cli", "mcp" })
            {
                var source = Directory.CreateDirectory(Path.Combine(root, "skills", $"excel-{surface}-report-formatting")).FullName;
                File.WriteAllText(Path.Combine(source, "SKILL.md"), "Authored skill");
            }
            var output = new AuthoredPackages(root).Skills("1.2.3", null, generateOnly: true, prepared: null);
            Assert.Equal(["excel-cli", "excel-cli-report-formatting", "excel-mcp-report-formatting"],
                Directory.GetDirectories(output).Select(Path.GetFileName).Order(StringComparer.Ordinal));
            Assert.Equal("Authored CLI discovery", File.ReadAllText(Path.Combine(output, "excel-cli", "SKILL.md")));
            Assert.Equal("1.2.3", File.ReadAllText(Path.Combine(output, "excel-cli", "VERSION")));
            Assert.False(Directory.Exists(Path.Combine(output, "excel-cli", "references")));
            foreach (var surface in new[] { "cli", "mcp" })
            {
                var content = File.ReadAllText(Path.Combine(output, $"excel-{surface}-report-formatting", "references", "report-formatting.md"));
                Assert.Contains("Shared policy uses sessionId responses.", content, StringComparison.Ordinal);
                Assert.Contains("""{"mCodeFile":"query.m"}""", content, StringComparison.Ordinal);
                Assert.DoesNotContain("```cli", content, StringComparison.Ordinal);
                Assert.DoesNotContain("```mcp", content, StringComparison.Ordinal);
                if (surface == "cli")
                {
                    Assert.Contains("```powershell", content, StringComparison.Ordinal);
                    Assert.Contains("excelcli -q range", content, StringComparison.Ordinal);
                    Assert.DoesNotContain("range(action:", content, StringComparison.Ordinal);
                }
                else
                {
                    Assert.Contains("```text", content, StringComparison.Ordinal);
                    Assert.Contains("workbook_session_id: sessionId", content, StringComparison.Ordinal);
                    Assert.DoesNotContain("excelcli -q", content, StringComparison.Ordinal);
                }
            }
        }
        finally { Directory.Delete(root, recursive: true); }
    }
}
