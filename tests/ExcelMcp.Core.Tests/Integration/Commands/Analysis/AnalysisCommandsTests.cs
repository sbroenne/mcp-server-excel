using Sbroenne.ExcelMcp.Core.Commands.Analysis;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Analysis;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Analysis")]
[Trait("RequiresExcel", "false")]
public sealed class AnalysisCommandsTests
{
    [Fact]
    public void GoalSeek_NullGoal_RejectsBeforeOpeningExcel()
    {
        var commands = new AnalysisCommands();

        var exception = Assert.Throws<ArgumentNullException>(
            () => commands.GoalSeek(null!, "Sheet1", "B1", null, "A1"));

        Assert.Equal("goal", exception.ParamName);
    }
}
