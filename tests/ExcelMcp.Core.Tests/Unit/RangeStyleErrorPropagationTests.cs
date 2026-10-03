using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeStyleErrorPropagationTests
{
    [Fact]
    public void GetStyle_LeavesExceptionHandlingToBatchAndService()
    {
        var callbacks = typeof(RangeCommands).GetNestedTypes(BindingFlags.NonPublic)
            .SelectMany(type => type.GetMethods(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
            .Where(method => method.Name.StartsWith("<GetStyle>", StringComparison.Ordinal))
            .ToArray();
        Assert.NotEmpty(callbacks);

        // Inspect the compiled callback: real Excel tests cover mixed/custom styles,
        // while this guard prevents catches from hiding arbitrary native failures.
        foreach (var callback in callbacks)
        {
            var body = callback.GetMethodBody();
            Assert.NotNull(body);
            Assert.DoesNotContain(body.ExceptionHandlingClauses,
                clause => clause.Flags == ExceptionHandlingClauseOptions.Clause
                    || clause.Flags == ExceptionHandlingClauseOptions.Filter);
        }
    }
}
