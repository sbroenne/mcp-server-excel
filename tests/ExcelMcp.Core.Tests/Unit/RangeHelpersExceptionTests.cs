using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

public class RangeHelpersExceptionTests
{
    [Fact]
    public void ResolveRange_DoesNotCatchBaseException()
    {
        var method = typeof(RangeHelpers).GetMethod(
            nameof(RangeHelpers.ResolveRange),
            [typeof(object), typeof(string), typeof(string), typeof(string).MakeByRefType()]);

        Assert.NotNull(method);

        var catchTypes = method.GetMethodBody()!
            .ExceptionHandlingClauses
            .Where(clause => clause.Flags == ExceptionHandlingClauseOptions.Clause)
            .Select(clause => clause.CatchType);

        Assert.DoesNotContain(typeof(Exception), catchTypes);
    }
}
