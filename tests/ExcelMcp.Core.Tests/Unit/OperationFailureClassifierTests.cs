using System.Reflection;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Feature", "ErrorHandling")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class OperationFailureClassifierTests
{
    [Fact]
    public void Classify_TypedOuterFailure_TakesPrecedenceOverComCause()
    {
#pragma warning disable CA2201 // Synthetic exception for pure classification.
        var com = new COMException("Private Excel detail", unchecked((int)0x800A03EC));
#pragma warning restore CA2201
        var error = new OperationFailureException(OperationFailureCategory.Permissions, "Access blocked", com);
        Assert.Equal("Permissions", OperationFailureClassifier.Classify(new TargetInvocationException(error)));
        Assert.Same(com, error.InnerException);
        Assert.Equal("0x800A03EC", OperationFailureClassifier.GetComHResult(new TargetInvocationException(error)));
        Assert.Null(OperationFailureClassifier.GetComHResult(new AggregateException(com, error)));
        Assert.Equal("0x800A03EC", OperationFailureClassifier.GetComHResult(new AggregateException(error)));
    }

    [Fact]
    public void Classify_UnknownMessage_DoesNotGuessCategory()
    {
        Assert.Null(OperationFailureClassifier.Classify(
            new InvalidOperationException("Timeout syntax permission privacy COM error")));
    }

    [Fact]
    public void Classify_MixedAggregate_DoesNotSelectFirstFailure()
    {
        Assert.Null(OperationFailureClassifier.Classify(new AggregateException(
            new ArgumentException("Bad argument"), new TimeoutException("Late"))));
        Assert.Null(OperationFailureClassifier.Classify(new AggregateException(
            new ArgumentException("Bad argument"), new InvalidOperationException("Unknown"))));
        Assert.Null(OperationFailureClassifier.Classify(new AggregateException()));
        Assert.Equal("InvalidInput", OperationFailureClassifier.Classify(new AggregateException(
            new ArgumentException("First"), new ArgumentException("Second"))));
    }

    [Fact]
    public void Classify_QueryCleanup_PreservesExplicitCategoryOverMixedCauses()
    {
        var error = new PowerQueryCommandException("Cleanup failed", "Cleanup",
            new AggregateException(new ArgumentException("Query"), new InvalidOperationException("Cleanup")));
        Assert.Equal("Cleanup", OperationFailureClassifier.Classify(error));
    }
}
