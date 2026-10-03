using System.Collections;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public partial class PersistentServiceDataModelCommandsTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ExplicitMeasureFormat_FactoryFailure_DoesNotReportSuccessOrChangeMeasure(
        bool updating, bool bindingFailure)
    {
        var batch = _fixture.BatchToken;
        var name = $"FormatFailure_{Guid.NewGuid():N}";
        const string originalFormula = "SUM(SalesTable[Amount])";
        const string sentinel = "Synthetic requested format factory failure";
        if (updating)
        {
            Assert.True(CreateMeasure(
                "SalesTable", name, originalFormula, "Currency", "Original description").Success);
        }

        // Inject a failure in our .NET format-factory map, not a mocked Excel object.
        // Service routing, the workbook, measure writes and readback all use real Excel.
        var field = typeof(DataModelCommands).GetField(
            "MeasureFormatFactories", BindingFlags.Static | BindingFlags.NonPublic);
        Assert.NotNull(field);
        var factories = Assert.IsAssignableFrom<IDictionary>(field.GetValue(null));
        var originalFactory = Assert.IsAssignableFrom<Delegate>(factories["Percentage"]);
        Exception failure = bindingFailure
            ? new RuntimeBinderException(sentinel)
            : Assert.IsType<COMException>(
                Marshal.GetExceptionForHR(unchecked((int)0x800A03EC), new IntPtr(-1)));
        try
        {
            var parameterType = originalFactory.Method.GetParameters()[0].ParameterType;
            var parameter = Expression.Parameter(parameterType, "model");
            factories["Percentage"] = Expression.Lambda(
                originalFactory.GetType(),
                Expression.Throw(Expression.Constant(failure), typeof(object)),
                parameter).Compile();
            var exception = Assert.ThrowsAny<Exception>(() =>
            {
                if (updating)
                {
                    _dataModelCommands.UpdateMeasure(
                        batch, name, "AVERAGE(SalesTable[Amount])", "Percentage", "Replacement description");
                }
                else
                {
                    CreateMeasure("SalesTable", name, originalFormula, "Percentage");
                }
            });
            Assert.Contains(failure.Message, exception.Message);
        }
        finally
        {
            factories["Percentage"] = originalFactory;
        }

        if (updating)
        {
            var read = _dataModelCommands.Read(batch, name);
            Assert.True(read.Success, read.ErrorMessage);
            Assert.Equal(originalFormula, read.DaxFormula);
            Assert.Equal("Original description", read.Description);
            Assert.NotNull(read.FormatInfo);
            Assert.Equal("Currency", read.FormatInfo.Type);
        }
        else
        {
            var listed = _dataModelCommands.ListMeasures(batch);
            Assert.True(listed.Success, listed.ErrorMessage);
            Assert.DoesNotContain(listed.Measures, measure => measure.Name == name);
        }
    }
}
