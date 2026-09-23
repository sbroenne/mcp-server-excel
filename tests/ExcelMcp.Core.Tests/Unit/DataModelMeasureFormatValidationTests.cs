using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "DataModel")]
[Trait("Speed", "Fast")]
public class DataModelMeasureFormatValidationTests
{
    private readonly DataModelCommands _commands = new();

    [Fact]
    public void CreateMeasure_WithUnknownFormat_RejectsBeforeBatchExecution()
    {
        using var batch = new RejectingBatch();

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.CreateMeasure(
                batch,
                "SalesTable",
                "Invalid Format",
                "SUM(SalesTable[Amount])",
                formatType: "Accounting"));

        Assert.Equal("formatType", exception.ParamName);
        Assert.Equal(
            "Unknown measure format type: 'Accounting'. Valid values: General, Currency, Decimal, Percentage, WholeNumber. (Parameter 'formatType')",
            exception.Message);
        Assert.Equal(0, batch.ExecuteCalls);
    }

    [Fact]
    public void UpdateMeasure_WithUnknownFormat_RejectsBeforeBatchExecution()
    {
        using var batch = new RejectingBatch();

        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.UpdateMeasure(batch, "Existing Measure", formatType: "Scientific"));

        Assert.Equal("formatType", exception.ParamName);
        Assert.Equal(
            "Unknown measure format type: 'Scientific'. Valid values: General, Currency, Decimal, Percentage, WholeNumber. (Parameter 'formatType')",
            exception.Message);
        Assert.Equal(0, batch.ExecuteCalls);
    }

    private sealed class RejectingBatch : IExcelBatch
    {
        public string WorkbookPath => "unused.xlsx";

        public ILogger Logger => NullLogger.Instance;

        public IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; } =
            new Dictionary<string, Excel.Workbook>();

        public bool HasTimedOutOperation => false;

        public int? ExcelProcessId => null;

        public TimeSpan OperationTimeout => TimeSpan.FromSeconds(1);

        public bool IsExcelVisible => false;

        public int ExecuteCalls { get; private set; }

        public void Dispose()
        {
        }

        public void Execute(
            Action<ExcelContext, CancellationToken> operation,
            CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            throw new InvalidOperationException("Batch execution was not expected.");
        }

        public T Execute<T>(
            Func<ExcelContext, CancellationToken, T> operation,
            CancellationToken cancellationToken = default)
        {
            ExecuteCalls++;
            throw new InvalidOperationException("Batch execution was not expected.");
        }

        public Excel.Workbook GetWorkbook(string filePath) =>
            throw new NotSupportedException();

        public bool IsExcelProcessAlive() => false;

        public void Save(CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public void UpdateWorkbookPath(string workbookPath) =>
            throw new NotSupportedException();
    }
}
