using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "false")]
public sealed class ChartExportFailureTests
{
    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 3)]
    [InlineData(true, 0)]
    [InlineData(true, 3)]
    public void ExportImage_BatchFailureRemovesNewOutput(bool overwrite, int length)
    {
        var output = Path.Combine(Path.GetTempPath(), $"excel-export-failure-{Guid.NewGuid():N}.png");
        using var batch = DispatchProxy.Create<IExcelBatch, ExportFailureBatch>();
        var failureBatch = Assert.IsAssignableFrom<ExportFailureBatch>(batch);
        failureBatch.OutputPath = output;
        failureBatch.Length = length;
        try
        {
            var error = Assert.Throws<IOException>(() =>
                new ChartCommands().ExportImage(batch, "Chart", output, overwrite: overwrite));
            Assert.Same(failureBatch.Failure, error);
            Assert.False(File.Exists(output));
        }
        finally
        {
            File.Delete(output);
        }
    }

    // The batch boundary injects an export failure after a partial write, without Excel.
    public class ExportFailureBatch : DispatchProxy
    {
        public string OutputPath { get; set; } = string.Empty;
        public int Length { get; set; }
        public IOException Failure { get; } = new("Injected chart export failure.");

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name == nameof(IDisposable.Dispose))
                return null;
            if (targetMethod?.Name == nameof(IExcelBatch.Execute))
            {
                File.WriteAllBytes(OutputPath, new byte[Length]);
                throw Failure;
            }
            throw new NotSupportedException($"Unexpected batch member: {targetMethod?.Name}");
        }
    }
}
