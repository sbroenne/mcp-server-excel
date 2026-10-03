using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRequiredMetadataTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    // Required-property error propagation is an internal contract that Service cannot expose
    // on demand: a healthy ModelTable implements all the requested properties.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RequiredMetadata_UnsupportedNativeProperty_DoesNotInventEmptyData(bool numeric)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                Assert.Equal(sheetName, ComUtilities.SafeGetString(sheet, "Name"));
                var failure = numeric
                    ? Assert.ThrowsAny<Exception>(() => ComUtilities.SafeGetInt(sheet, "RecordCount"))
                    : Assert.ThrowsAny<Exception>(() => ComUtilities.SafeGetString(sheet, "SourceName"));
                Assert.True(failure is COMException or RuntimeBinderException, failure.ToString());
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }
}
