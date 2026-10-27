using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

internal sealed class DialogWindowControlStub(string details) : IExcelBatchWindowControl
{
    public int DetailReads { get; private set; }
    public string? DialogDetails
    {
        get
        {
            DetailReads++;
            return details;
        }
    }

    public string? DialogVisibilityError => null;
    public ExcelVisibilityResult SetVisibility(bool visible) =>
        throw new InvalidOperationException("Recovery guidance must not change visibility directly.");
}
