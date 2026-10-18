namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Macro workbook that tests save and reopen from an untrusted temporary folder,
/// so Excel's own macro settings block its macros after the first reopen.
/// </summary>
public sealed class PersistentServiceReopenedVbaFixture :
    PersistentServiceWorkbookFixture
{
    public PersistentServiceReopenedVbaFixture() :
        base("PersistentServiceReopenedVba.xlsm")
    {
    }
}
