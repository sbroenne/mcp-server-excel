using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure.Session;

internal interface ISessionService
{
    string Create(string filePath);
    bool Close(string sessionId, bool save = false);
    IReadOnlyList<SessionDescriptor> List();
    IExcelBatch GetBatch(string sessionId);
}
