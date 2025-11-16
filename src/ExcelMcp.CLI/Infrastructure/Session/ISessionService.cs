using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure.Session;

internal interface ISessionService
{
    string Create(string filePath);
    bool Save(string sessionId);
    bool Close(string sessionId);
    IReadOnlyList<SessionDescriptor> List();
    IExcelBatch GetBatch(string sessionId);
}
