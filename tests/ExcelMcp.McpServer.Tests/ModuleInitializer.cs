using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLifetimeTestFramework", "Sbroenne.ExcelMcp.McpServer.Tests")]

namespace Sbroenne.ExcelMcp.McpServer.Tests;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Init() => TestRunExcelLifetime.StartForTestHost();
}
