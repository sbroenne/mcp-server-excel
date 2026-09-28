using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework(
    "Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLifetimeTestFramework",
    "Sbroenne.ExcelMcp.Service.Tests")]

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class ModuleInitializer
{
    [ModuleInitializer]
    internal static void Initialize()
    {
        TestRunExcelLifetime.StartForTestHost();
        ExcelBatch.SuppressVisibleDuringOpen = true;
    }
}
