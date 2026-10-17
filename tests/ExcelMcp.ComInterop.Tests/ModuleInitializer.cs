using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLifetimeTestFramework", "Sbroenne.ExcelMcp.ComInterop.Tests")]

namespace Sbroenne.ExcelMcp.ComInterop.Tests;

internal static class ModuleInit
{
    internal static TestRunExcelLifetime? Lifetime => TestRunExcelLifetime.CurrentHost;

    [ModuleInitializer]
    internal static void Init()
    {
        TestRunExcelLifetime.StartForTestHost();
    }
}
