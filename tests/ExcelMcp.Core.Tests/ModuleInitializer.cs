using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLifetimeTestFramework", "Sbroenne.ExcelMcp.Core.Tests")]

namespace Sbroenne.ExcelMcp.Core.Tests;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Init()
    {
        TestRunExcelLifetime.StartForTestHost();
    }
}
