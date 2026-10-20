using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLeftoverTestFramework", "Sbroenne.ExcelMcp.ComInterop.Tests")]

namespace Sbroenne.ExcelMcp.ComInterop.Tests;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Init()
    {
        TestRunExcelLeftoverCheck.Start();
    }
}
