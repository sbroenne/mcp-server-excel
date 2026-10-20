using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework(
    "Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLeftoverTestFramework",
    "Sbroenne.ExcelMcp.Service.Tests")]

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class ModuleInitializer
{
    [ModuleInitializer]
    internal static void Initialize()
    {
        TestRunExcelLeftoverCheck.Start();
    }
}
