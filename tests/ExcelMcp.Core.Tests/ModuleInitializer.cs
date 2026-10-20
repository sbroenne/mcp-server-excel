using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLeftoverTestFramework", "Sbroenne.ExcelMcp.Core.Tests")]

namespace Sbroenne.ExcelMcp.Core.Tests;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Init()
    {
        TestRunExcelLeftoverCheck.Start();
    }
}
