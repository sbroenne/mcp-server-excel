using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
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
        // Suppress "start visible during open" in tests to avoid flashing Excel windows.
        // Production uses Visible=true during workbook open so enterprise auth/sign-in
        // dialogs are interactable (PR #577). Tests don't need this behavior.
        ExcelBatch.SuppressVisibleDuringOpen = true;
    }
}
