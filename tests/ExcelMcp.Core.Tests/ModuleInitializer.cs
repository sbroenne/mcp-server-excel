using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Shared;

[assembly: Xunit.TestFramework("Sbroenne.ExcelMcp.Tests.Infrastructure.ExcelLifetimeTestFramework", "Sbroenne.ExcelMcp.Core.Tests")]

namespace Sbroenne.ExcelMcp.Core.Tests;

internal static class ModuleInit
{
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
