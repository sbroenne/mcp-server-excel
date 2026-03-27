using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Tests;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Init()
    {
        // Suppress "start visible during open" in tests to avoid flashing Excel windows.
        // Production uses Visible=true during workbook open so enterprise auth/sign-in
        // dialogs are interactable (PR #577). Tests don't need this behavior.
        ExcelBatch.SuppressVisibleDuringOpen = true;
    }
}
