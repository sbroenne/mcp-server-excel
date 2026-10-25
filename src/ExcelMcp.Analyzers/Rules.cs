using Microsoft.CodeAnalysis;

namespace Sbroenne.ExcelMcp.Analyzers;

/// <summary>
/// Repository rules enforced at compile time. Each message says what to do instead,
/// because the reader is usually a coding agent fixing the build.
/// </summary>
internal static class Rules
{
    private const string ErrorHandling = "ErrorHandling";
    private const string ComSafety = "ComSafety";
    private const string Maintainability = "Maintainability";

    public static readonly DiagnosticDescriptor EmptyCatch = new(
        "EXL001",
        "Empty catch block swallows errors",
        "Empty catch block hides failures. Let the exception propagate, handle it with real code, or filter the exact expected error with 'when (...)'.",
        ErrorHandling, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor BroadCatch = new(
        "EXL002",
        "Broad catch must rethrow or filter a specific error",
        "catch ({0}) neither rethrows nor filters a specific error. Propagate the failure or filter a named, specific HResult. Deliberate failure-transport boundaries need a local SuppressMessage with a reason explaining how the failure is reported.",
        ErrorHandling, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnjustifiedDynamicCast = new(
        "EXL005",
        "Casting to dynamic is not allowed",
        "Casting to dynamic bypasses compile-time API checks. Use the typed member; comments do not exempt late binding.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnjustifiedSuppression = new(
        "EXL006",
        "Suppression needs a reason",
        "{0} must say why the warning does not apply here. Prefer fixing the code; otherwise add the reason ({1}).",
        Maintainability, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DirectDynamicCast = new(
        "EXL010",
        "Use Convert for values read from dynamic COM",
        "Excel may return this value as a different numeric type (often double). Use Convert.To{0}(value, CultureInfo.InvariantCulture) instead of a direct cast to {1}.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DynamicAccess = new(
        "EXL011",
        "Dynamic access is not allowed",
        "Dynamic declarations and late-bound calls bypass compile-time API checks. Use typed interop; do not bypass a missing PIA member with dynamic access.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor ChainedComAccess = new(
        "EXL004",
        "Store acquired COM references before using them",
        "'{0}' can acquire an intermediate COM object. Store it in a local and release it in finally; session-owned App and Book are excluded.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor DiscardedComResult = new(
        "EXL007",
        "Do not discard an acquired COM reference",
        "'{0}' returns a COM object. Store the result and release it in finally.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor WorkbookPackageAccess = new(
        "EXL009",
        "Workbook contents must be accessed through Excel",
        "'{0}' accesses an Excel workbook package. Use Excel COM in production; package access belongs only in test fixtures.",
        ComSafety, DiagnosticSeverity.Error, isEnabledByDefault: true);
}
