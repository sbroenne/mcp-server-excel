using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

// CA2201: COMException is "reserved by the runtime", but these regression tests must
// fabricate a COMException carrying a precise HRESULT (0x800A03EC) and message to
// deterministically reproduce the COM failure classified by VbaCommands.IsVbaTrustError.
// There is no other way to exercise that classification logic without a live Excel COM
// server raising the (environment-specific, non-reproducible) error. Test-only.
#pragma warning disable CA2201

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

/// <summary>
/// Regression tests for issue #671: <c>vba(import)</c> (and the other VBA lifecycle
/// commands) misreported generic COM failures as "VBA trust access is not enabled".
///
/// HRESULT 0x800A03EC is the GENERIC Office automation error ("Exception occurred"),
/// reused for many unrelated failures. The previous catch filter relabeled EVERY
/// 0x800A03EC as a trust error, which masked real failures and sent users to re-check
/// Trust Center settings that were already correct - on a Japanese Office build in the
/// report, where the COM message text differs from the English "programmatic access".
///
/// These tests must FAIL against the old "every 0x800A03EC is a trust error" logic and
/// PASS after the fix.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaTrustErrorClassificationTests
{
    // DISP_E_EXCEPTION - the generic Office automation error.
    private const int GenericOfficeAutomationError = unchecked((int)0x800A03EC);

    [Fact]
    public void IsVbaTrustError_GenericAutomationError_ClassifiedByActualTrustState()
    {
        // A 0x800A03EC carrying a non-trust message must only be treated as a trust error
        // when trust is ACTUALLY disabled in the registry (locale-independent). When trust
        // is enabled, the real error must be surfaced - this is the core of issue #671.
        var genericComError = new COMException("Some unrelated automation failure", GenericOfficeAutomationError);

        bool expected = !VbaCommands.IsVbaTrustEnabled();

        Assert.Equal(expected, VbaCommands.IsVbaTrustError(genericComError));
    }

    [Fact]
    public void IsVbaTrustError_ProgrammaticAccessMessage_AlwaysTrustError()
    {
        // The genuine English trust message is always recognized, regardless of registry
        // state, preserving the helpful guidance when trust really is the problem.
        var trustError = new COMException(
            "Programmatic access to Visual Basic Project is not trusted",
            GenericOfficeAutomationError);

        Assert.True(VbaCommands.IsVbaTrustError(trustError));
    }

    [Fact]
    public void IsVbaTrustError_SpecificNonGenericHResult_IsNotTrustError()
    {
        // A specific (non-generic) HRESULT is never the trust error and must be surfaced
        // unchanged so callers see the real failure.
        var otherError = new COMException("A specific object error", unchecked((int)0x800AC3D4));

        Assert.False(VbaCommands.IsVbaTrustError(otherError));
    }

    [Fact]
    public void BuildGenericComErrorMessage_IncludesHResultOriginalMessageAndDiagnostics()
    {
        // For genuine (non-trust) 0x800A03EC failures the surfaced message must carry enough
        // context to triage the still-unknown underlying cause of issue #671: the HRESULT, the
        // original COM description, and the COM environment diagnostics.
        var comEx = new COMException(
            "\u30d7\u30ed\u30b0\u30e9\u30df\u30f3\u30b0\u306b\u3088\u308b Visual Basic \u30d7\u30ed\u30b8\u30a7\u30af\u30c8",
            VbaCommands.GenericOfficeAutomationError);

        string message = VbaCommands.BuildGenericComErrorMessage(comEx);

        Assert.Contains("0x800A03EC", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(comEx.Message, message, StringComparison.Ordinal);
        Assert.Contains("COM Diagnostics:", message, StringComparison.Ordinal);
        // It must NOT reassert the (incorrect) trust diagnosis.
        Assert.DoesNotContain("trust access is not enabled", message, StringComparison.OrdinalIgnoreCase);
    }
}
#pragma warning restore CA2201
