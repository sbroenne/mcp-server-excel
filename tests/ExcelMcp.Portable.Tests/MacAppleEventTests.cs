using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Utilities;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacAppleEventTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeErrorsRetainStatusAndDistinguishRepliesFromTransportFailures(bool isExcelReply)
    {
        var error = Assert.Throws<MacAppleEventException>(() =>
            MacAppleEvents.Check(-50, "setd", isExcelReply));
        Assert.Equal(-50, error.Status);
        Assert.Equal("setd", error.Operation);
        Assert.Equal(isExcelReply, error.IsExcelReply);
        Assert.Equal("Excel Apple Event 'setd' failed with OSStatus -50.", error.Message);
    }

    [Theory]
    [InlineData(false, "Excel rejected worksheet name 'Rejected'. Worksheet 'Sheet1' was not renamed.")]
    [InlineData(true, "Excel rejected worksheet name 'Rejected'. The new worksheet 'Sheet1' remains in the workbook; inspect it and remove it if appropriate.")]
    public async Task NativeNamingErrorPreservesSharedContextAndOriginalReplyAcrossTheChildBoundary(bool isNewSheet, string expectedMessage)
    {
        var message = WorksheetCommandValidation.NamingRejectionMessage("Rejected", "Sheet1", isNewSheet);
        Assert.Equal(expectedMessage, message);
        var nativeError = new MacAppleEventException(-1004, "Excel reply: Workbook structure is protected.", true);
        var namingError = new MacExcelOperationException("ComInterop", message, nativeError,
            remoteExceptionType: nameof(InvalidOperationException), remoteInnerError: nativeError.Message);
        var backend = new MacExcelBackend((_, _, _) => Task.FromResult(
            new MacProcessResult(0, MacAutomationHost.SerializeFailure(namingError), "")));
        object arguments = isNewSheet
            ? new { filePath = "opaque.xlsx", sheetName = "Rejected" }
            : new { filePath = "opaque.xlsx", oldName = "Sheet1", newName = "Rejected" };

        var transported = await Assert.ThrowsAsync<MacExcelOperationException>(() =>
            backend.InvokeAsync(isNewSheet ? "sheet.create" : "sheet.rename", arguments,
                TimeSpan.FromSeconds(10)));
        var response = transported.ToServiceResponse();
        Assert.False(response.Success);
        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Equal(nameof(InvalidOperationException), response.ExceptionType);
        Assert.Equal($"InvalidOperationException: {message}", response.ErrorMessage);
        Assert.Equal(nativeError.Message, response.InnerError);
        Assert.Null(response.HResult);
    }

    [Theory]
    [InlineData("ComInterop")]
    [InlineData("HelperMissing")]
    public async Task UntypedChildFailuresPreserveExistingServiceDiagnostics(string category)
    {
        const string message = "Native operation failed; no mutation was confirmed.";
        var backend = new MacExcelBackend((_, _, _) => Task.FromResult(
            new MacProcessResult(0,
                MacAutomationHost.SerializeFailure(new MacExcelOperationException(category, message)), "")));
        var error = await Assert.ThrowsAsync<MacExcelOperationException>(() =>
            backend.InvokeAsync("sheet.rename", new { filePath = "opaque.xlsx" }, TimeSpan.FromSeconds(10)));
        var response = error.ToServiceResponse();
        Assert.False(response.Success);
        Assert.Equal(category, response.ErrorCategory);
        Assert.Equal(message, response.ErrorMessage);
        Assert.Equal(nameof(MacExcelOperationException), response.ExceptionType);
        Assert.Null(response.InnerError);
        Assert.Null(response.HResult);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeTimeoutRemainsTimeoutAcrossTheChildBoundaryForBothErrorOrigins(bool isExcelReply)
    {
        var timeout = Assert.Throws<TimeoutException>(() =>
            MacAppleEvents.Check(-1712, "setd", isExcelReply));
        var backend = new MacExcelBackend((_, _, _) => Task.FromResult(
            new MacProcessResult(0, MacAutomationHost.SerializeFailure(timeout), "")));
        var transported = await Assert.ThrowsAsync<TimeoutException>(() =>
            backend.InvokeAsync("sheet.rename", new { filePath = "opaque.xlsx" }, TimeSpan.FromSeconds(10)));
        Assert.Equal(timeout.Message, transported.Message);
        Assert.Contains("uncertain", transported.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("PlatformNotSupported")]
    [InlineData("Timeout")]
    [InlineData("ComInterop")]
    [InlineData("HelperMissing")]
    [InlineData("HelperIncompatible")]
    public void ChildFailurePreservesCategoryAndMessage(string expectedCategory)
    {
        const string message = "Native operation failed; no mutation was confirmed.";
        Exception error = expectedCategory switch
        {
            "PlatformNotSupported" => new PlatformNotSupportedException(message),
            "Timeout" => new TimeoutException(message),
            "ComInterop" => new InvalidOperationException(message),
            _ => new MacExcelOperationException(expectedCategory, message)
        };

        using var response = JsonDocument.Parse(MacAutomationHost.SerializeFailure(error));
        Assert.False(response.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(expectedCategory, response.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(message, response.RootElement.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public void DescriptorUsesTheNativeTwoBytePacking()
    {
        Assert.Equal(4, Marshal.OffsetOf<MacAutomationAccess.Descriptor>("DataHandle").ToInt32());
        Assert.Equal(4 + IntPtr.Size, Marshal.SizeOf<MacAutomationAccess.Descriptor>());
    }

    [Fact]
    public void CodesUseBigEndianCharacterOrder()
    {
        Assert.Equal(0x636F7265u, MacAppleEvents.Code("core"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("three")]
    [InlineData("a\u00e9cd")]
    public void InvalidCodesFailExplicitly(string code)
    {
        Assert.Throws<ArgumentException>(() => MacAppleEvents.Code(code));
    }

    [Theory]
    [InlineData(0.001, 1)]
    [InlineData(1, 60)]
    [InlineData(1.01, 61)]
    public void TimeoutRoundsUpToPositiveNativeTicks(double seconds, int expected)
    {
        Assert.Equal((nint)expected, MacAppleEvents.TimeoutTicks(TimeSpan.FromSeconds(seconds)));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void UnboundedOrNonpositiveTimeoutIsRejected(double seconds)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            MacAppleEvents.TimeoutTicks(TimeSpan.FromSeconds(seconds)));
    }

    [Fact]
    public void EnumDescriptorsPreserveTheirNativeCodeWithoutGuessingItsMeaning()
    {
        Assert.Equal(0x0270FFFFu, MacAppleEvents.DecodeScalar(
            MacAppleEvents.Code("enum"), BitConverter.GetBytes(0x0270FFFFu))!.GetValue<uint>());
    }

    [Fact]
    public void ScalarsPreserveTheirNativeTypes()
    {
        Assert.Equal(-42, MacAppleEvents.DecodeScalar(MacAppleEvents.Code("long"), BitConverter.GetBytes(-42))!.GetValue<int>());
        Assert.Equal(1.25, MacAppleEvents.DecodeScalar(MacAppleEvents.Code("doub"), BitConverter.GetBytes(1.25))!.GetValue<double>());
        Assert.True(MacAppleEvents.DecodeScalar(MacAppleEvents.Code("bool"), [1])!.GetValue<bool>());
        Assert.Equal("Data", MacAppleEvents.DecodeScalar(MacAppleEvents.Code("utxt"), Encoding.Unicode.GetBytes("Data"))!.GetValue<string>());
    }

    [Fact]
    public void TypedMissingValueIsDecodedAsMissingNotAsAnOrdinaryNumber()
    {
        Assert.Null(MacAppleEvents.DecodeScalar(
            MacAppleEvents.Code("type"), BitConverter.GetBytes(MacAppleEvents.Code("msng"))));
    }

    [Theory]
    [InlineData("bool", new byte[] { 2 })]
    [InlineData("long", new byte[] { 1 })]
    [InlineData("utxt", new byte[] { 1 })]
    [InlineData("TEXT", new byte[] { 1 })]
    [InlineData("enum", new byte[] { 1 })]
    public void UnknownOrMalformedScalarsNeverBecomeSuccess(string type, byte[] data)
    {
        Assert.Throws<InvalidOperationException>(() => MacAppleEvents.DecodeScalar(MacAppleEvents.Code(type), data));
    }

    [Fact]
    public void ChildTimeout_PreservesTheExactPositiveBudget()
    {
        Assert.Equal(TimeSpan.FromTicks(12345), MacAutomationHost.ReadTimeout("12345"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("-1")]
    [InlineData("NaN")]
    [InlineData("10000000000000000000")]
    public void ChildTimeout_RejectsMalformedOrNonpositiveBudgets(string ticks)
    {
        Assert.Throws<InvalidOperationException>(() => MacAutomationHost.ReadTimeout(ticks));
    }

    [Fact]
    public void NativeTimeoutRetainsUncertainOutcome()
    {
        var error = Assert.Throws<TimeoutException>(() => MacAppleEvents.Check(-1712, "setd"));
        Assert.Contains("uncertain", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void NativeDescriptorsHaveDeterministicIdempotentLifetime()
    {
        if (!OperatingSystem.IsMacOS())
        {
            return;
        }
        var descriptor = MacAppleEvents.Text("descriptor test");
        Assert.Equal("descriptor test", descriptor.Decode()!.GetValue<string>());
        descriptor.Dispose();
        descriptor.Dispose();
        Assert.Throws<ObjectDisposedException>(() => descriptor.Decode());
    }

    [Fact]
    public void NativeListsRetainCopiesAfterTheirSourceDescriptorsAreDisposed()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Assert.Throws<PlatformNotSupportedException>(() => MacAppleEvents.List());
            return;
        }
        using var matrix = MacAppleEvents.List();
        using (var row = MacAppleEvents.List())
        {
            using (var formula = MacAppleEvents.Text("=SEQUENCE(3)"))
            {
                MacAppleEvents.Append(row, formula);
            }
            MacAppleEvents.Append(matrix, row);
        }
        Assert.True(System.Text.Json.Nodes.JsonNode.DeepEquals(
            System.Text.Json.Nodes.JsonNode.Parse("[[\"=SEQUENCE(3)\"]]"), matrix.Decode()));
        matrix.Dispose();
        using var value = MacAppleEvents.Text("=1");
        Assert.Throws<ObjectDisposedException>(() => MacAppleEvents.Append(matrix, value));
    }

    [Fact]
    public void NativeListAppendRejectsRecordsInsteadOfChangingTheirKeys()
    {
        if (!OperatingSystem.IsMacOS())
        {
            Assert.Throws<PlatformNotSupportedException>(() => MacAppleEvents.List());
            return;
        }
        using var record = MacAppleEvents.Record(MacAppleEvents.Code("reco"));
        using var value = MacAppleEvents.Text("=1");
        var error = Assert.Throws<ArgumentException>(() => MacAppleEvents.Append(record, value));
        Assert.Equal("list", error.ParamName);
    }
}
