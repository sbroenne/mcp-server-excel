using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "Workbook")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class DocumentPropertyReadFailureTests
{
    private static DocumentPropertyInfo ReadProperty(NativeProperty property, DocumentPropertyScope scope) =>
        WorkbookCommands.ReadDocumentProperty(property.ReadName, property.ReadType, property.ReadValue, scope);

    [Theory]
    [InlineData(DocumentPropertyScope.BuiltIn, unchecked((int)0x800A03EC))]
    [InlineData(DocumentPropertyScope.BuiltIn, unchecked((int)0x80070005))]
    [InlineData(DocumentPropertyScope.BuiltIn, unchecked((int)0x800706BE))]
    [InlineData(DocumentPropertyScope.BuiltIn, unchecked((int)0x80010108))]
    [InlineData(DocumentPropertyScope.Custom, unchecked((int)0x80004005))]
    public void Value_UnexpectedNativeFailure_PropagatesOriginalException(DocumentPropertyScope scope, int hresult)
    {
        var error = NativeFailure(hresult);
        var property = new NativeProperty { ReadValue = () => throw error };

        var thrown = Assert.Throws<COMException>(() => ReadProperty(property, scope));

        Assert.Same(error, thrown);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void IdentityGetter_Failure_IsNotMisclassifiedAsUnreadableValue(bool nameFails)
    {
        var error = NativeFailure(unchecked((int)0x80004005));
        var property = new NativeProperty
        {
            ReadName = () => nameFails ? throw error : "Last print date",
            ReadType = () => nameFails ? 3 : throw error
        };

        var thrown = Assert.Throws<COMException>(() => ReadProperty(property, DocumentPropertyScope.BuiltIn));

        Assert.Same(error, thrown);
    }

    [Fact]
    public void BuiltInValue_EFail_ReportsNativeFailureWithoutClaimingUndefined()
    {
        var error = NativeFailure(unchecked((int)0x80004005));
        var property = new NativeProperty { ReadValue = () => throw error };

        var result = ReadProperty(property, DocumentPropertyScope.BuiltIn);

        Assert.Equal("Last print date", result.Name);
        Assert.Equal("built-in", result.Scope);
        Assert.Equal("date", result.ValueType);
        Assert.False(result.ValueAvailable);
        Assert.Null(result.Value);
        Assert.Contains("0x80004005", result.ValueError, StringComparison.Ordinal);
        Assert.Contains(error.Message, result.ValueError, StringComparison.Ordinal);
        Assert.Contains("may be undefined", result.ValueError, StringComparison.Ordinal);
        Assert.Contains("may have failed", result.ValueError, StringComparison.Ordinal);
    }

    [Fact]
    public void KnownUnavailableMember_ReportsErrorInsteadOfSuccessfulEmptyValue()
    {
        var error = NativeFailure(unchecked((int)0x80020003));
        var property = new NativeProperty { ReadValue = () => throw error };

        var result = ReadProperty(property, DocumentPropertyScope.Custom);

        Assert.False(result.ValueAvailable);
        Assert.Null(result.Value);
        Assert.Equal("custom", result.Scope);
        Assert.Contains("0x80020003", result.ValueError, StringComparison.Ordinal);
        Assert.Contains(error.Message, result.ValueError, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Filled title")]
    public void SuccessfulValue_RetainsExactValueAndHasNoError(string value)
    {
        var property = new NativeProperty
        {
            ReadName = () => "Title",
            ReadType = () => 4,
            ReadValue = () => value
        };

        var result = ReadProperty(property, DocumentPropertyScope.BuiltIn);

        Assert.True(result.ValueAvailable);
        Assert.Equal(value, result.Value);
        Assert.Null(result.ValueError);
        Assert.Equal("Title", result.Name);
        Assert.Equal("string", result.ValueType);
        Assert.Equal("built-in", result.Scope);
    }

#pragma warning disable CA2201 // Inject exact native property HRESULTs without starting Excel.
    private static COMException NativeFailure(int hresult) => new("Native Value getter failed.", hresult);
#pragma warning restore CA2201

    private sealed class NativeProperty
    {
        public Func<string> ReadName { get; init; } = () => "Last print date";
        public Func<int> ReadType { get; init; } = () => 3;
        public Func<object?> ReadValue { get; init; } = () => null;
    }
}
