using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaProcedureSourceTests
{
    [Fact]
    public void ParseSingleProcedure_MultilineFunctionSignature_ReturnsNameAndKind()
    {
        const string source = """
            Public Function BuildLabel( _
                ByVal firstName As String, _
                ByVal lastName As String) As String
                BuildLabel = firstName & " " & lastName
            End Function
            """;

        var result = VbaProcedureSource.ParseSingleProcedure(source);

        Assert.Equal("BuildLabel", result.Name);
        Assert.Equal("Function", result.Kind);
    }

    [Theory]
    [InlineData("Property Get", "Property Get")]
    [InlineData("Property Let", "Property Let")]
    [InlineData("Property Set", "Property Set")]
    public void ParseSingleProcedure_PropertyAccessor_ReturnsAccessorKind(
        string declaration,
        string expectedKind)
    {
        string source = $"{declaration} Value() As Variant{Environment.NewLine}End Property";

        var result = VbaProcedureSource.ParseSingleProcedure(source);

        Assert.Equal("Value", result.Name);
        Assert.Equal(expectedKind, result.Kind);
    }

    [Theory]
    [InlineData("Sub First()\nEnd Sub\nSub Second()\nEnd Sub")]
    [InlineData("Sub First()\nEnd Function")]
    [InlineData("Option Explicit\nSub First()\nEnd Sub")]
    [InlineData("Public Sub KeepMe()\nEnd Sub: Public Sub Extra()\nEnd Sub")]
    [InlineData("Public Sub KeepMe(): End Sub: Public Sub Extra()\nEnd Sub")]
    [InlineData("Public Sub KeepMe()\nDebug.Print \"text\": End Sub: Public Sub Extra()\nEnd Sub")]
    public void ParseSingleProcedure_AnythingBeyondOneProcedure_Throws(
        string source)
    {
        Assert.Throws<ArgumentException>(
            () => VbaProcedureSource.ParseSingleProcedure(source));
    }

    [Fact]
    public void ParseSingleProcedure_ColonsInStatementsStringsAndComments_AcceptsOneProcedure()
    {
        const string source = """
            Public Sub KeepMe()
                Debug.Print "End Sub: Public Sub Extra()": Debug.Print "quoted "" : "" text"
                ' End Sub: Public Sub Extra()
                Debug.Print "ok": Rem End Sub: Public Sub Extra()
            End Sub
            """;

        var result = VbaProcedureSource.ParseSingleProcedure(source);
        Assert.Equal("KeepMe", result.Name);
        Assert.Equal("Sub", result.Kind);
    }

    [Fact]
    public void GetBodyLineCount_ExcludesTrailingCommentsAndBlankLines()
    {
        Assert.Equal(3, VbaProcedureSource.GetBodyLineCount(
            "Public Sub Target()\n    Debug.Print \"End Sub\"\nEnd Sub ' closing comment\n\n' trailing note\n",
            "Sub"));
        Assert.Throws<ArgumentException>(() => VbaProcedureSource.GetBodyLineCount(
            "Public Sub Target()\nEnd Function", "Sub"));
    }

    [Fact]
    public void ComputeHash_DifferentLineEndings_ProducesSameFingerprint()
    {
        const string source = "Sub First()\r\nEnd Sub";

        Assert.Equal(
            VbaProcedureSource.ComputeHash(source),
            VbaProcedureSource.ComputeHash(source.Replace("\r\n", "\n", StringComparison.Ordinal)));
    }
}
