using System.Reflection;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Layer", "Core")]
[Trait("Feature", "ConditionalFormat")]
[Trait("Speed", "Fast")]
public sealed class ConditionalFormattingVariantTests
{
    private static readonly Func<object?, int?> ReadInt =
        typeof(ConditionalFormattingCommands).GetMethod("ReadOptionalFormattingInt", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<object?, int?>>();
    private static readonly Func<object?, bool?> ReadBool =
        typeof(ConditionalFormattingCommands).GetMethod("ReadOptionalFormattingBool", BindingFlags.Static | BindingFlags.NonPublic)!
            .CreateDelegate<Func<object?, bool?>>();

    [Fact]
    public void UnsetVariants_AreUnknown_NotZeroOrFalse()
    {
        Assert.Null(ReadInt(null));
        Assert.Null(ReadInt(DBNull.Value));
        Assert.Null(ReadBool(null));
        Assert.Null(ReadBool(DBNull.Value));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-4142)]
    [InlineData(65535)]
    public void SetNumericVariant_IsRetained(int value) => Assert.Equal(value, ReadInt(value));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SetBooleanVariant_IsRetained(bool value) => Assert.Equal(value, ReadBool(value));

    [Fact]
    public void UnexpectedVariant_IsNotConvertedToAbsence()
    {
        Assert.Throws<FormatException>(() => ReadInt("invalid native integer"));
        Assert.Throws<FormatException>(() => ReadBool("invalid native boolean"));
    }
}
