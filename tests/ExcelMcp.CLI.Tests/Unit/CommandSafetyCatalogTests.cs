using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "Safety")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CommandSafetyCatalogTests
{
    [Fact]
    public void GeneratedCatalog_ClassifiesEveryPublicServiceAction()
    {
        var expectedCommands = typeof(ServiceCategoryAttribute).Assembly
            .GetTypes()
            .Where(type => type.IsInterface)
            .Select(type => new
            {
                Type = type,
                Category = type.GetCustomAttributes(typeof(ServiceCategoryAttribute), inherit: false)
                    .Cast<ServiceCategoryAttribute>()
                    .SingleOrDefault()?.Category
            })
            .Where(item => item.Category is not null)
            .SelectMany(item => item.Type.GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.DeclaredOnly).Select(method => new
            {
                Command = $"{item.Category}.{method.GetCustomAttributes(typeof(ServiceActionAttribute), inherit: false).Cast<ServiceActionAttribute>().SingleOrDefault()?.Action ?? ToKebabCase(method.Name)}"
            }))
            .Select(item => item.Command)
            .OrderBy(command => command, StringComparer.Ordinal)
            .ToArray();

        Assert.NotEmpty(expectedCommands);
        Assert.Equal(expectedCommands.Length, ServiceRegistry.SafetyDescriptors.Count);
        Assert.All(expectedCommands, command =>
        {
            var descriptor = ServiceRegistry.GetSafetyDescriptor(command);
            Assert.True(descriptor.ExplicitlyClassified, $"{command} was not explicitly classified.");
        });
    }

    [Fact]
    public void RangeActions_DistinguishReadFromValueMutation()
    {
        var read = ServiceRegistry.GetSafetyDescriptor("range.get-values");
        var write = ServiceRegistry.GetSafetyDescriptor("range.set-values");

        Assert.False(read.IsMutation);
        Assert.True(write.IsMutation);
        Assert.Equal("values", write.MutationKind);
        Assert.Equal("rangeSemantic", write.VerificationLevel);
    }

    [Fact]
    public void RangeFormatMutation_UsesPartialVerification()
    {
        var format = ServiceRegistry.GetSafetyDescriptor("rangeformat.set-style");

        Assert.True(format.IsMutation);
        Assert.Equal("formatting", format.MutationKind);
        Assert.Equal("rangeFingerprint", format.VerificationLevel);
    }

    [Theory]
    [InlineData("diag.ping")]
    [InlineData("diag.echo")]
    [InlineData("diag.validate-params")]
    public void DiagnosticTransportActions_AreReadOnly(string command)
    {
        var descriptor = ServiceRegistry.GetSafetyDescriptor(command);

        Assert.False(descriptor.IsMutation);
        Assert.Equal("none", descriptor.MutationKind);
        Assert.Equal("none", descriptor.VerificationLevel);
        Assert.Equal("none", descriptor.RecoveryRisk);
    }

    [Fact]
    public void UnknownAction_FailsClosedAsMutation()
    {
        var descriptor = ServiceRegistry.GetSafetyDescriptor("future.unseen-action");

        Assert.True(descriptor.IsMutation);
        Assert.False(descriptor.ExplicitlyClassified);
        Assert.Equal("unknown", descriptor.MutationKind);
    }

    private static string ToKebabCase(string value)
    {
        var characters = new List<char>(value.Length + 8);
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (index > 0 && char.IsUpper(character))
            {
                characters.Add('-');
            }

            characters.Add(char.ToLowerInvariant(character));
        }

        return new string([.. characters]);
    }
}
