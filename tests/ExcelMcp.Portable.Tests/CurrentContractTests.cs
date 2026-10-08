using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Service.Mac;
using Sbroenne.ExcelMcp.Service;
using ServiceCategoryAttribute = Sbroenne.ExcelMcp.Core.Attributes.ServiceCategoryAttribute;
using NoSessionAttribute = Sbroenne.ExcelMcp.Core.Attributes.NoSessionAttribute;
using System.Reflection;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class CurrentContractTests
{
    private sealed class UnimplementedMacSheetCommands : MacSheetCommandsBase;

    [Fact]
    public void GeneratedBlockedActionFailsBeforeAccessingTheBatch()
    {
        var commands = new UnimplementedMacSheetCommands();
        var error = Assert.Throws<PlatformNotSupportedException>(() => commands.Copy(null!, "Source", "Target"));
        Assert.Equal(MacCommandCapabilities.Get("sheet.copy").UnavailableMessage, error.Message);
    }

    [Fact]
    public void GeneratedAvailableActionCannotBecomeSuccessWithoutAnOverride()
    {
        var commands = new UnimplementedMacSheetCommands();
        var error = Assert.Throws<PlatformNotSupportedException>(() => commands.List(null!));
        Assert.Contains("verified Mac command override", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MacCommandBasesImplementEverySessionContractWithExactSignaturesAndDefaults()
    {
        var contracts = typeof(ISheetCommands).Assembly.GetTypes()
            .Where(type => type.IsInterface && type.IsDefined(typeof(ServiceCategoryAttribute), false)
                && !type.IsDefined(typeof(NoSessionAttribute), false)).ToArray();
        Assert.NotEmpty(contracts);
        foreach (var contract in contracts)
        {
            var category = contract.GetCustomAttribute<ServiceCategoryAttribute>()!;
            var baseType = typeof(ExcelMcpService).Assembly.GetType(
                $"Sbroenne.ExcelMcp.Service.Mac.Mac{category.PascalName}CommandsBase");
            Assert.NotNull(baseType);
            Assert.True(baseType.IsAbstract && !baseType.IsPublic);
            Assert.True(contract.IsAssignableFrom(baseType), contract.FullName);
            var mapping = baseType.GetInterfaceMap(contract);
            foreach (var method in contract.GetMethods())
            {
                var index = Array.IndexOf(mapping.InterfaceMethods, method);
                Assert.True(index >= 0, method.Name);
                var implementation = mapping.TargetMethods[index];
                Assert.True(implementation.IsVirtual && !implementation.IsFinal);
                Assert.Equal(method.ReturnType, implementation.ReturnType);
                var expected = method.GetParameters();
                var actual = implementation.GetParameters();
                Assert.Equal(expected.Length, actual.Length);
                for (var parameter = 0; parameter < expected.Length; parameter++)
                {
                    Assert.Equal(expected[parameter].Name, actual[parameter].Name);
                    Assert.Equal(expected[parameter].ParameterType, actual[parameter].ParameterType);
                    Assert.Equal(expected[parameter].HasDefaultValue, actual[parameter].HasDefaultValue);
                    if (expected[parameter].HasDefaultValue)
                    {
                        Assert.Equal(expected[parameter].DefaultValue, actual[parameter].DefaultValue);
                    }
                }
            }
        }
    }

    [Theory]
    [InlineData(typeof(IRangeLinkCommands), "SetCellProtection", "SetCellLock")]
    [InlineData(typeof(IRangeLinkCommands), "GetCellProtection", "GetCellLock")]
    [InlineData(typeof(ICalculationModeCommands), "GetSettings", "GetMode")]
    [InlineData(typeof(ICalculationModeCommands), "SetSettings", "SetMode")]
    public void CurrentMethodsReplaceRemovedContracts(Type contract, string current, string removed)
    {
        Assert.NotNull(contract.GetMethod(current));
        Assert.Null(contract.GetMethod(removed));
    }

    [Fact]
    public void TableFilterUsesTheTypedContract()
    {
        var method = typeof(ITableColumnCommands).GetMethod("ApplyFilter");
        Assert.NotNull(method);
        Assert.Equal("FilterOptions", method.GetParameters()[^1].ParameterType.Name);
        Assert.Null(typeof(ITableColumnCommands).GetMethod("ApplyFilterValues"));
    }

    [Theory]
    [InlineData("rangelink.set-cell-protection")]
    [InlineData("rangelink.get-cell-protection")]
    public void ChangedContractsRemainGatedUntilParityIsReverified(string command)
    {
        Assert.False(MacCommandCapabilities.Get(command).IsAvailable);
    }

    [Fact]
    public void OfficeAddInCannotEnableUnverifiedMacActions()
    {
        Assert.DoesNotContain("OfficeAddIn", Enum.GetNames<MacCapabilityTier>());
        Assert.DoesNotContain(MacCommandCapabilities.Inventory,
            item => item.RequiredTier.ToString() == "OfficeAddIn");
    }

    [Fact]
    public async Task FileTestCannotReplaceExcelValidationWithAPathOnlySuccess()
    {
        using var service = new ExcelMcpService(new MacExcelBackend());
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.test",
            Args = System.Text.Json.JsonSerializer.Serialize(
                new { filePath = Path.GetFullPath("Missing.xlsx") })
        });
        Assert.False(response.Success);
        Assert.Equal("PlatformNotSupported", response.ErrorCategory);
    }
}
