using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Collection("Mac backend state")]
[Trait("RequiresExcel", "false")]
public sealed class MacGeneratedDispatchTests
{
    [Fact]
    public void EveryMacSessionContractHasARegisteredPlatformImplementation()
    {
        var commandSet = typeof(ExcelMcpService).Assembly.GetType("Sbroenne.ExcelMcp.Service.PlatformCommandSet");
        Assert.NotNull(commandSet);
        var createMac = commandSet.GetMethod("CreateMac");
        Assert.NotNull(createMac);
        var commands = createMac.Invoke(null, null);
        foreach (var contract in typeof(Sbroenne.ExcelMcp.Core.Commands.ISheetCommands).Assembly.GetTypes()
                     .Where(type => type.IsInterface
                         && type.IsDefined(typeof(Sbroenne.ExcelMcp.Core.Attributes.ServiceCategoryAttribute), false)
                         && !type.IsDefined(typeof(Sbroenne.ExcelMcp.Core.Attributes.NoSessionAttribute), false)))
        {
            var category = contract.GetCustomAttributes(typeof(Sbroenne.ExcelMcp.Core.Attributes.ServiceCategoryAttribute), false)
                .Cast<Sbroenne.ExcelMcp.Core.Attributes.ServiceCategoryAttribute>().Single();
            var property = commandSet.GetProperty(category.PascalName ?? throw new InvalidOperationException("Missing category name."));
            Assert.NotNull(property);
            Assert.Equal(contract, property.PropertyType);
            var implementation = property.GetValue(commands);
            Assert.NotNull(implementation);
            Assert.IsAssignableFrom(contract, implementation);
            var mapping = implementation.GetType().GetInterfaceMap(contract);
            var actionEnum = contract.Assembly.GetType($"Sbroenne.ExcelMcp.Generated.{category.PascalName}Action");
            Assert.NotNull(actionEnum);
            var toolAttribute = contract.GetCustomAttributes(typeof(Sbroenne.ExcelMcp.Core.Attributes.McpToolAttribute), false)
                .Cast<Sbroenne.ExcelMcp.Core.Attributes.McpToolAttribute>().SingleOrDefault();
            var commandCategory = toolAttribute?.ToolName.Replace("_", "")
                ?? category.PascalName.ToLowerInvariant();
            var registry = typeof(Sbroenne.ExcelMcp.Generated.ServiceRegistry).GetNestedType(category.PascalName!);
            Assert.NotNull(registry);
            var toActionString = registry.GetMethod("ToActionString");
            Assert.NotNull(toActionString);
            foreach (var method in contract.GetMethods())
            {
                var action = toActionString.Invoke(null, [Enum.Parse(actionEnum, method.Name)]);
                var capability = MacCommandCapabilities.Get($"{commandCategory}.{action}");
                var target = mapping.TargetMethods[Array.IndexOf(mapping.InterfaceMethods, method)];
                Assert.Equal(capability.IsAvailable, target.DeclaringType == implementation.GetType());
            }
        }
    }

    [Theory]
    [InlineData("range.get-values", """{"rangeAddress":"A1"}""", "sheetName")]
    [InlineData("range.set-values", """{"sheetName":"Sheet1","rangeAddress":"A1","values":[[1]],"overwritePolicy":"not-a-policy"}""", "overwritePolicy")]
    public async Task InvalidInputsFailBeforeCallingExcel(string command, string arguments, string parameter)
    {
        var rangeCalls = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (start.ArgumentList.Any(argument => argument.StartsWith("range.", StringComparison.Ordinal))) rangeCalls++;
            return Task.FromResult(new MacProcessResult(0, """{"success":true}""", ""));
        });
        var directory = Directory.CreateTempSubdirectory("excelmcp-generated-dispatch-");
        try
        {
            using var service = new ExcelMcpService(backend);
            var created = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.create",
                Args = JsonSerializer.Serialize(new { filePath = Path.Combine(directory.FullName, "test.xlsx") })
            });
            Assert.True(created.Success, created.ErrorMessage);
            using var session = JsonDocument.Parse(created.Result!);
            var sessionId = session.RootElement.GetProperty("sessionId").GetString();
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = sessionId,
                Args = arguments
            });
            Assert.False(response.Success);
            Assert.Equal("InvalidInput", response.ErrorCategory);
            Assert.Contains(parameter, response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(command, response.Command);
            Assert.Equal(sessionId, response.SessionId);
            Assert.Equal(0, rangeCalls);
        }
        finally
        {
            File.Delete(Path.Combine(directory.FullName, "test.xlsx"));
            directory.Delete();
        }
    }
}
