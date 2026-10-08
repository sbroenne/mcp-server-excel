using Sbroenne.ExcelMcp.Service.Mac;
using Sbroenne.ExcelMcp.Core.Models;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

[Trait("RequiresExcel", "false")]
public sealed class MacHelperCompatibilityTests
{
    private static readonly string[] RequiredPrimitives = ["helper.info"];
    private static readonly string[] ExpectedMaintenancePrimitives = ["future.primitive", "helper.info"];

    [Fact]
    public void NativeTextResponsesPreserveCompatibleVersionAndConfirmBuild()
    {
        var info = MacHelperProtocol.ValidateResponse(
            JsonValue.Create("""{"version":"1.8.2","primitives":["helper.info"]}"""), RequiredPrimitives);
        Assert.Equal("1.8.2", info.Version);
        Assert.Equal(RequiredPrimitives, info.Primitives);
        MacNativeHelper.ValidateBuildResponse(JsonValue.Create("""{"success":true,"errorMessage":""}"""));
    }

    [Theory]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{}")]
    [InlineData("[]")]
    public void NonTextNativeHandshakeHasExplicitMalformedCategory(string response)
    {
        var error = Assert.Throws<MacExcelOperationException>(() =>
            MacHelperProtocol.ValidateResponse(JsonNode.Parse(response), RequiredPrimitives));
        Assert.Equal("HelperMalformed", error.ErrorCategory);
    }

    [Theory]
    [InlineData("\"not-json\"")]
    [InlineData("42")]
    [InlineData("\"{}\"")]
    [InlineData("\"{\\\"success\\\":true,\\\"errorMessage\\\":123}\"")]
    [InlineData("\"{\\\"success\\\":true,\\\"errorMessage\\\":\\\"failed\\\"}\"")]
    public void InvalidNativeBuildConfirmationIsNotSuccess(string response)
    {
        var error = Assert.Throws<MacExcelOperationException>(() =>
            MacNativeHelper.ValidateBuildResponse(JsonNode.Parse(response)));
        Assert.Equal("HelperBuildFailed", error.ErrorCategory);
    }

    [Fact]
    public void NativeBuildFailurePreservesVbaErrorDetails()
    {
        var response = JsonValue.Create("""{"success":false,"errorNumber":58,"errorMessage":"The output already exists."}""");
        var error = Assert.Throws<MacExcelOperationException>(() => MacNativeHelper.ValidateBuildResponse(response));
        Assert.Equal("HelperBuildFailed", error.ErrorCategory);
        Assert.Contains("VBA 58", error.Message, StringComparison.Ordinal);
        Assert.Contains("The output already exists.", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("service.helper-check", """{"helperVersion":"1.0.0"}""")]
    [InlineData("service.helper-build", """{"unrecognized":true}""")]
    [InlineData("service.helper-build", """{}""")]
    [InlineData("service.helper-build", """{"workbookPath":"/bootstrap.xlsx","outputPath":"/helper.xlam","helperVersion":"1.0.0"}""")]
    public async Task InvalidMaintenanceArgumentsFailBeforeCallingExcel(string command, string arguments)
    {
        var calls = 0;
        var backend = new MacExcelBackend((_, _, _) =>
        {
            calls++;
            throw new InvalidOperationException("Invalid maintenance input must not reach Excel.");
        });
        using var service = new ExcelMcpService(backend);
        var response = await service.ProcessAsync(new ServiceRequest { Command = command, Args = arguments });
        Assert.False(response.Success);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(command, response.Command);
        Assert.NotEmpty(response.ErrorMessage!);
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task MaintenanceReadinessReturnsIndependentHelperVersionAndPrimitives()
    {
        var backend = new MacExcelBackend((start, _, _) =>
        {
            Assert.Contains("helper.check", start.ArgumentList);
            return Task.FromResult(new MacProcessResult(0,
                """{"success":true,"helper":{"version":"1.8.2","primitives":["helper.info","future.primitive"]}}""", ""));
        });
        using var service = new ExcelMcpService(backend);
        var response = await service.ProcessAsync(new ServiceRequest { Command = "service.helper-check" });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
        var result = JsonNode.Parse(response.Result!)!;
        Assert.Equal("1.8.2", result["version"]!.GetValue<string>());
        Assert.Equal(ExpectedMaintenancePrimitives,
            result["primitives"]!.AsArray().Select(value => value!.GetValue<string>()));
    }
    [Theory]
    [InlineData("""{"version":"1.0.0","primitives":["helper.info"]}""", true)]
    [InlineData("""{"version":"1.9.4+build.2","primitives":["helper.info","future.primitive"]}""", true)]
    [InlineData("""{"version":"2.0.0","primitives":["helper.info"]}""", false)]
    [InlineData("""{"version":"1.0","primitives":["helper.info"]}""", false)]
    [InlineData("""{"version":"01.0.0","primitives":["helper.info"]}""", false)]
    [InlineData("""{"version":"1.0.0","primitives":[]}""", false)]
    [InlineData("""{"version":"1.0.0","primitives":["helper.info","helper.info"]}""", false)]
    [InlineData("""{"version":"1.0.0","primitives":["helper.info",null]}""", false)]
    [InlineData("null", false)]
    [InlineData("", false)]
    [InlineData("""{"version":"1.0.0-alpha.01","primitives":["helper.info"]}""", false)]
    [InlineData("""{"version":"1.0.0-alpha.1","primitives":["helper.info"]}""", true)]
    [InlineData("""{"version":"1.0.0","version":"2.0.0","primitives":["helper.info"]}""", false)]
    [InlineData("{not-json}", false)]
    public void ReadinessDependsOnHelperMajorAndRequiredPrimitivesNotServerVersion(string json, bool compatible)
    {
        var error = Record.Exception(() => MacHelperProtocol.Validate(json, RequiredPrimitives));
        if (compatible)
        {
            Assert.Null(error);
        }
        else
        {
            Assert.IsType<MacExcelOperationException>(error);
        }
    }

    [Theory]
    [InlineData("""{"version":"1.9.0","primitives":["helper.info"]}""", "HelperIncompatible")]
    [InlineData("""{"version":"2.0.0","primitives":["helper.info","probe.required"]}""", "HelperIncompatible")]
    [InlineData("null", "HelperMissing")]
    [InlineData("""{"version":"1.0.0","primitives":null}""", "HelperMalformed")]
    public void RequiredPrimitiveCheckStopsBeforeDispatchingWorkbookMutation(string handshake, string category)
    {
        var mutations = 0;
        var backend = new MacExcelBackend((start, _, _) =>
        {
            if (!start.ArgumentList.Contains("helper.check")) mutations++;
            return Task.FromResult(new MacProcessResult(0, $$"""{"success":true,"helper":{{handshake}}}""", ""));
        });
        var session = new MacExcelSession
        {
            SessionId = "readiness",
            FilePath = Path.GetFullPath("opaque.xlsx"),
            IsVisible = false,
            OperationTimeout = TimeSpan.FromSeconds(10)
        };
        using var batch = new MacExcelBatch(backend, session);
        var error = Assert.Throws<MacExcelOperationException>(() =>
            batch.Invoke<OperationResult>("range.set-values", new
            {
                sheetName = "Sheet1",
                rangeAddress = "A1",
                values = new List<List<object?>> { new() { 1 } }
            }, MutationPrimitives));
        Assert.Equal(category, error.ErrorCategory);
        Assert.Equal(0, mutations);
    }

    private static readonly string[] MutationPrimitives = ["helper.info", "probe.required"];
}
