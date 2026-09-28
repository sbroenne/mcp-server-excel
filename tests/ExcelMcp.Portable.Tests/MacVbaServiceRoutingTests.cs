using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacVbaServiceRoutingTests
{
    private static readonly string[] MarkerProcedures = ["WriteMarker"];

    [Fact]
    public async Task AdvertisedButUnprovenListDoesNotDispatch()
    {
        var helperCalls = 0;
        var path = TempWorkbookPath();
        using var service = CreateService(
            Capabilities(false, "vba.list"),
            (workbookPath, action, arguments, timeout) =>
            {
                helperCalls++;
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            });

        var sessionId = await OpenAsync(service, path);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "vba.list",
            SessionId = sessionId
        });

        Assert.False(response.Success);
        Assert.Equal("PlatformNotSupported", response.ErrorCategory);
        Assert.Equal(0, helperCalls);
        File.Delete(path);
    }

    [Fact]
    public async Task ExactListOptInMapsStrictPublicResult()
    {
        var path = TempWorkbookPath();
        using var service = CreateService(
            Capabilities(false, "vba.list"),
            (workbookPath, action, arguments, timeout) =>
            {
                Assert.Equal("vba.list", action);
                return Task.FromResult(JsonSerializer.SerializeToElement(new
                {
                    modules = new[]
                    {
                        new
                        {
                            name = "Module1",
                            type = 1,
                            lineCount = 3,
                            procedures = MarkerProcedures
                        }
                    }
                }));
            },
            "vba.list");

        var sessionId = await OpenAsync(service, path);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "vba.list",
            SessionId = sessionId
        });

        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(Assert.IsType<string>(response.Result));
        var script = Assert.Single(result.RootElement.GetProperty("scripts").EnumerateArray());
        Assert.Equal("Module1", script.GetProperty("name").GetString());
        Assert.Equal("Module", script.GetProperty("type").GetString());
        Assert.Equal("WriteMarker", Assert.Single(script.GetProperty("procedures").EnumerateArray()).GetString());
        File.Delete(path);
    }

    [Fact]
    public async Task ImportResolvesFileAndMapsSourceArgument()
    {
        var path = TempWorkbookPath();
        var sourcePath = Path.Combine(Path.GetTempPath(), $"excelmcp-vba-{Guid.NewGuid():N}.bas");
        await File.WriteAllTextAsync(sourcePath, "Option Explicit\nPublic Sub Probe()\nEnd Sub");
        JsonObject? observed = null;
        using var service = CreateService(
            Capabilities(false, "vba.import"),
            (workbookPath, action, arguments, timeout) =>
            {
                observed = JsonSerializer.SerializeToNode(arguments)!.AsObject();
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            },
            "vba.import");

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "vba.import",
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(new
                {
                    moduleName = "Module1",
                    vbaCodeFile = sourcePath
                })
            });

            Assert.True(response.Success, response.ErrorMessage);
            Assert.Equal("Module1", observed!["moduleName"]!.GetValue<string>());
            Assert.Equal(
                "Option Explicit\nPublic Sub Probe()\nEnd Sub",
                observed["source"]!.GetValue<string>().ReplaceLineEndings("\n"));
        }
        finally
        {
            File.Delete(path);
            File.Delete(sourcePath);
        }
    }

    [Fact]
    public async Task RunUsesWorkbookQualifiedHelperContractAndPublicTimeout()
    {
        var path = TempWorkbookPath();
        JsonObject? observed = null;
        TimeSpan? observedTimeout = null;
        using var service = CreateService(
            Capabilities(false, "vba.run"),
            (workbookPath, action, arguments, timeout) =>
            {
                observed = JsonSerializer.SerializeToNode(arguments)!.AsObject();
                observedTimeout = timeout;
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            },
            "vba.run");

        var sessionId = await OpenAsync(service, path);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "vba.run",
            SessionId = sessionId,
            Args = """{"procedureName":"FixtureModule.WriteMarker","timeout":17,"parameters":["accepted"]}"""
        });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal("FixtureModule.WriteMarker", observed!["procedureName"]!.GetValue<string>());
        Assert.Equal("accepted", Assert.Single(observed["parameters"]!.AsArray())!.GetValue<string>());
        Assert.NotNull(observedTimeout);
        Assert.True(observedTimeout <= TimeSpan.FromSeconds(17));
        File.Delete(path);
    }

    private static ExcelMcpService CreateService(
        JsonElement capabilities,
        MacPowerQueryHelperDispatch dispatch,
        params string[] enabledActions) =>
        new(
            CreateBackend(),
            (workbookPath, timeout) => Task.FromResult(capabilities),
            dispatch,
            macVbaCandidateActions: new HashSet<string>(enabledActions, StringComparer.Ordinal),
            macVbaPreflight: new MacVbaPreflightResult(
                MacMacroExecutionAvailability.Available,
                MacVbaProjectModelAccess.Enabled));

    private static MacExcelBackend CreateBackend() =>
        new((start, input, cancellationToken) =>
        {
            var command = AutomationCommandOrNull(start);
            return Task.FromResult(new MacProcessResult(
                0,
                """{"success":true,"errorMessage":""}""",
                ""));
        });

    private static string TempWorkbookPath() =>
        Path.Combine(Path.GetTempPath(), $"excelmcp-vba-{Guid.NewGuid():N}.xlsm");

    private static async Task<string> OpenAsync(ExcelMcpService service, string path)
    {
        File.WriteAllBytes(path, []);
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = $$"""{"filePath":"{{path}}","show":false,"timeoutSeconds":10}"""
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(Assert.IsType<string>(response.Result));
        return result.RootElement.GetProperty("sessionId").GetString()!;
    }

    private static JsonElement Capabilities(bool proven, params string[] actions) =>
        JsonSerializer.SerializeToElement(new
        {
            supportedActions = actions,
            trustReadiness = new
            {
                vbaProjectReadable = true
            },
            provenMethods = new
            {
                vbaListView = proven,
                vbaMutation = proven,
                vbaRun = proven
            }
        });

    private static string? AutomationCommandOrNull(ProcessStartInfo start)
    {
        if (start.FileName == "/usr/bin/open")
        {
            return null;
        }
        var arguments = start.ArgumentList.ToArray();
        var marker = Array.IndexOf(arguments, MacAutomationHost.Marker);
        Assert.True(marker >= 0 && marker + 1 < arguments.Length);
        return arguments[marker + 1];
    }
}
