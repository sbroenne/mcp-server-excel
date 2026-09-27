using System.Diagnostics;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryServiceRoutingTests
{
    [Fact]
    public async Task HelperMutationSelectsRouteBeforeWorkbookPackageInspection()
    {
        var backendCommands = new List<string>();
        var helperActions = new List<string>();
        var backend = CreateBackend(backendCommands);
        using var service = new ExcelMcpService(
            backend,
            (path, timeout) => Task.FromResult(Capabilities("powerquery.rename")),
            (path, action, arguments, timeout) =>
            {
                helperActions.Add(action);
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            });
        var path = TempWorkbookPath();

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.rename",
                SessionId = sessionId,
                Args = """{"oldName":"Sales","newName":"Revenue"}"""
            });

            Assert.True(response.Success);
            Assert.Equal(["powerquery.rename"], helperActions);
            Assert.DoesNotContain("workbook.state", backendCommands);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task PackageReadDoesNotProbeOptionalHelper()
    {
        var directory = Directory.CreateTempSubdirectory("excelmcp-pq-routing-");
        var path = Path.Combine(directory.FullName, "empty.xlsx");
        using (System.IO.Compression.ZipFile.Open(
                   path,
                   System.IO.Compression.ZipArchiveMode.Create))
        {
        }
        var capabilityCalls = 0;
        var backendCommands = new List<string>();
        var backend = CreateBackend(backendCommands);
        using var service = new ExcelMcpService(
            backend,
            (workbookPath, timeout) =>
            {
                capabilityCalls++;
                return Task.FromResult(Capabilities("powerquery.list"));
            },
            (workbookPath, action, arguments, timeout) =>
                throw new InvalidOperationException("Helper must not run."));
        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.list",
                SessionId = sessionId
            });

            Assert.True(response.Success);
            Assert.Equal(0, capabilityCalls);
            Assert.Contains("workbook.state", backendCommands);
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    [Fact]
    public async Task AdvertisedButUnprovenMutationDoesNotDispatch()
    {
        var helperCalls = 0;
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(false, "powerquery.rename")),
            (workbookPath, action, arguments, timeout) =>
            {
                helperCalls++;
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.rename",
                SessionId = sessionId,
                Args = """{"oldName":"Sales","newName":"Revenue"}"""
            });

            Assert.False(response.Success);
            Assert.Equal("PlatformNotSupported", response.ErrorCategory);
            Assert.Equal(0, helperCalls);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task StructuredHelperFailurePreservesPublicErrorCategory()
    {
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, timeout) => Task.FromResult(Capabilities("powerquery.rename")),
            (workbookPath, action, arguments, timeout) =>
                throw new MacVbaHelperException(
                    "Conflict",
                    "query_conflict",
                    "A query with that name already exists."));

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.rename",
                SessionId = sessionId,
                Args = """{"oldName":"Sales","newName":"Revenue"}"""
            });

            Assert.False(response.Success);
            Assert.Equal("Conflict", response.ErrorCategory);
            Assert.Equal(
                "A query with that name already exists.",
                response.ErrorMessage);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task UncertainHelperMutationIsNotRetriedWithPackageMutation()
    {
        var helperCalls = 0;
        var backendCommands = new List<string>();
        var backend = CreateBackend(backendCommands);
        using var service = new ExcelMcpService(
            backend,
            (path, timeout) => Task.FromResult(Capabilities("powerquery.update")),
            (path, action, arguments, timeout) =>
            {
                helperCalls++;
                throw new TimeoutException("Helper mutation timed out.");
            });
        var path = TempWorkbookPath();

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = sessionId,
                Args = """{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""
            });

            Assert.False(response.Success);
            Assert.Equal("Timeout", response.ErrorCategory);
            Assert.Equal(1, helperCalls);
            Assert.DoesNotContain("workbook.state", backendCommands);
            Assert.DoesNotContain("session.close-if-saved", backendCommands);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RefreshIsNotRoutedWhenMethodIsNotInProvenContract()
    {
        TimeSpan? observedTimeout = null;
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, timeout) => Task.FromResult(Capabilities("powerquery.refresh")),
            (workbookPath, action, arguments, timeout) =>
            {
                observedTimeout = timeout;
                return Task.FromResult(JsonSerializer.SerializeToElement(new
                {
                    queryName = "Sales",
                    hasErrors = false,
                    errorMessages = Array.Empty<string>(),
                    refreshTime = DateTime.UtcNow,
                    isConnectionOnly = false
                }));
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.refresh",
                SessionId = sessionId,
                Args = """{"queryName":"Sales","timeout":17}"""
            });

            Assert.False(response.Success);
            Assert.Equal("PlatformNotSupported", response.ErrorCategory);
            Assert.Null(observedTimeout);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("-1")]
    [InlineData("1.5")]
    [InlineData("2147484")]
    public async Task RefreshRejectsTimeoutOutsidePublicContract(string timeout)
    {
        var helperCalls = 0;
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, operationTimeout) =>
                Task.FromResult(Capabilities("powerquery.refresh")),
            (workbookPath, action, arguments, operationTimeout) =>
            {
                helperCalls++;
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.refresh",
                SessionId = sessionId,
                Args = $$"""{"queryName":"Sales","timeout":{{timeout}}}"""
            });

            Assert.False(response.Success);
            Assert.Equal("InvalidInput", response.ErrorCategory);
            Assert.Equal(0, helperCalls);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static MacExcelBackend CreateBackend(List<string> commands) =>
        new((start, input, cancellationToken) =>
        {
            var command = AutomationCommandOrNull(start);
            if (command is not null)
            {
                commands.Add(command);
            }
            var output = command == "workbook.state"
                ? """{"success":true,"errorMessage":"","saved":true}"""
                : """{"success":true,"errorMessage":""}""";
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });

    private static string TempWorkbookPath() =>
        Path.Combine(Path.GetTempPath(), $"excelmcp-pq-{Guid.NewGuid():N}.xlsx");

    private static async Task<string> OpenAsync(ExcelMcpService service, string path)
    {
        if (!File.Exists(path))
        {
            File.WriteAllBytes(path, []);
        }
        var response = await service.ProcessAsync(new ServiceRequest
        {
            Command = "session.open",
            Args = $$"""{"filePath":"{{path}}","show":false,"timeoutSeconds":10}"""
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(Assert.IsType<string>(response.Result));
        return result.RootElement.GetProperty("sessionId").GetString()!;
    }

    private static JsonElement Capabilities(params string[] actions) =>
        Capabilities(true, actions);

    private static JsonElement Capabilities(
        bool powerQueryMutationProven,
        params string[] actions) =>
        JsonSerializer.SerializeToElement(new
        {
            helperVersion = "1.0.0",
            protocolVersion = 1,
            supportedActions = actions,
            provenMethods = new
            {
                powerQueryList = true,
                powerQueryMutation = powerQueryMutationProven
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
