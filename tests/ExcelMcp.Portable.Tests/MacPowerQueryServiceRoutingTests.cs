using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacPowerQueryServiceRoutingTests
{
    [Fact]
    public async Task HelperMutationUsesOnlyAdvertisedHelperRoute()
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
    public async Task ReadWithoutProvenHelperIsUnsupportedWithoutWorkbookInspection()
    {
        var path = TempWorkbookPath();
        var capabilityCalls = 0;
        var helperCalls = 0;
        var backendCommands = new List<string>();
        var backend = CreateBackend(backendCommands);
        using var service = new ExcelMcpService(
            backend,
            (workbookPath, timeout) =>
            {
                capabilityCalls++;
                return Task.FromResult(Capabilities(false, "powerquery.list"));
            },
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
                Command = "powerquery.list",
                SessionId = sessionId
            });

            Assert.False(response.Success);
            Assert.Equal("PlatformNotSupported", response.ErrorCategory);
            Assert.Equal(1, capabilityCalls);
            Assert.Equal(0, helperCalls);
            Assert.DoesNotContain("workbook.state", backendCommands);
        }
        finally
        {
            File.Delete(path);
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
    public async Task ExactCandidateOptInDispatchesOnlySelectedAction()
    {
        var helperActions = new List<string>();
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(
                    false,
                    "powerquery.create",
                    "powerquery.delete")),
            (workbookPath, action, arguments, timeout) =>
            {
                helperActions.Add(action);
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            },
            new HashSet<string>(["powerquery.create"], StringComparer.Ordinal));

        try
        {
            var sessionId = await OpenAsync(service, path);
            var createResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = sessionId,
                Args =
                    """{"queryName":"Sales","mCode":"let Source = 1 in Source","loadDestination":"connection-only"}"""
            });
            var deleteResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.delete",
                SessionId = sessionId,
                Args = """{"queryName":"Sales"}"""
            });

            Assert.True(createResponse.Success, createResponse.ErrorMessage);
            Assert.False(deleteResponse.Success);
            Assert.Equal("PlatformNotSupported", deleteResponse.ErrorCategory);
            Assert.Equal(["powerquery.create"], helperActions);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task DefaultCreateDispatchesAtomicWorksheetContract()
    {
        JsonObject? observedArguments = null;
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(true, "powerquery.create")),
            (workbookPath, action, arguments, timeout) =>
            {
                Assert.Equal("powerquery.create", action);
                observedArguments = Assert.IsType<JsonObject>(arguments);
                return Task.FromResult(JsonSerializer.SerializeToElement(new { }));
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.create",
                SessionId = sessionId,
                Args = """{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""
            });

            Assert.True(response.Success, response.ErrorMessage);
            Assert.NotNull(observedArguments);
            Assert.Equal(
                "load-to-table",
                observedArguments["destination"]!.GetValue<string>());
            Assert.Equal("Sales", observedArguments["sheetName"]!.GetValue<string>());
            Assert.Equal("A1", observedArguments["cellAddress"]!.GetValue<string>());
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
    public async Task UncertainHelperMutationIsNotRetriedThroughAnotherRoute()
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
    public async Task TimeoutWithWorkbookStillOpenRequiresRecovery()
    {
        var helperCalls = 0;
        var path = TempWorkbookPath();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            var command = AutomationCommandOrNull(start);
            if (command == "session.close")
            {
                throw new TimeoutException("Close timed out.");
            }
            var output = command == "session.is-open"
                ? """{"success":true,"errorMessage":"","open":true}"""
                : """{"success":true,"errorMessage":""}""";
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });
        using var service = new ExcelMcpService(
            backend,
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(true, "powerquery.rename")),
            (workbookPath, action, arguments, timeout) =>
            {
                helperCalls++;
                throw new TimeoutException("Helper mutation timed out.");
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var timeoutResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.rename",
                SessionId = sessionId,
                Args = """{"oldName":"Sales","newName":"Revenue"}"""
            });
            var retryResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.rename",
                SessionId = sessionId,
                Args = """{"oldName":"Sales","newName":"Revenue"}"""
            });

            Assert.False(timeoutResponse.Success);
            Assert.Equal("Timeout", timeoutResponse.ErrorCategory);
            Assert.False(retryResponse.Success);
            Assert.Contains(
                "requires manual recovery",
                retryResponse.ErrorMessage,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, helperCalls);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task RollbackFailureWithWorkbookStillOpenRequiresRecovery()
    {
        var helperCalls = 0;
        var path = TempWorkbookPath();
        var backend = new MacExcelBackend((start, input, cancellationToken) =>
        {
            var command = AutomationCommandOrNull(start);
            if (command == "session.close")
            {
                throw new TimeoutException("Close timed out.");
            }
            var output = command == "session.is-open"
                ? """{"success":true,"errorMessage":"","open":true}"""
                : """{"success":true,"errorMessage":""}""";
            return Task.FromResult(new MacProcessResult(0, output, ""));
        });
        using var service = new ExcelMcpService(
            backend,
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(true, "powerquery.update")),
            (workbookPath, action, arguments, timeout) =>
            {
                helperCalls++;
                throw new MacVbaHelperException(
                    "RecoveryRequired",
                    "rollback_failed",
                    "The helper could not restore the original query state.");
            });

        try
        {
            var sessionId = await OpenAsync(service, path);
            var failureResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = sessionId,
                Args = """{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""
            });
            var retryResponse = await service.ProcessAsync(new ServiceRequest
            {
                Command = "powerquery.update",
                SessionId = sessionId,
                Args = """{"queryName":"Sales","mCode":"let Source = 1 in Source"}"""
            });

            Assert.False(failureResponse.Success);
            Assert.Equal("RecoveryRequired", failureResponse.ErrorCategory);
            Assert.False(retryResponse.Success);
            Assert.Contains(
                "requires manual recovery",
                retryResponse.ErrorMessage,
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, helperCalls);
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
            (workbookPath, timeout) =>
                Task.FromResult(Capabilities(false, "powerquery.refresh")),
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

    [Fact]
    public async Task RefreshSharesExplicitTimeoutAcrossCapabilityProbeAndDispatch()
    {
        TimeSpan? capabilityTimeout = null;
        TimeSpan? dispatchTimeout = null;
        var path = TempWorkbookPath();
        using var service = new ExcelMcpService(
            CreateBackend([]),
            async (workbookPath, timeout) =>
            {
                capabilityTimeout = timeout;
                await Task.Delay(20);
                return Capabilities("powerquery.refresh");
            },
            (workbookPath, action, arguments, timeout) =>
            {
                dispatchTimeout = timeout;
                return Task.FromResult(JsonSerializer.SerializeToElement(new
                {
                    queryName = "Sales",
                    hasErrors = false,
                    errorMessages = Array.Empty<string>(),
                    refreshTime = DateTimeOffset.UtcNow,
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

            Assert.True(response.Success, response.ErrorMessage);
            Assert.Equal(TimeSpan.FromSeconds(17), capabilityTimeout);
            Assert.NotNull(dispatchTimeout);
            Assert.True(dispatchTimeout < capabilityTimeout);
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
            return Task.FromResult(new MacProcessResult(
                0,
                """{"success":true,"errorMessage":""}""",
                ""));
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
        bool actionsProven,
        params string[] actions)
    {
        var provenMethods = new JsonObject
        {
            ["powerQueryList"] = actionsProven
        };
        foreach (var action in actions)
        {
            var proofProperty = action switch
            {
                "powerquery.create" => "powerQueryCreate",
                "powerquery.update" => "powerQueryUpdate",
                "powerquery.rename" => "powerQueryRename",
                "powerquery.delete" => "powerQueryDelete",
                "powerquery.refresh" => "powerQueryRefresh",
                "powerquery.refresh-all" => "powerQueryRefreshAll",
                "powerquery.load-to" => "powerQueryLoadTo",
                "powerquery.unload" => "powerQueryUnload",
                "powerquery.evaluate" => "powerQueryEvaluate",
                _ => null
            };
            if (proofProperty is not null)
            {
                provenMethods[proofProperty] = actionsProven;
            }
        }
        return JsonSerializer.SerializeToElement(new
        {
            helperVersion = "1.0.0",
            protocolVersion = 1,
            supportedActions = actions,
            provenMethods
        });
    }

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
