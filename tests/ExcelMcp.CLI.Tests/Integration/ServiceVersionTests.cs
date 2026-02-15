using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using SharedProtocol = Sbroenne.ExcelMcp.ComInterop.ServiceClient.ServiceProtocol;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for service version negotiation.
/// Validates that the service reports its version and that version matching works.
/// </summary>
[Collection("Service")]
[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("Feature", "ServiceVersion")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ServiceVersionTests
{
    private readonly ITestOutputHelper _output;

    public ServiceVersionTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public async Task ServiceVersion_ReturnsVersion()
    {
        using var client = new ServiceClient(connectTimeout: TimeSpan.FromSeconds(5));
        var response = await client.SendAsync(
            new Sbroenne.ExcelMcp.Service.ServiceRequest { Command = "service.version" });

        _output.WriteLine($"Response: Success={response.Success}, Result={response.Result}");

        Assert.True(response.Success);
        Assert.NotNull(response.Result);

        using var doc = JsonDocument.Parse(response.Result);
        var version = doc.RootElement.GetProperty("version").GetString();
        Assert.False(string.IsNullOrWhiteSpace(version));
        _output.WriteLine($"Service version: {version}");
    }

    [Fact]
    public async Task ServiceVersion_MatchesClientVersion()
    {
        using var client = new ServiceClient(connectTimeout: TimeSpan.FromSeconds(5));
        var response = await client.SendAsync(
            new Sbroenne.ExcelMcp.Service.ServiceRequest { Command = "service.version" });

        Assert.True(response.Success);
        Assert.NotNull(response.Result);

        using var doc = JsonDocument.Parse(response.Result);
        var serviceVersion = doc.RootElement.GetProperty("version").GetString();
        var clientVersion = SharedProtocol.Version;

        _output.WriteLine($"Service: {serviceVersion}, Client: {clientVersion}");
        Assert.Equal(clientVersion, serviceVersion);
    }

    [Fact]
    public async Task ValidateServiceVersion_SucceedsWhenVersionsMatch()
    {
        // Should not throw when versions match (same build)
        await ServiceManager.ValidateServiceVersionAsync();
    }

    [Fact]
    public void ClientVersion_IsNotEmpty()
    {
        var version = SharedProtocol.Version;
        _output.WriteLine($"Client version: {version}");
        Assert.False(string.IsNullOrWhiteSpace(version));
    }
}
