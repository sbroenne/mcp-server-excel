using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands.Session;
using Sbroenne.ExcelMcp.CLI.Commands.Sheet;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
[Trait("Speed", "Medium")]
public sealed class SessionCommandIntegrationTests : IClassFixture<TempDirectoryFixture>
{
    private readonly string _tempDir;

    public SessionCommandIntegrationTests(TempDirectoryFixture fixture)
    {
        _tempDir = fixture.TempDir;
    }

    [Fact]
    public async Task SessionCommands_OpenListClose_ManagesLifecycle()
    {
        var filePath = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SessionCommandIntegrationTests),
            nameof(SessionCommands_OpenListClose_ManagesLifecycle),
            _tempDir);

        using var sessionService = new SessionService();

        var openConsole = new TestCliConsole();
        var openCommand = new SessionOpenCommand(sessionService, openConsole);
        var openExit = openCommand.Execute(null!, new SessionOpenCommand.Settings
        {
            FilePath = filePath
        }, CancellationToken.None);

        Assert.Equal(0, openExit);
        using var openJson = JsonDocument.Parse(openConsole.GetLastJson());
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));

        var listConsole = new TestCliConsole();
        var listCommand = new SessionListCommand(sessionService, listConsole);
        var listExit = listCommand.Execute(null!, CancellationToken.None);

        Assert.Equal(0, listExit);
        using var listJson = JsonDocument.Parse(listConsole.GetLastJson());
        var sessions = listJson.RootElement.GetProperty("sessions").EnumerateArray().ToList();
        Assert.Contains(sessions, s => s.GetProperty("sessionId").GetString() == sessionId);

        var closeConsole = new TestCliConsole();
        var closeCommand = new SessionCloseCommand(sessionService, closeConsole);
        var closeExit = closeCommand.Execute(null!, new SessionCloseCommand.Settings
        {
            SessionId = sessionId!
        }, CancellationToken.None);

        Assert.Equal(0, closeExit);

        var finalListConsole = new TestCliConsole();
        var finalListCommand = new SessionListCommand(sessionService, finalListConsole);
        var finalListExit = finalListCommand.Execute(null!, CancellationToken.None);

        Assert.Equal(0, finalListExit);
        using var finalListJson = JsonDocument.Parse(finalListConsole.GetLastJson());
        Assert.Empty(finalListJson.RootElement.GetProperty("sessions").EnumerateArray());
    }

    [Fact]
    public async Task SheetCommand_CreateAndList_Worksheets()
    {
        var filePath = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SessionCommandIntegrationTests),
            nameof(SheetCommand_CreateAndList_Worksheets),
            _tempDir);

        using var sessionService = new SessionService();

        var openConsole = new TestCliConsole();
        var openCommand = new SessionOpenCommand(sessionService, openConsole);
        openCommand.Execute(null!, new SessionOpenCommand.Settings
        {
            FilePath = filePath
        }, CancellationToken.None);

        using var openJson = JsonDocument.Parse(openConsole.GetLastJson());
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString()!;

        var createConsole = new TestCliConsole();
        var sheetCommand = new SheetCommand(sessionService, new SheetCommands(), createConsole);
        var createExit = sheetCommand.Execute(null!, new SheetCommand.Settings
        {
            Action = "create",
            SessionId = sessionId,
            SheetName = "CliSheet"
        }, CancellationToken.None);

        Assert.Equal(0, createExit);
        using var createJson = JsonDocument.Parse(createConsole.GetLastJson());
        Assert.True(createJson.RootElement.GetProperty("success").GetBoolean());

        var listConsole = new TestCliConsole();
        var sheetListCommand = new SheetCommand(sessionService, new SheetCommands(), listConsole);
        var listExit = sheetListCommand.Execute(null!, new SheetCommand.Settings
        {
            Action = "list",
            SessionId = sessionId
        }, CancellationToken.None);

        Assert.Equal(0, listExit);
        using var listJson = JsonDocument.Parse(listConsole.GetLastJson());
        var worksheets = listJson.RootElement.GetProperty("worksheets").EnumerateArray().ToList();
        Assert.Contains(worksheets, w => string.Equals(w.GetProperty("name").GetString(), "CliSheet", StringComparison.OrdinalIgnoreCase));

        var closeConsole = new TestCliConsole();
        var closeCommand = new SessionCloseCommand(sessionService, closeConsole);
        var closeExit = closeCommand.Execute(null!, new SessionCloseCommand.Settings
        {
            SessionId = sessionId
        }, CancellationToken.None);

        Assert.Equal(0, closeExit);
    }
}
