using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Commands.Session;
using Sbroenne.ExcelMcp.CLI.Commands.Slicer;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

[Trait("Layer", "CLI")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Slicers")]
[Trait("Speed", "Medium")]
public sealed class SlicerCommandIntegrationTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public SlicerCommandIntegrationTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SlicerCommand_ListSlicers_ReturnsEmptyForNewWorkbook()
    {
        var filePath = _fixture.CreateTestFile();
        using var sessionService = new SessionService();

        // Open session
        var openConsole = new TestCliConsole();
        var openCommand = new SessionOpenCommand(sessionService, openConsole);
        openCommand.Execute(null!, new SessionOpenCommand.Settings { FilePath = filePath }, CancellationToken.None);

        using var openJson = JsonDocument.Parse(openConsole.GetLastJson());
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString()!;

        // List slicers (should be empty)
        var listConsole = new TestCliConsole();
        var slicerCommand = new SlicerCommand(sessionService, new PivotTableCommands(), new TableCommands(), listConsole);
        var listExit = slicerCommand.Execute(null!, new SlicerCommand.Settings
        {
            Action = "list-slicers",
            SessionId = sessionId
        }, CancellationToken.None);

        Assert.Equal(ExitCodes.Success, listExit);
        using var listJson = JsonDocument.Parse(listConsole.GetLastJson());
        Assert.True(listJson.RootElement.GetProperty("success").GetBoolean());

        // Close session
        var closeConsole = new TestCliConsole();
        var closeCommand = new SessionCloseCommand(sessionService, closeConsole);
        closeCommand.Execute(null!, new SessionCloseCommand.Settings { SessionId = sessionId }, CancellationToken.None);
    }

    [Fact]
    public void SlicerCommand_MissingSession_ReturnsError()
    {
        using var sessionService = new SessionService();

        var console = new TestCliConsole();
        var slicerCommand = new SlicerCommand(sessionService, new PivotTableCommands(), new TableCommands(), console);
        var exit = slicerCommand.Execute(null!, new SlicerCommand.Settings
        {
            Action = "list-slicers",
            SessionId = "" // Missing session ID
        }, CancellationToken.None);

        Assert.Equal(ExitCodes.MissingSession, exit);
        Assert.True(console.ErrorMessages.Count > 0);
        Assert.Contains("Session ID is required", console.ErrorMessages[0]);
    }

    [Fact]
    public void SlicerCommand_UnknownAction_ReturnsError()
    {
        var filePath = _fixture.CreateTestFile();
        using var sessionService = new SessionService();

        // Open session
        var openConsole = new TestCliConsole();
        var openCommand = new SessionOpenCommand(sessionService, openConsole);
        openCommand.Execute(null!, new SessionOpenCommand.Settings { FilePath = filePath }, CancellationToken.None);

        using var openJson = JsonDocument.Parse(openConsole.GetLastJson());
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString()!;

        // Unknown action
        var console = new TestCliConsole();
        var slicerCommand = new SlicerCommand(sessionService, new PivotTableCommands(), new TableCommands(), console);
        var exit = slicerCommand.Execute(null!, new SlicerCommand.Settings
        {
            Action = "invalid-action",
            SessionId = sessionId
        }, CancellationToken.None);

        Assert.Equal(ExitCodes.UnknownAction, exit);
        Assert.True(console.ErrorMessages.Count > 0);
        Assert.Contains("Unknown slicer action", console.ErrorMessages[0]);

        // Close session
        var closeConsole = new TestCliConsole();
        var closeCommand = new SessionCloseCommand(sessionService, closeConsole);
        closeCommand.Execute(null!, new SessionCloseCommand.Settings { SessionId = sessionId }, CancellationToken.None);
    }

    [Fact]
    public void SlicerCommand_CreateSlicerMissingParams_ReturnsError()
    {
        var filePath = _fixture.CreateTestFile();
        using var sessionService = new SessionService();

        // Open session
        var openConsole = new TestCliConsole();
        var openCommand = new SessionOpenCommand(sessionService, openConsole);
        openCommand.Execute(null!, new SessionOpenCommand.Settings { FilePath = filePath }, CancellationToken.None);

        using var openJson = JsonDocument.Parse(openConsole.GetLastJson());
        var sessionId = openJson.RootElement.GetProperty("sessionId").GetString()!;

        // Create slicer without required params
        var console = new TestCliConsole();
        var slicerCommand = new SlicerCommand(sessionService, new PivotTableCommands(), new TableCommands(), console);
        var exit = slicerCommand.Execute(null!, new SlicerCommand.Settings
        {
            Action = "create-slicer",
            SessionId = sessionId
            // Missing: PivotTableName, FieldName, DestinationSheet, Position
        }, CancellationToken.None);

        Assert.Equal(ExitCodes.MissingParameter, exit);
        Assert.True(console.ErrorMessages.Count > 0);
        Assert.Contains("required", console.ErrorMessages[0]);

        // Close session
        var closeConsole = new TestCliConsole();
        var closeCommand = new SessionCloseCommand(sessionService, closeConsole);
        closeCommand.Execute(null!, new SessionCloseCommand.Settings { SessionId = sessionId }, CancellationToken.None);
    }
}
