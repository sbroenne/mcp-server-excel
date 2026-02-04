using Sbroenne.ExcelMcp.CLI.Daemon;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

/// <summary>
/// Mock dialog service for testing tray interactions.
/// Records all dialog calls and returns configured responses.
/// </summary>
internal sealed class MockDialogService : IDialogService
{
    private readonly Queue<DialogResult> _responses = new();
    private readonly List<(string Type, string Message, string Title)> _dialogCalls = new();

    /// <summary>
    /// Queue a response for the next dialog call.
    /// </summary>
    public void QueueResponse(DialogResult result) => _responses.Enqueue(result);

    /// <summary>
    /// Get all dialog calls that were made.
    /// </summary>
    public IReadOnlyList<(string Type, string Message, string Title)> DialogCalls => _dialogCalls;

    /// <summary>
    /// Clear recorded calls and queued responses.
    /// </summary>
    public void Reset()
    {
        _responses.Clear();
        _dialogCalls.Clear();
    }

    public DialogResult ShowYesNoCancel(string message, string title)
    {
        _dialogCalls.Add(("YesNoCancel", message, title));
        return _responses.Count > 0 ? _responses.Dequeue() : DialogResult.Cancel;
    }

    public DialogResult ShowOkCancel(string message, string title)
    {
        _dialogCalls.Add(("OkCancel", message, title));
        return _responses.Count > 0 ? _responses.Dequeue() : DialogResult.Cancel;
    }

    public DialogResult ShowYesNo(string message, string title)
    {
        _dialogCalls.Add(("YesNo", message, title));
        return _responses.Count > 0 ? _responses.Dequeue() : DialogResult.No;
    }

    public void ShowInfo(string message, string title)
    {
        _dialogCalls.Add(("Info", message, title));
    }

    public void ShowError(string message, string title)
    {
        _dialogCalls.Add(("Error", message, title));
    }
}

/// <summary>
/// Unit tests for DaemonTray dialog interactions.
/// Tests the decision logic without actual Windows Forms UI.
/// </summary>
[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "DaemonTray")]
[Trait("Speed", "Fast")]
public sealed class DaemonTrayDialogTests
{
    [Fact]
    public void MockDialogService_QueuedResponses_ReturnedInOrder()
    {
        // Arrange
        var mock = new MockDialogService();
        mock.QueueResponse(DialogResult.Yes);
        mock.QueueResponse(DialogResult.No);
        mock.QueueResponse(DialogResult.Cancel);

        // Act & Assert
        Assert.Equal(DialogResult.Yes, mock.ShowYesNoCancel("msg1", "title1"));
        Assert.Equal(DialogResult.No, mock.ShowYesNoCancel("msg2", "title2"));
        Assert.Equal(DialogResult.Cancel, mock.ShowYesNoCancel("msg3", "title3"));
    }

    [Fact]
    public void MockDialogService_NoQueuedResponses_ReturnsDefaults()
    {
        // Arrange
        var mock = new MockDialogService();

        // Act & Assert - defaults
        Assert.Equal(DialogResult.Cancel, mock.ShowYesNoCancel("msg", "title"));
        Assert.Equal(DialogResult.Cancel, mock.ShowOkCancel("msg", "title"));
        Assert.Equal(DialogResult.No, mock.ShowYesNo("msg", "title"));
    }

    [Fact]
    public void MockDialogService_RecordsAllCalls()
    {
        // Arrange
        var mock = new MockDialogService();

        // Act
        mock.ShowYesNoCancel("message1", "title1");
        mock.ShowOkCancel("message2", "title2");
        mock.ShowYesNo("message3", "title3");
        mock.ShowInfo("message4", "title4");
        mock.ShowError("message5", "title5");

        // Assert
        Assert.Equal(5, mock.DialogCalls.Count);
        Assert.Equal(("YesNoCancel", "message1", "title1"), mock.DialogCalls[0]);
        Assert.Equal(("OkCancel", "message2", "title2"), mock.DialogCalls[1]);
        Assert.Equal(("YesNo", "message3", "title3"), mock.DialogCalls[2]);
        Assert.Equal(("Info", "message4", "title4"), mock.DialogCalls[3]);
        Assert.Equal(("Error", "message5", "title5"), mock.DialogCalls[4]);
    }

    [Fact]
    public void MockDialogService_Reset_ClearsState()
    {
        // Arrange
        var mock = new MockDialogService();
        mock.QueueResponse(DialogResult.Yes);
        mock.ShowYesNoCancel("msg", "title");

        // Act
        mock.Reset();

        // Assert
        Assert.Empty(mock.DialogCalls);
        // After reset, should return default again
        Assert.Equal(DialogResult.Cancel, mock.ShowYesNoCancel("msg", "title"));
    }

    [Fact]
    public void WindowsFormsDialogService_Implements_IDialogService()
    {
        // Verify the production implementation exists and implements the interface
        var service = new WindowsFormsDialogService();
        Assert.IsAssignableFrom<IDialogService>(service);
    }
}
