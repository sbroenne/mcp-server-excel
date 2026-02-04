namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// Result of a user dialog prompt.
/// Abstracts MessageBox.Show results for testability.
/// </summary>
internal enum DialogResult
{
    Yes,
    No,
    Cancel,
    OK
}

/// <summary>
/// Interface for dialog interactions, enabling unit testing of tray logic.
/// </summary>
internal interface IDialogService
{
    /// <summary>
    /// Shows a Yes/No/Cancel dialog.
    /// </summary>
    DialogResult ShowYesNoCancel(string message, string title);

    /// <summary>
    /// Shows an OK/Cancel dialog.
    /// </summary>
    DialogResult ShowOkCancel(string message, string title);

    /// <summary>
    /// Shows a Yes/No dialog.
    /// </summary>
    DialogResult ShowYesNo(string message, string title);

    /// <summary>
    /// Shows an information dialog.
    /// </summary>
    void ShowInfo(string message, string title);

    /// <summary>
    /// Shows an error dialog.
    /// </summary>
    void ShowError(string message, string title);
}

/// <summary>
/// Production implementation using Windows Forms MessageBox.
/// </summary>
internal sealed class WindowsFormsDialogService : IDialogService
{
    public DialogResult ShowYesNoCancel(string message, string title)
    {
        var result = System.Windows.Forms.MessageBox.Show(
            message, title,
            System.Windows.Forms.MessageBoxButtons.YesNoCancel,
            System.Windows.Forms.MessageBoxIcon.Question);

        return result switch
        {
            System.Windows.Forms.DialogResult.Yes => DialogResult.Yes,
            System.Windows.Forms.DialogResult.No => DialogResult.No,
            _ => DialogResult.Cancel
        };
    }

    public DialogResult ShowOkCancel(string message, string title)
    {
        var result = System.Windows.Forms.MessageBox.Show(
            message, title,
            System.Windows.Forms.MessageBoxButtons.OKCancel,
            System.Windows.Forms.MessageBoxIcon.Question);

        return result == System.Windows.Forms.DialogResult.OK
            ? DialogResult.OK
            : DialogResult.Cancel;
    }

    public DialogResult ShowYesNo(string message, string title)
    {
        var result = System.Windows.Forms.MessageBox.Show(
            message, title,
            System.Windows.Forms.MessageBoxButtons.YesNo,
            System.Windows.Forms.MessageBoxIcon.Error);

        return result == System.Windows.Forms.DialogResult.Yes
            ? DialogResult.Yes
            : DialogResult.No;
    }

    public void ShowInfo(string message, string title)
    {
        System.Windows.Forms.MessageBox.Show(
            message, title,
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }

    public void ShowError(string message, string title)
    {
        System.Windows.Forms.MessageBox.Show(
            message, title,
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Error);
    }
}
