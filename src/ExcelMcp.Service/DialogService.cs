namespace Sbroenne.ExcelMcp.Service;

/// <summary>
/// Result of a user dialog prompt.
/// Abstracts MessageBox.Show results for testability.
/// </summary>
public enum DialogResult
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

    /// <summary>
    /// Shows an About dialog with clickable hyperlinks.
    /// </summary>
    void ShowAbout(string productName, string version, string description, string githubUrl, string docsUrl);
}

/// <summary>
/// Production implementation using Windows Forms MessageBox.
/// </summary>
public sealed class WindowsFormsDialogService : IDialogService
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

    public void ShowAbout(string productName, string version, string description, string githubUrl, string docsUrl)
    {
        using var form = new System.Windows.Forms.Form
        {
            Text = "About ExcelMCP",
            Size = new Size(420, 240),
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false
        };

        // Icon
        var iconBox = new System.Windows.Forms.PictureBox
        {
            Image = System.Drawing.SystemIcons.Information.ToBitmap(),
            SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize,
            Location = new Point(20, 20)
        };

        // Product name
        var nameLabel = new System.Windows.Forms.Label
        {
            Text = productName,
            Font = new Font(System.Windows.Forms.Control.DefaultFont.FontFamily, 10, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(70, 20)
        };

        // Version
        var versionLabel = new System.Windows.Forms.Label
        {
            Text = $"Version: {version}",
            AutoSize = true,
            Location = new Point(70, 45)
        };

        // Description
        var descLabel = new System.Windows.Forms.Label
        {
            Text = description,
            AutoSize = true,
            Location = new Point(70, 75)
        };

        // GitHub link
        var githubLabel = new System.Windows.Forms.Label
        {
            Text = "GitHub:",
            AutoSize = true,
            Location = new Point(70, 105)
        };

        var githubLink = new System.Windows.Forms.LinkLabel
        {
            Text = githubUrl,
            AutoSize = true,
            Location = new Point(125, 105)
        };
        githubLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(githubUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        // Docs link
        var docsLabel = new System.Windows.Forms.Label
        {
            Text = "Docs:",
            AutoSize = true,
            Location = new Point(70, 130)
        };

        var docsLink = new System.Windows.Forms.LinkLabel
        {
            Text = docsUrl,
            AutoSize = true,
            Location = new Point(125, 130)
        };
        docsLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docsUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        // OK button
        var okButton = new System.Windows.Forms.Button
        {
            Text = "OK",
            DialogResult = System.Windows.Forms.DialogResult.OK,
            Size = new Size(80, 28),
            Location = new Point(160, 165)
        };
        form.AcceptButton = okButton;

        form.Controls.AddRange([iconBox, nameLabel, versionLabel, descLabel, githubLabel, githubLink, docsLabel, docsLink, okButton]);
        form.ShowDialog();
    }
}


