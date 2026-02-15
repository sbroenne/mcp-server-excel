using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

/// <summary>
/// System tray icon for the ExcelMCP CLI daemon process.
/// Shows running sessions and allows closing them or stopping the service.
/// Ported from the old ExcelMcp.Service.ServiceTray with auto-update removed.
/// </summary>
internal sealed class CliServiceTray : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly ContextMenuStrip _contextMenu;
    private readonly ToolStripMenuItem _sessionsMenu;
    private readonly SessionManager _sessionManager;
    private readonly Action _requestShutdown;
    private readonly System.Windows.Forms.Timer _refreshTimer;
    private bool _disposed;
    private DateTime _lastBalloonShown = DateTime.MinValue;

    public CliServiceTray(SessionManager sessionManager, Action requestShutdown)
    {
        _sessionManager = sessionManager;
        _requestShutdown = requestShutdown;

        _contextMenu = new ContextMenuStrip();

        // Sessions submenu
        _sessionsMenu = new ToolStripMenuItem("Sessions (0)");
        _sessionsMenu.DropDownItems.Add(new ToolStripMenuItem("No active sessions") { Enabled = false });
        _contextMenu.Items.Add(_sessionsMenu);

        _contextMenu.Items.Add(new ToolStripSeparator());

        // About
        var aboutItem = new ToolStripMenuItem("About...");
        aboutItem.Click += (_, _) => ShowAbout();
        _contextMenu.Items.Add(aboutItem);

        _contextMenu.Items.Add(new ToolStripSeparator());

        // Exit
        var exitItem = new ToolStripMenuItem("Exit");
        exitItem.Click += (_, _) => ExitService();
        _contextMenu.Items.Add(exitItem);

        // Load icon
        var icon = LoadEmbeddedIcon();

        _notifyIcon = new NotifyIcon
        {
            Icon = icon,
            Text = "ExcelMCP CLI Service",
            ContextMenuStrip = _contextMenu,
            Visible = true
        };

        _notifyIcon.DoubleClick += (_, _) => ShowSessions();

        // Refresh timer
        _refreshTimer = new System.Windows.Forms.Timer { Interval = 2000 };
        _refreshTimer.Tick += (_, _) => RefreshSessionsMenu();
        _refreshTimer.Start();

        RefreshSessionsMenu();

        // Check for updates after a short delay so the UI is responsive at startup
        CheckForUpdateAsync();
    }

    private static Icon LoadEmbeddedIcon()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = "Sbroenne.ExcelMcp.CLI.Resources.excelcli.ico";

        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream != null)
        {
            return new Icon(stream);
        }

        return SystemIcons.Application;
    }

    /// <summary>
    /// Checks NuGet for a newer version after a 5-second delay and shows a balloon tip if available.
    /// </summary>
    private async void CheckForUpdateAsync()
    {
        try
        {
            await Task.Delay(TimeSpan.FromSeconds(5));

            if (_disposed) return;

            var currentVersion = GetCurrentVersion();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cts.Token);

            if (_disposed || latestVersion == null) return;

            if (CompareVersions(currentVersion, latestVersion) < 0)
            {
                ShowBalloon(
                    "Update Available",
                    $"ExcelMCP CLI {latestVersion} is available (current: {currentVersion}).\n" +
                    "Run: dotnet tool update --global Sbroenne.ExcelMcp.CLI");
            }
        }
        catch
        {
            // Version check should never crash the service
        }
    }

    private void RefreshSessionsMenu()
    {
        if (_disposed) return;

        try
        {
            var sessions = _sessionManager.GetActiveSessions();

            if (_contextMenu.InvokeRequired)
            {
                _contextMenu.Invoke(RefreshSessionsMenu);
                return;
            }

            _sessionsMenu.Text = $"Sessions ({sessions.Count})";
            _sessionsMenu.DropDownItems.Clear();

            if (sessions.Count == 0)
            {
                _sessionsMenu.DropDownItems.Add(new ToolStripMenuItem("No active sessions") { Enabled = false });
            }
            else
            {
                foreach (var session in sessions)
                {
                    var fileName = Path.GetFileName(session.FilePath);
                    var sessionMenu = new ToolStripMenuItem(fileName);
                    sessionMenu.ToolTipText = $"Session: {session.SessionId}\nPath: {session.FilePath}";

                    // Close session with save prompt
                    var closeItem = new ToolStripMenuItem("Close Session...");
                    closeItem.Click += (_, _) => PromptCloseSession(session.SessionId, fileName);
                    sessionMenu.DropDownItems.Add(closeItem);

                    _sessionsMenu.DropDownItems.Add(sessionMenu);
                }

                _sessionsMenu.DropDownItems.Add(new ToolStripSeparator());

                var closeAllItem = new ToolStripMenuItem("Close All Sessions");
                closeAllItem.Click += (_, _) => CloseAllSessions();
                _sessionsMenu.DropDownItems.Add(closeAllItem);
            }

            _notifyIcon.Text = sessions.Count > 0
                ? $"ExcelMCP CLI - {sessions.Count} session(s)"
                : "ExcelMCP CLI Service";
        }
        catch (Exception)
        {
            // UI refresh errors should not crash the service
        }
    }

    private void PromptCloseSession(string sessionId, string fileName)
    {
        var result = MessageBox.Show(
            $"Do you want to save changes to '{fileName}' before closing?",
            "Close Session",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (result == System.Windows.Forms.DialogResult.Cancel)
            return;

        CloseSession(sessionId, save: result == System.Windows.Forms.DialogResult.Yes);
    }

    private void CloseSession(string sessionId, bool save)
    {
        try
        {
            _sessionManager.CloseSession(sessionId, save: save);
            RefreshSessionsMenu();
            ShowBalloon("Session Closed",
                save ? "Session saved and closed." : "Session closed without saving.");
        }
        catch (Exception ex)
        {
            ShowBalloon("Error", $"Failed to close session: {ex.Message}", ToolTipIcon.Error);
        }
    }

    private void CloseAllSessions()
    {
        try
        {
            var sessions = _sessionManager.GetActiveSessions().ToList();
            foreach (var session in sessions)
            {
                _sessionManager.CloseSession(session.SessionId, save: false);
            }
            RefreshSessionsMenu();
            ShowBalloon("Sessions Closed", $"Closed {sessions.Count} session(s).");
        }
        catch (Exception ex)
        {
            ShowBalloon("Error", $"Failed to close sessions: {ex.Message}", ToolTipIcon.Error);
        }
    }

    private void ShowSessions()
    {
        // Debounce: prevent duplicate balloons when clicking on/near balloon tips
        if ((DateTime.Now - _lastBalloonShown).TotalSeconds < 2)
            return;

        var sessions = _sessionManager.GetActiveSessions();
        if (sessions.Count == 0)
        {
            ShowBalloon("ExcelMCP CLI Service", "No active sessions.");
        }
        else
        {
            var message = string.Join("\n", sessions.Select(s => $"• {Path.GetFileName(s.FilePath)}"));
            ShowBalloon($"Active Sessions ({sessions.Count})", message);
        }
    }

    private void ShowBalloon(string title, string message, ToolTipIcon icon = ToolTipIcon.Info)
    {
        _lastBalloonShown = DateTime.Now;
        _notifyIcon.ShowBalloonTip(3000, title, message, icon);
    }

    private static async void ShowAbout()
    {
        var version = GetCurrentVersion();

        string? latestVersion = null;
        try
        {
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(3));
            latestVersion = await NuGetVersionChecker.GetLatestVersionAsync(cts.Token);
        }
        catch
        {
            // Version check failed — show dialog without update info
        }

        var updateAvailable = latestVersion != null && CompareVersions(version, latestVersion) < 0;

        using var form = new Form
        {
            Text = "About ExcelMCP CLI",
            Size = new Size(420, updateAvailable ? 300 : 260),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterScreen,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false
        };

        var iconBox = new PictureBox
        {
            Image = SystemIcons.Information.ToBitmap(),
            SizeMode = PictureBoxSizeMode.AutoSize,
            Location = new Point(20, 20)
        };

        var nameLabel = new Label
        {
            Text = "ExcelMCP CLI Service",
            Font = new Font(Control.DefaultFont.FontFamily, 10, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(70, 20)
        };

        var versionLabel = new Label
        {
            Text = $"Version: {version}",
            AutoSize = true,
            Location = new Point(70, 45)
        };

        var descLabel = new Label
        {
            Text = "Excel automation for coding agents.",
            AutoSize = true,
            Location = new Point(70, 75)
        };

        const string githubUrl = "https://github.com/sbroenne/mcp-server-excel";
        const string docsUrl = "https://excelmcpserver.dev/";

        var githubLabel = new Label { Text = "GitHub:", AutoSize = true, Location = new Point(70, 105) };
        var githubLink = new LinkLabel { Text = githubUrl, AutoSize = true, Location = new Point(125, 105) };
        githubLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(githubUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        var docsLabel = new Label { Text = "Docs:", AutoSize = true, Location = new Point(70, 130) };
        var docsLink = new LinkLabel { Text = docsUrl, AutoSize = true, Location = new Point(125, 130) };
        docsLink.Click += (_, _) =>
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docsUrl) { UseShellExecute = true }); }
            catch { /* Ignore navigation errors */ }
        };

        var buttonY = 165;
        form.Controls.AddRange([iconBox, nameLabel, versionLabel, descLabel, githubLabel, githubLink, docsLabel, docsLink]);

        if (updateAvailable)
        {
            var updateLabel = new Label
            {
                Text = $"Update available: {version} \u2192 {latestVersion}",
                ForeColor = Color.DarkGreen,
                Font = new Font(Control.DefaultFont, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(70, 160)
            };

            var updateCmd = new TextBox
            {
                Text = "dotnet tool update --global Sbroenne.ExcelMcp.CLI",
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = form.BackColor,
                Location = new Point(70, 180),
                Size = new Size(320, 20)
            };

            form.Controls.AddRange([updateLabel, updateCmd]);
            buttonY = 210;
        }

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = System.Windows.Forms.DialogResult.OK,
            Size = new Size(80, 28),
            Location = new Point(160, buttonY)
        };
        form.AcceptButton = okButton;
        form.Controls.Add(okButton);

        form.ShowDialog();
    }

    private static int CompareVersions(string current, string latest)
    {
        if (Version.TryParse(current, out var currentVer) && Version.TryParse(latest, out var latestVer))
            return currentVer.CompareTo(latestVer);
        return string.Compare(current, latest, StringComparison.Ordinal);
    }

    private static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        return informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString() ?? "0.0.0";
    }

    private void ExitService()
    {
        var sessions = _sessionManager.GetActiveSessions();
        if (sessions.Count > 0)
        {
            var result = MessageBox.Show(
                $"There are {sessions.Count} active session(s).\n\n" +
                "Do you want to save all sessions before exiting?",
                "Exit ExcelMCP CLI",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == System.Windows.Forms.DialogResult.Cancel)
                return;

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    foreach (var session in sessions)
                    {
                        _sessionManager.CloseSession(session.SessionId, save: true);
                    }
                    ShowBalloon("Sessions Saved", $"Saved and closed {sessions.Count} session(s).");
                }
                catch (Exception ex)
                {
                    var continueResult = MessageBox.Show(
                        $"Error saving sessions: {ex.Message}\n\nExit anyway?",
                        "Error",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Error);

                    if (continueResult != System.Windows.Forms.DialogResult.Yes)
                        return;
                }
            }
            else
            {
                try
                {
                    foreach (var session in sessions)
                    {
                        _sessionManager.CloseSession(session.SessionId, save: false);
                    }
                }
                catch (Exception ex)
                {
                    ShowBalloon("Warning", $"Error closing sessions: {ex.Message}", ToolTipIcon.Warning);
                }
            }
        }

        _requestShutdown();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _refreshTimer.Stop();
        _refreshTimer.Dispose();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _contextMenu.Dispose();
    }
}
