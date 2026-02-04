using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Daemon;

/// <summary>
/// System tray icon for the Excel CLI daemon.
/// Shows running sessions and allows closing them or stopping the daemon.
/// </summary>
internal sealed class DaemonTray : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly ContextMenuStrip _contextMenu;
    private readonly ToolStripMenuItem _sessionsMenu;
    private readonly ToolStripMenuItem? _updateMenuItem;
    private readonly SessionManager _sessionManager;
    private readonly Action _requestShutdown;
    private readonly System.Windows.Forms.Timer _refreshTimer;
    private bool _disposed;
    private UpdateInfo? _availableUpdate;

    public DaemonTray(SessionManager sessionManager, Action requestShutdown)
    {
        _sessionManager = sessionManager;
        _requestShutdown = requestShutdown;

        // Initialize Windows Forms
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Create context menu
        _contextMenu = new ContextMenuStrip();

        // Sessions submenu
        _sessionsMenu = new ToolStripMenuItem("Sessions (0)");
        _sessionsMenu.DropDownItems.Add(new ToolStripMenuItem("No active sessions") { Enabled = false });
        _contextMenu.Items.Add(_sessionsMenu);

        _contextMenu.Items.Add(new ToolStripSeparator());

        // Update menu item (initially hidden, shown when update is available)
        _updateMenuItem = new ToolStripMenuItem("Update CLI")
        {
            Visible = false
        };
        _updateMenuItem.Click += (_, _) => UpdateCli();
        _contextMenu.Items.Add(_updateMenuItem);

        // Stop daemon
        var stopItem = new ToolStripMenuItem("Stop Daemon");
        stopItem.Click += (_, _) => StopDaemon();
        _contextMenu.Items.Add(stopItem);

        // Load icon from embedded resource
        var icon = LoadEmbeddedIcon();

        // Create notify icon
        _notifyIcon = new NotifyIcon
        {
            Icon = icon,
            Text = "Excel CLI Daemon",
            ContextMenuStrip = _contextMenu,
            Visible = true
        };

        _notifyIcon.DoubleClick += (_, _) => ShowSessions();

        // Timer to refresh sessions menu periodically
        _refreshTimer = new System.Windows.Forms.Timer { Interval = 2000 };
        _refreshTimer.Tick += (_, _) => RefreshSessionsMenu();
        _refreshTimer.Start();

        // Initial refresh
        RefreshSessionsMenu();
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

        // Fallback to system icon
        return SystemIcons.Application;
    }

    private void RefreshSessionsMenu()
    {
        if (_disposed) return;

        try
        {
            var sessions = _sessionManager.GetActiveSessions();

            // Update on UI thread
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

                    // Close session (with save prompt)
                    var closeItem = new ToolStripMenuItem("Close Session...");
                    closeItem.Click += (_, _) => PromptCloseSession(session.SessionId, fileName);
                    sessionMenu.DropDownItems.Add(closeItem);

                    _sessionsMenu.DropDownItems.Add(sessionMenu);
                }

                // Add separator and "Close All" option
                _sessionsMenu.DropDownItems.Add(new ToolStripSeparator());

                var closeAllItem = new ToolStripMenuItem("Close All Sessions");
                closeAllItem.Click += (_, _) => CloseAllSessions();
                _sessionsMenu.DropDownItems.Add(closeAllItem);
            }

            // Update tooltip with session count
            _notifyIcon.Text = sessions.Count > 0
                ? $"Excel CLI Daemon - {sessions.Count} session(s)"
                : "Excel CLI Daemon";
        }
        catch
        {
            // Ignore errors during refresh
        }
    }

    private void PromptCloseSession(string sessionId, string fileName)
    {
        var result = MessageBox.Show(
            $"Do you want to save changes to '{fileName}' before closing?",
            "Close Session",
            MessageBoxButtons.YesNoCancel,
            MessageBoxIcon.Question);

        if (result == DialogResult.Cancel)
            return;

        CloseSession(sessionId, save: result == DialogResult.Yes);
    }

    private void CloseSession(string sessionId, bool save)
    {
        try
        {
            _sessionManager.CloseSession(sessionId, save: save);
            RefreshSessionsMenu();
            ShowBalloon("Session Closed", save ? "Session saved and closed." : "Session closed without saving.");
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
        var sessions = _sessionManager.GetActiveSessions();
        if (sessions.Count == 0)
        {
            ShowBalloon("Excel CLI Daemon", "No active sessions.");
        }
        else
        {
            var message = string.Join("\n", sessions.Select(s => $"• {Path.GetFileName(s.FilePath)}"));
            ShowBalloon($"Active Sessions ({sessions.Count})", message);
        }
    }

    private void ShowBalloon(string title, string message, ToolTipIcon icon = ToolTipIcon.Info)
    {
        _notifyIcon.ShowBalloonTip(3000, title, message, icon);
    }

    /// <summary>
    /// Shows an update notification to the user.
    /// Thread-safe - can be called from any thread.
    /// </summary>
    public void ShowUpdateNotification(string title, string message)
    {
        if (_disposed) return;

        // Store update info for the update menu
        _availableUpdate = new UpdateInfo
        {
            CurrentVersion = message.Contains("current:") ? message.Split("current:")[1].Split(')')[0].Trim() : "unknown",
            LatestVersion = message.Contains("Version") ? message.Split("Version")[1].Split("is")[0].Trim() : "unknown",
            UpdateAvailable = true
        };

        // Show update menu option
        if (_updateMenuItem != null && _contextMenu.InvokeRequired)
        {
            _contextMenu.Invoke(() =>
            {
                _updateMenuItem.Visible = true;
                _updateMenuItem.Text = $"Update to {_availableUpdate.LatestVersion}";
            });
        }
        else if (_updateMenuItem != null)
        {
            _updateMenuItem.Visible = true;
            _updateMenuItem.Text = $"Update to {_availableUpdate.LatestVersion}";
        }

        // Create a custom balloon tip with clickable instructions
        var fullMessage = message + "\n\nClick the 'Update CLI' menu option to update.";
        ShowBalloon(title, fullMessage, ToolTipIcon.Info);
    }

    private void UpdateCli()
    {
        if (_availableUpdate == null)
            return;

        var updateCommand = ToolInstallationDetector.GetUpdateCommand();
        
        // Show confirmation dialog with update command
        var result = MessageBox.Show(
            $"Update Excel CLI from {_availableUpdate.CurrentVersion} to {_availableUpdate.LatestVersion}?\n\n" +
            $"This will run:\n{updateCommand}\n\n" +
            "The daemon will restart after the update.",
            "Update Excel CLI",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Question);

        if (result != DialogResult.OK)
            return;

        // Show progress
        ShowBalloon("Updating...", "Please wait while the CLI is updated.", ToolTipIcon.Info);

        // Run update in background
        Task.Run(async () =>
        {
            var (success, output) = await ToolInstallationDetector.TryUpdateAsync();
            
            // Show result on UI thread
            if (_contextMenu.InvokeRequired)
            {
                _contextMenu.Invoke(() => ShowUpdateResult(success, output));
            }
            else
            {
                ShowUpdateResult(success, output);
            }
        });
    }

    private void ShowUpdateResult(bool success, string output)
    {
        if (success)
        {
            var result = MessageBox.Show(
                "CLI updated successfully!\n\nThe daemon will now restart to use the new version.",
                "Update Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            // Hide update menu item
            if (_updateMenuItem != null)
            {
                _updateMenuItem.Visible = false;
            }
            _availableUpdate = null;

            // Restart daemon
            _requestShutdown();
        }
        else
        {
            var updateCommand = ToolInstallationDetector.GetUpdateCommand();
            MessageBox.Show(
                $"Update failed:\n{output}\n\nYou can manually update by running:\n{updateCommand}",
                "Update Failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private void StopDaemon()
    {
        var sessions = _sessionManager.GetActiveSessions();
        if (sessions.Count > 0)
        {
            var result = MessageBox.Show(
                $"There are {sessions.Count} active session(s).\n\n" +
                "Do you want to save all sessions before stopping the daemon?",
                "Stop Excel CLI Daemon",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Cancel)
            {
                return;
            }

            // Save all sessions if requested
            if (result == DialogResult.Yes)
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
                        $"Error saving sessions: {ex.Message}\n\nStop daemon anyway?",
                        "Error",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Error);

                    if (continueResult != DialogResult.Yes)
                        return;
                }
            }
            else
            {
                // Close without saving
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
