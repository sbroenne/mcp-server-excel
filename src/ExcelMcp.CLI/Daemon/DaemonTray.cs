using System.Reflection;
using Sbroenne.ExcelMcp.ComInterop.Session;

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
    private readonly SessionManager _sessionManager;
    private readonly Action _requestShutdown;
    private readonly System.Windows.Forms.Timer _refreshTimer;
    private bool _disposed;

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

        // Status item
        var statusItem = new ToolStripMenuItem("Excel CLI Daemon") { Enabled = false };
        _contextMenu.Items.Add(statusItem);

        _contextMenu.Items.Add(new ToolStripSeparator());

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

                    // Close without save
                    var closeItem = new ToolStripMenuItem("Close");
                    closeItem.Click += (_, _) => CloseSession(session.SessionId, save: false);
                    sessionMenu.DropDownItems.Add(closeItem);

                    // Close with save
                    var saveCloseItem = new ToolStripMenuItem("Save && Close");
                    saveCloseItem.Click += (_, _) => CloseSession(session.SessionId, save: true);
                    sessionMenu.DropDownItems.Add(saveCloseItem);

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
    /// Can be called from any thread (will invoke on UI thread if needed).
    /// </summary>
    public void ShowUpdateNotification(string title, string message)
    {
        if (_disposed) return;

        // Ensure we're on the UI thread
        if (_notifyIcon.InvokeRequired)
        {
            _notifyIcon.Invoke(() => ShowBalloon(title, message, ToolTipIcon.Info));
        }
        else
        {
            ShowBalloon(title, message, ToolTipIcon.Info);
        }
    }

    private void StopDaemon()
    {
        var sessions = _sessionManager.GetActiveSessions();
        if (sessions.Count > 0)
        {
            var result = MessageBox.Show(
                $"There are {sessions.Count} active session(s). Close all sessions and stop the daemon?",
                "Stop Excel CLI Daemon",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
            {
                return;
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
