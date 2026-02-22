using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI;

internal sealed class Program
{
    private static readonly string[] VersionFlags = ["--version", "-v"];
    private static readonly string[] QuietFlags = ["--quiet", "-q"];

    private static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        // Register assembly resolver for office.dll (Microsoft.Office.Core), which is a
        // .NET Framework GAC assembly that .NET Core cannot find via standard probing.
        // office.dll is copied to our output directory by Directory.Build.targets.
        RegisterOfficeAssemblyResolver();

        // Determine if we should show the banner:
        // - Not when --quiet/-q flag is passed
        // - Not when output is redirected (piped to another process or file)
        var isQuiet = args.Any(arg => QuietFlags.Contains(arg, StringComparer.OrdinalIgnoreCase));
        var isPiped = Console.IsOutputRedirected;
        var showBanner = !isQuiet && !isPiped;

        // Remove --quiet/-q from args before passing to Spectre.Console.Cli
        var filteredArgs = args.Where(arg => !QuietFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)).ToArray();

        if (filteredArgs.Length == 0)
        {
            if (showBanner) RenderHeader();
            AnsiConsole.MarkupLine("[dim]No command supplied. Use [green]--help[/] for usage examples.[/]");
            return 0;
        }

        if (filteredArgs.Any(arg => VersionFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)))
        {
            return await HandleVersionAsync();
        }

        // Handle "service run" — runs the CLI daemon with tray icon (no banner)
        // Optional: --pipe-name <name> to override the default CLI pipe (used by tests)
        if (filteredArgs.Length >= 2
            && string.Equals(filteredArgs[0], "service", StringComparison.OrdinalIgnoreCase)
            && string.Equals(filteredArgs[1], "run", StringComparison.OrdinalIgnoreCase))
        {
            string? pipeNameOverride = null;
            for (int i = 2; i < filteredArgs.Length - 1; i++)
            {
                if (string.Equals(filteredArgs[i], "--pipe-name", StringComparison.OrdinalIgnoreCase))
                {
                    pipeNameOverride = filteredArgs[i + 1];
                    break;
                }
            }
            return RunServiceDaemon(pipeNameOverride);
        }

        if (showBanner) RenderHeader();

        var app = new CommandApp();

        app.Configure(config =>
        {
            config.SetApplicationName("excelcli");
            config.SetApplicationVersion(GetCurrentVersion());
            config.SetExceptionHandler((ex, _) =>
            {
                AnsiConsole.MarkupLine($"[red]Unhandled error:[/] {ex.Message.EscapeMarkup()}");
            });

            // Service lifecycle commands
            config.AddBranch("service", branch =>
            {
                branch.SetDescription("Service lifecycle management: start, stop, status.");
                branch.AddCommand<ServiceStartCommand>("start")
                    .WithDescription("Start the ExcelMCP Service if not already running.");
                branch.AddCommand<ServiceStopCommand>("stop")
                    .WithDescription("Gracefully stop the ExcelMCP Service.");
                branch.AddCommand<ServiceStatusCommand>("status")
                    .WithDescription("Show service status (running, PID, sessions, uptime).");
            });

            // Batch command — execute multiple commands in a single process launch
            config.AddCommand<BatchCommand>("batch")
                .WithDescription("Execute multiple commands from a JSON file or stdin. Outputs NDJSON (one result per line).");

            // Session commands
            config.AddBranch("session", branch =>
            {
                branch.SetDescription("Session management. WORKFLOW: open -> use sessionId -> close (--save to persist).");
                branch.AddCommand<SessionCreateCommand>("create")
                    .WithDescription("Create a new Excel file, open it, and create a session.");
                branch.AddCommand<SessionOpenCommand>("open")
                    .WithDescription("Open an Excel file and create a session.");
                branch.AddCommand<SessionCloseCommand>("close")
                    .WithDescription("Close a session. Use --save to persist changes.");
                branch.AddCommand<SessionListCommand>("list")
                    .WithDescription("List active sessions.");
                branch.AddCommand<SessionSaveCommand>("save")
                    .WithDescription("Save a session without closing it.");
            });

            // Sheet commands
            // =============================================
            // All service commands are auto-generated from
            // Core interfaces marked with [ServiceCategory].
            // =============================================
            CliCommandRegistration.RegisterCommands(config);
        });

        try
        {
            return app.Run(filteredArgs);
        }
        catch (CommandRuntimeException ex)
        {
            AnsiConsole.MarkupLine($"[red]Command error:[/] {ex.Message.EscapeMarkup()}");
            return -1;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Fatal error:[/] {ex.Message.EscapeMarkup()}");
            if (AnsiConsole.Profile.Capabilities.Ansi)
            {
                AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            }
            return -1;
        }
    }

    private static void RegisterOfficeAssemblyResolver()
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name);
            if (!string.Equals(name.Name, "office", StringComparison.OrdinalIgnoreCase))
                return null;

            return ResolveOfficeDll();
        };
    }

    /// <summary>
    /// Resolves office.dll (Microsoft.Office.Core) from multiple locations.
    /// office.dll is a .NET Framework GAC assembly that .NET Core cannot find automatically.
    /// It is present when Microsoft Office is installed, but not in the .NET Core probing paths.
    /// Search order:
    ///   1. AppContext.BaseDirectory (copied by Directory.Build.targets in local dev builds)
    ///   2. .NET Framework GAC - v16 then v15 (v15 is accepted by the CLR for v16 requests)
    ///   3. Office installation directory (click-to-run Office 365 doesn't register in GAC)
    /// </summary>
    private static Assembly? ResolveOfficeDll()
    {
        // 1. Local build output (Directory.Build.targets copies office.dll here in dev builds)
        var localPath = Path.Combine(AppContext.BaseDirectory, "office.dll");
        if (File.Exists(localPath))
            return Assembly.LoadFrom(localPath);

        // 2. .NET Framework GAC — v16 preferred, v15 accepted (CLR honours AssemblyResolve return regardless of version)
        string[] gacPaths =
        [
            @"C:\Windows\assembly\GAC_MSIL\office\16.0.0.0__71e9bce111e9429c\OFFICE.DLL",
            @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL",
        ];
        foreach (var gacPath in gacPaths)
        {
            if (File.Exists(gacPath))
                return Assembly.LoadFrom(gacPath);
        }

        // 3. Office 365 click-to-run installation directories (Office registers its own copy)
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string[] officeDirs =
        [
            Path.Combine(programFiles, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-inv16"),
            Path.Combine(programFiles, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-inv16"),
            Path.Combine(programFilesX86, @"Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in"),
        ];
        foreach (var dir in officeDirs)
        {
            var officePath = Path.Combine(dir, "OFFICE.dll");
            if (File.Exists(officePath))
                return Assembly.LoadFrom(officePath);
        }

        return null;
    }

    private static void RenderHeader()
    {
        // Write banner to stderr so it never pollutes JSON output on stdout,
        // regardless of whether stdout is piped, redirected, or captured
        // (Console.IsOutputRedirected is false in VS Code integrated terminal
        // even when capturing with $result = excelcli ...).
        var err = AnsiConsole.Create(new AnsiConsoleSettings { Out = new AnsiConsoleOutput(Console.Error) });
        err.Write(new FigletText("Excel CLI").Color(Spectre.Console.Color.Blue));
        err.MarkupLine("[dim]Excel automation powered by ExcelMcp Core[/]");
        err.MarkupLine("[yellow]Workflow:[/] [green]session open <file>[/] → run commands with [green]--session <id>[/] → [green]session close --save[/].");
        err.MarkupLine("[dim]A background service manages sessions for performance.[/]");
        err.WriteLine();
    }

    private static async Task<int> HandleVersionAsync()
    {
        var currentVersion = GetCurrentVersion();
        var latestVersion = await NuGetVersionChecker.GetLatestVersionAsync();
        var updateAvailable = latestVersion != null && CompareVersions(currentVersion, latestVersion) < 0;

        // Always show banner for version output
        RenderHeader();

        // Show friendly update message if available
        if (updateAvailable)
        {
            AnsiConsole.MarkupLine($"[yellow]⚠ Update available:[/] [dim]{currentVersion}[/] → [green]{latestVersion}[/]");
            AnsiConsole.MarkupLine($"[cyan]Run:[/] [white]dotnet tool update --global Sbroenne.ExcelMcp.CLI[/]");
            AnsiConsole.MarkupLine($"[cyan]Release notes:[/] [blue]https://github.com/sbroenne/mcp-server-excel/releases/latest[/]");
        }
        else if (latestVersion != null)
        {
            AnsiConsole.MarkupLine($"[green]✓ You're running the latest version:[/] [white]{currentVersion}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine($"[yellow]⚠ Could not check for updates[/]");
            AnsiConsole.MarkupLine($"[dim]Current version: {currentVersion}[/]");
        }

        return 0;
    }

    private static string GetCurrentVersion()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        // Strip git hash suffix (e.g., "1.2.0+abc123" -> "1.2.0")
        return informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString() ?? "0.0.0";
    }

    private static int CompareVersions(string current, string latest)
    {
        if (Version.TryParse(current, out var currentVer) && Version.TryParse(latest, out var latestVer))
            return currentVer.CompareTo(latestVer);
        return string.Compare(current, latest, StringComparison.Ordinal);
    }

    /// <summary>
    /// Runs the CLI as a daemon process with system tray icon.
    /// The service listens on the CLI pipe name (shared across CLI invocations).
    /// Auto-exits after 10 minutes of inactivity with no active sessions.
    /// </summary>
    private static int RunServiceDaemon(string? pipeNameOverride = null)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var pipeName = pipeNameOverride ?? Service.ServiceSecurity.GetCliPipeName();
        var service = new Service.ExcelMcpService();

        // Capture the UI synchronization context after Application starts
        SynchronizationContext? uiContext = null;

        // Start pipe server on background thread with 10-minute idle timeout
        var serviceTask = Task.Run(() => service.RunAsync(pipeName, idleTimeout: TimeSpan.FromMinutes(10)));

        // When service shuts down (idle timeout or remote shutdown), exit the WinForms loop
        serviceTask.ContinueWith(_ =>
        {
            if (uiContext != null)
            {
                uiContext.Post(_ => Application.ExitThread(), null);
            }
            else
            {
                Application.ExitThread();
            }
        }, TaskScheduler.Default);

        // Run WinForms message loop with tray icon on main thread
        using var tray = new CliServiceTray(service.SessionManager, () =>
        {
            service.RequestShutdown();
            Application.ExitThread();
        });

        uiContext = SynchronizationContext.Current;
        Application.Run();

        // Wait for service to finish
        serviceTask.Wait(TimeSpan.FromSeconds(5));
        service.Dispose();

        return 0;
    }
}

