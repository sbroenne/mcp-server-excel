using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Generated;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Telemetry;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI;

internal sealed class Program
{
    private static readonly string[] VersionFlags = ["--version", "-v"];
    private static readonly string[] QuietFlags = ["--quiet", "-q"];
    private static readonly HashSet<string> StdinSentinelOptions = new(
        ["--input", "-i", "--values", "--formulas", "--formats", "--rows"],
        StringComparer.OrdinalIgnoreCase);

    private static Task<int> Main(string[] args) => RunEntryPointAsync(args);

    /// <summary>
    /// The process entry path: enables telemetry, runs the command, then signals reporting exit.
    /// Tests call it directly so they cover the same routes as <c>Main</c>.
    /// </summary>
    internal static async Task<int> RunEntryPointAsync(
        string[] args,
        CliCommandRuntime? runtime = null)
    {
        if (Service.ExcelController.IsWorkerMode(args))
            return await Service.ExcelController.RunWorkerAsync(args);
        using var runtimeScope = CliCommandRuntime.Push(runtime ?? CliCommandRuntime.Current);
        InitializeOperation(args);
        try
        {
            try
            {
                CliTelemetry.Enable();
            }
            catch (InvalidOperationException ex)
            {
                return CliErrorOutput.WriteException(ex);
            }
            return await RunAsync(args, runtime, initializeOperation: false);
        }
        finally
        {
            CliTelemetry.SignalExit();
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Top-level CLI boundary writes the exact exception as JSON or stderr diagnostics and returns nonzero; it must not let a fatal command failure appear successful.")]
    internal static async Task<int> RunAsync(
        string[] args,
        CliCommandRuntime? runtime = null,
        bool initializeOperation = true)
    {
        if (runtime is null)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
        }

        using var runtimeScope = CliCommandRuntime.Push(runtime ?? CliCommandRuntime.Current);
        if (initializeOperation)
            InitializeOperation(args);

        // Determine if we should show the banner:
        // - Not when --quiet/-q flag is passed
        // - Not when output is redirected (piped to another process or file)
        var isQuiet = args.Any(arg => QuietFlags.Contains(arg, StringComparer.OrdinalIgnoreCase));
        var isPiped = CliCommandRuntime.Current.IsOutputRedirected;
        var showBanner = !isQuiet && !isPiped;

        // Remove quiet flags and normalize standalone dash option values before Spectre parsing.
        var filteredArgs = NormalizeStandaloneDashOptionValues(
            args.Where(arg => !QuietFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)).ToArray());

        if (filteredArgs.Length == 0)
        {
            if (showBanner) RenderHeader();
            WriteDiagnosticMarkupLine("[dim]No command supplied. Use [green]--help[/] for commands. For a known sequence in one call, use [green]batch --help[/].[/]");
            return 0;
        }

        if (filteredArgs.Any(arg => VersionFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)))
        {
            return await HandleVersionAsync();
        }

        if (ActionHelp.TryWrite(filteredArgs))
        {
            return 0;
        }
        // Handle "service run" — runs the CLI daemon with tray icon (no banner)
        // Optional: --pipe-name <name> to override the default CLI pipe (used by tests)
        if (filteredArgs.Length >= 2
            && string.Equals(filteredArgs[0], "service", StringComparison.OrdinalIgnoreCase)
            && string.Equals(filteredArgs[1], "run", StringComparison.OrdinalIgnoreCase))
        {
            string? pipeNameOverride = null;
            for (int i = 2; i < filteredArgs.Length; i++)
            {
                if (!string.Equals(filteredArgs[i], "--pipe-name", StringComparison.OrdinalIgnoreCase))
                    return CliErrorOutput.WriteError($"Unexpected argument '{filteredArgs[i]}' for service run. Only --pipe-name <name> is supported.");
                if (i + 1 == filteredArgs.Length ||
                    (filteredArgs[i + 1].StartsWith('-') && filteredArgs[i + 1] != "-"))
                    return CliErrorOutput.WriteError("--pipe-name requires a name for service run.");
                pipeNameOverride ??= filteredArgs[i + 1];
                i++;
            }
            CliTelemetry.StartDaemonReporting();
            var daemonRunner = CliCommandRuntime.Current.ServiceDaemonRunner ?? RunServiceDaemon;
            return daemonRunner(pipeNameOverride);
        }

        if (showBanner) RenderHeader();

        var app = new CommandApp();

        app.Configure(config =>
        {
            config.ConfigureConsole(CreateOutputConsole());
            config.SetApplicationName("excelcli");
            config.SetApplicationVersion(GetCurrentVersion());
            config.Settings.StrictParsing = true;
            config.SetExceptionHandler((ex, _) =>
            {
                CliErrorOutput.WriteException(ex);
            });

            // Service lifecycle commands
            config.AddBranch("service", branch =>
            {
                branch.SetDescription("Service lifecycle management: start, stop, status.");
                branch.AddCommand<ServiceStartCommand>("start")
                    .WithDescription("Start the ExcelMCP Service if needed and wait up to 30 seconds for readiness.");
                branch.AddCommand<ServiceStopCommand>("stop")
                    .WithDescription("Stop the CLI service normally; preserve unsaved workbooks and Excel left by earlier services.");
                branch.AddCommand<ServiceStatusCommand>("status")
                    .WithDescription("Show stopped, starting, running, or unresponsive daemon state; readiness checks wait up to 10 seconds.");
            });

            // Batch command — execute multiple commands in a single process launch
            config.AddCommand<BatchCommand>("batch")
                .WithDescription("Run a known sequence in one CLI call, avoiding repeated launches. See batch --help for input examples.");

            // Session commands
            config.AddBranch("session", branch =>
            {
                branch.SetDescription("Session management. Use discover to find an already-open workbook, then attach --workbook-id without reopening it. WORKFLOW: open or attach -> use sessionId -> close (--save to persist). Use --show for IRM/auth prompts.");
                branch.AddCommand<SessionDiscoverCommand>("discover")
                    .WithDescription("Find already-open workbooks without reopening them. Results include saved state and AutoSave setting.");
                branch.AddCommand<SessionAttachCommand>("attach")
                    .WithDescription("Attach to a workbook returned by session discover. Does not reopen the workbook; existing AutoSave remains unchanged.");
                branch.AddCommand<SessionCreateCommand>("create")
                    .WithDescription("Create a new Excel file, open it, and create a session. Add --show for visible Excel.");
                branch.AddCommand<SessionOpenCommand>("open")
                    .WithDescription("Open an Excel file with VBA macros disabled and create a session. Add --show for visible Excel.");
                branch.AddCommand<SessionNoticesCommand>("notices")
                    .WithDescription("Inspect workbook restrictions without enabling content. Reports Protected View, read-only access, known VBA opening state, and inspection gaps.");
                branch.AddCommand<SessionSecurityCommand>("enable-editing")
                    .WithDescription("Leave Protected View only with --approved and --allow-remembered-trust. Excel can remember trust for this file. Does not enable VBA macros.");
                branch.AddCommand<SessionSecurityCommand>("enable-connections")
                    .WithDescription("Enable all workbook data connections with --approved and --allow-remembered-trust. Can contact data sources or start refresh; completion is not established.");
                branch.AddCommand<SessionSecurityCommand>("enable-macros")
                    .WithDescription("Enable VBA by reopening a saved local workbook with --approved and --allow-reopen. Only for a session-owned workbook without neighbors; opening macros can run.");
                branch.AddCommand<SessionCloseCommand>("close")
                    .WithDescription("Close a session. Use --save to persist changes.");
                branch.AddCommand<SessionSaveCommand>("save")
                    .WithDescription("Save the current workbook without closing its session. Excel must be ready.");
                branch.AddCommand<SessionListCommand>("list")
                    .WithDescription("List active sessions with canClose, excelState, operationPending and blocker details. Confirmed blocking dialogs automatically reveal Excel and include readable dialog text in blockingMessage; partial or unavailable reads are reported. Ask the user to respond in Excel. Dialog types are not identified and input values are not read.");
                branch.AddCommand<SessionTestCommand>("test")
                    .WithDescription("Test file existence, validity, openability, and IRM/AIP read-only requirements.");
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
            var telemetryObserver = CliCommandRuntime.Current.TelemetryObserver;
            var exitCode = await (telemetryObserver == null
                ? CliTelemetry.TrackCliInvocationAsync(filteredArgs, () => app.RunAsync(filteredArgs))
                : CliTelemetry.TrackCliInvocationAsync(
                    filteredArgs,
                    () => app.RunAsync(filteredArgs),
                    telemetryObserver));
            if (exitCode == 0 &&
                string.Equals(filteredArgs[0], "batch", StringComparison.OrdinalIgnoreCase) &&
                filteredArgs.Any(argument => argument is "--help" or "-h"))
            {
                BatchHelp.Write();
            }
            return exitCode;
        }
        catch (CommandRuntimeException ex)
        {
            return CliErrorOutput.WriteException(ex, "InvalidInput");
        }
        catch (Exception ex)
        {
            return CliErrorOutput.WriteException(ex);
        }
    }

    private static void InitializeOperation(string[] args)
    {
        CliCommandRuntime.Current.OperationTimer.Restart();
        var parts = args.Where(arg => !QuietFlags.Contains(arg, StringComparer.OrdinalIgnoreCase))
            .TakeWhile(arg => !arg.StartsWith('-')).Take(2).ToArray();
        CliCommandRuntime.Current.CommandName = parts.Length > 0 && parts[0] is "batch" or "list-actions"
            ? parts[0]
            : string.Join(".", parts).ToLowerInvariant();
    }

    internal static string[] NormalizeStandaloneDashOptionValues(string[] args)
    {
        var normalized = new List<string>(args.Length);
        for (var index = 0; index < args.Length; index++)
        {
            var argument = args[index];
            if (index + 1 < args.Length
                && StdinSentinelOptions.Contains(argument)
                && string.Equals(args[index + 1], "-", StringComparison.Ordinal))
            {
                normalized.Add($"{argument}=-");
                index++;
                continue;
            }

            normalized.Add(argument);
        }

        return [.. normalized];
    }

    private static void RenderHeader()
    {
        // Write banner to stderr so it never pollutes JSON output on stdout,
        // regardless of whether stdout is piped, redirected, or captured
        // (Console.IsOutputRedirected is false in VS Code integrated terminal
        // even when capturing with $result = excelcli ...).
        var err = CreateErrorConsole();
        err.Write(new FigletText("Excel CLI").Color(Spectre.Console.Color.Blue));
        err.MarkupLine("[dim]Automate Microsoft Excel from the command line[/]");
        err.MarkupLine("[yellow]Workflow:[/] [green]session open <file>[/] → run commands with [green]--session <id>[/] → [green]session close --save[/].");
        err.MarkupLine("[dim]For a known sequence in one call, use [green]batch --help[/].[/]");
        err.MarkupLine("[dim]A background service manages sessions for performance.[/]");
        err.WriteLine();
    }

    private static async Task<int> HandleVersionAsync()
    {
        var currentVersion = GetCurrentVersion();
        var latestVersion = await CliCommandRuntime.Current.LatestVersionProvider();
        var updateAvailable = latestVersion != null && CompareVersions(currentVersion, latestVersion) < 0;

        // Always show banner for version output
        RenderHeader();

        // Show friendly update message if available
        if (updateAvailable)
        {
            var output = CreateOutputConsole();
            output.MarkupLine($"[yellow]⚠ Update available:[/] [dim]{currentVersion}[/] → [green]{latestVersion}[/]");
            output.MarkupLine($"[cyan]Download:[/] [blue]https://github.com/sbroenne/mcp-server-excel/releases/latest[/]");
        }
        else if (latestVersion != null)
        {
            CreateOutputConsole().MarkupLine(
                $"[green]✓ You're running the latest version:[/] [white]{currentVersion}[/]");
        }
        else
        {
            var output = CreateOutputConsole();
            output.MarkupLine("[yellow]⚠ Could not check for updates[/]");
            output.MarkupLine($"[dim]Current version: {currentVersion}[/]");
        }

        return 0;
    }

    internal static void WriteDiagnosticMarkupLine(string markup)
    {
        CreateErrorConsole().MarkupLine(markup);
    }

    private static IAnsiConsole CreateErrorConsole()
    {
        return AnsiConsole.Create(new AnsiConsoleSettings
        {
            Out = new AnsiConsoleOutput(CliCommandRuntime.Current.Error)
        });
    }

    private static IAnsiConsole CreateOutputConsole()
    {
        return AnsiConsole.Create(new AnsiConsoleSettings
        {
            Out = new AnsiConsoleOutput(CliCommandRuntime.Current.Output)
        });
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
    /// Remains running until explicitly stopped or its process exits.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Daemon process entry boundary observes the service task failure, writes its original message to stderr, returns exit code 1, and still disposes owned resources.")]
    private static int RunServiceDaemon(string? pipeNameOverride = null)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        var pipeName = pipeNameOverride ?? Service.ServiceSecurity.GetCliPipeName();

        // Acquire a named OS mutex for the lifetime of this daemon process.
        // If another daemon is already running for this pipe/user, exit immediately
        // instead of creating a duplicate process with a duplicate tray icon.
        var mutexName = DaemonAutoStart.GetDaemonMutexName(pipeName);
        var daemonMutex = new Mutex(initiallyOwned: true, mutexName, out var createdNew);
        var ownsDaemonMutex = createdNew;
        if (!createdNew)
        {
            try
            {
                ownsDaemonMutex = daemonMutex.WaitOne(TimeSpan.Zero);
            }
            catch (AbandonedMutexException)
            {
                ownsDaemonMutex = true;
            }

            if (!ownsDaemonMutex)
            {
                // Another daemon is already running — exit silently.
                daemonMutex.Dispose();
                return 0;
            }
        }
        Mutex? legacyDaemonMutex = null;
        var ownsLegacyDaemonMutex = false;
        try
        {
            legacyDaemonMutex = new Mutex(
                initiallyOwned: true,
                DaemonStartupLock.GetLegacyDaemonMutexName(pipeName),
                out var legacyCreatedNew);
            ownsLegacyDaemonMutex = legacyCreatedNew;
            if (!legacyCreatedNew)
            {
                try
                {
                    ownsLegacyDaemonMutex = legacyDaemonMutex.WaitOne(TimeSpan.Zero);
                }
                catch (AbandonedMutexException)
                {
                    ownsLegacyDaemonMutex = true;
                }
            }

            if (!ownsLegacyDaemonMutex)
            {
                legacyDaemonMutex.Dispose();
                if (ownsDaemonMutex)
                {
                    daemonMutex.ReleaseMutex();
                }
                daemonMutex.Dispose();
                return 0;
            }
        }
        catch (IOException)
        {
            legacyDaemonMutex?.Dispose();
            legacyDaemonMutex = null;
        }
        catch (ArgumentException)
        {
            legacyDaemonMutex?.Dispose();
            legacyDaemonMutex = null;
        }

        var daemonIdentity = DaemonProcessTracker.RegisterCurrentProcess(pipeName);
        Service.ExcelController? service = null;
        try
        {
            service = new Service.ExcelController(CliTelemetry.TrackDaemonRequestAsync);
            service.ExcelProcessObserved += process =>
                DaemonProcessTracker.RecordExcelProcesses(pipeName, daemonIdentity,
                    [new DaemonProcessTracker.ProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime)]);
            service.ExcelProcessPreserved += process =>
                DaemonProcessTracker.PreserveExcelProcess(pipeName, daemonIdentity,
                    new DaemonProcessTracker.ProcessIdentity(process.ProcessId, process.StartedAtUtcFileTime));

            // Capture the UI synchronization context after Application starts
            SynchronizationContext? uiContext = null;

            // Run WinForms message loop with tray icon on main thread
            using var tray = new CliServiceTray(() => service.Sessions,
                (sessionId, save) => service.ProcessAsync(new Service.ServiceRequest
                {
                    Command = "session.close",
                    SessionId = sessionId,
                    Args = System.Text.Json.JsonSerializer.Serialize(new { save })
                }), async () =>
            {
                var report = await service.ShutdownAsync();
                try
                {
                    if (report is not null)
                    {
                        await Console.Error.WriteLineAsync(report);
                        var reportPath = DaemonProcessTracker.GetTrackingFilePath(pipeName) + ".shutdown.txt";
                        await File.WriteAllTextAsync(reportPath, report);
                        await Console.Error.WriteLineAsync($"Excel was left open. Shutdown details: {reportPath}");
                    }
                }
                finally
                {
                    await service.RequestShutdownAsync();
                }
                return report;
            });

            uiContext = SynchronizationContext.Current;

            // Start accepting RPC only after daemon host initialization succeeds.
            // Otherwise auto-start clients can observe a successful ping from a pipe
            // server whose owning WinForms/tray host is still able to fail and exit.
            var serviceTask = Task.Run(() => service.RunAsync(pipeName));

            // When service shuts down, exit the WinForms loop.
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

            Application.Run();

            try
            {
                // Wait for service to finish
                serviceTask.GetAwaiter().GetResult();
                return 0;
            }
            catch (Exception ex)
            {
                WriteDiagnosticMarkupLine($"[red]Service error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
            finally
            {
                DaemonProcessTracker.RecordExcelProcesses(pipeName, daemonIdentity,
                    service.ExcelProcesses.Select(process => new DaemonProcessTracker.ProcessIdentity(
                        process.ProcessId, process.StartedAtUtcFileTime)).ToArray());
                service.Dispose();
                service = null;
            }
        }
        finally
        {
            if (service != null)
            {
                DaemonProcessTracker.RecordExcelProcesses(pipeName, daemonIdentity,
                    service.ExcelProcesses.Select(process => new DaemonProcessTracker.ProcessIdentity(
                        process.ProcessId, process.StartedAtUtcFileTime)).ToArray());
                service.Dispose();
            }

            // Release the daemon mutex so a new daemon can start if needed.
            if (ownsLegacyDaemonMutex)
                legacyDaemonMutex!.ReleaseMutex();
            legacyDaemonMutex?.Dispose();
            if (ownsDaemonMutex)
                daemonMutex.ReleaseMutex();
            daemonMutex.Dispose();
        }
    }

}
