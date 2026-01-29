using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Daemon;
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

        // Handle daemon run command before Spectre.Console
        if (args.Length >= 2 && args[0] == "daemon" && args[1] == "run")
        {
            return await RunDaemonAsync();
        }

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
            VersionReporter.WriteVersion();
            return 0;
        }

        if (showBanner) RenderHeader();

        var app = new CommandApp();

        app.Configure(config =>
        {
            config.SetApplicationName("excelcli");
            config.SetExceptionHandler((ex, _) =>
            {
                AnsiConsole.MarkupLine($"[red]Unhandled error:[/] {ex.Message.EscapeMarkup()}");
            });

            config.AddCommand<VersionCommand>("version")
                .WithDescription("Display excelcli version. Use --check to check for updates.");

            // Daemon commands
            config.AddBranch("daemon", branch =>
            {
                branch.SetDescription("Daemon management. The daemon holds Excel sessions across CLI invocations.");
                branch.AddCommand<DaemonStartCommand>("start")
                    .WithDescription("Start the daemon in the background.");
                branch.AddCommand<DaemonStopCommand>("stop")
                    .WithDescription("Stop the daemon.");
                branch.AddCommand<DaemonStatusCommand>("status")
                    .WithDescription("Show daemon status and active sessions.");
            });

            // Session commands
            config.AddBranch("session", branch =>
            {
                branch.SetDescription("Session management. WORKFLOW: open -> use sessionId -> close (--save to persist).");
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
            config.AddCommand<SheetCommand>("sheet")
                .WithDescription("Worksheet operations: list, create, rename, copy, delete, move.");

            // Range commands
            config.AddCommand<RangeCommand>("range")
                .WithDescription("Range operations: get/set values, formulas, number formats, copy, clear.");

            // Table commands
            config.AddCommand<TableCommand>("table")
                .WithDescription("Table operations: list, create, read, rename, delete, resize, style, append, get-data.");

            // PowerQuery commands
            config.AddCommand<PowerQueryCommand>("powerquery")
                .WithDescription("Power Query operations: list, view, create, update, refresh, delete, load-to.");

            // PivotTable commands
            config.AddCommand<PivotTableCommand>("pivottable")
                .WithDescription("PivotTable operations: list, read, create, delete, refresh.");

            // Chart commands
            config.AddCommand<ChartCommand>("chart")
                .WithDescription("Chart operations: list, read, create, delete, move, fit-to-range.");

            // ChartConfig commands
            config.AddCommand<ChartConfigCommand>("chartconfig")
                .WithDescription("Chart configuration: set-title, set-axis-title, add-series, set-style, data-labels, gridlines.");

            // Connection commands
            config.AddCommand<ConnectionCommand>("connection")
                .WithDescription("Connection operations: list, view, create, test, refresh, delete.");

            // NamedRange commands
            config.AddCommand<NamedRangeCommand>("namedrange")
                .WithDescription("Named range operations: list, read, write, create, update, delete.");

            // ConditionalFormat commands
            config.AddCommand<ConditionalFormatCommand>("conditionalformat")
                .WithDescription("Conditional formatting: add-rule, clear-rules.");

            // VBA commands
            config.AddCommand<VbaCommand>("vba")
                .WithDescription("VBA operations: list, view, import, update, run, delete.");

            // DataModel commands
            config.AddCommand<DataModelCommand>("datamodel")
                .WithDescription("Data Model operations: list-tables, list-measures, create/update/delete measures, evaluate DAX.");

            // Slicer commands
            config.AddCommand<SlicerCommand>("slicer")
                .WithDescription("Slicer operations: create, list, set-selection, delete (for PivotTables and Tables).");
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

    private static async Task<int> RunDaemonAsync()
    {
        using var daemon = new ExcelDaemon();
        try
        {
            await daemon.RunAsync();
            return 0;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("already running"))
        {
            Console.Error.WriteLine("Daemon is already running.");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Daemon error: {ex.Message}");
            return 1;
        }
    }

    private static void RenderHeader()
    {
        AnsiConsole.Write(new FigletText("Excel CLI").Color(Spectre.Console.Color.Blue));
        AnsiConsole.MarkupLine("[dim]Excel automation powered by ExcelMcp Core[/]");
        AnsiConsole.MarkupLine("[yellow]Workflow:[/] [green]session open <file>[/] → run commands with [green]--session <id>[/] → [green]session close --save[/].");
        AnsiConsole.MarkupLine("[dim]A background daemon manages sessions for performance.[/]");
        AnsiConsole.WriteLine();
    }
}
