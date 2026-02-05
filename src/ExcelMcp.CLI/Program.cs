using System.Reflection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Service.Infrastructure;
using Sbroenne.ExcelMcp.Generated;
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

        // Handle service run command before Spectre.Console (internal, not documented)
        if (args.Length >= 2 && args[0] == "service" && args[1] == "run")
        {
            return await RunServiceAsync();
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
            return await HandleVersionAsync();
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
            config.AddCommand<SheetCommand>("sheet")
                .WithDescription(DescribeActions(
                    "Worksheet operations.",
                    ServiceRegistry.Sheet.ValidActions
                        .Concat(ServiceRegistry.SheetStyle.ValidActions)));

            // Range commands
            config.AddCommand<RangeCommand>("range")
                .WithDescription(DescribeActions(
                    "Range operations.",
                    ServiceRegistry.Range.ValidActions
                        .Concat(ServiceRegistry.RangeEdit.ValidActions)
                        .Concat(ServiceRegistry.RangeFormat.ValidActions)
                        .Concat(ServiceRegistry.RangeLink.ValidActions)));

            // Table commands
            config.AddCommand<TableCommand>("table")
                .WithDescription(DescribeActions(
                    "Table operations.",
                    ServiceRegistry.Table.ValidActions));

            // PowerQuery commands
            config.AddCommand<PowerQueryCommand>("powerquery")
                .WithDescription(DescribeActions(
                    "Power Query operations.",
                    ServiceRegistry.PowerQuery.ValidActions));

            // PivotTable commands
            config.AddCommand<PivotTableCommand>("pivottable")
                .WithDescription(DescribeActions(
                    "PivotTable operations.",
                    ServiceRegistry.PivotTable.ValidActions
                        .Concat(ServiceRegistry.PivotTableField.ValidActions)
                        .Concat(ServiceRegistry.PivotTableCalc.ValidActions)));

            // Chart commands
            config.AddCommand<ChartCommand>("chart")
                .WithDescription(DescribeActions(
                    "Chart operations.",
                    ServiceRegistry.Chart.ValidActions));

            // ChartConfig commands
            config.AddCommand<ChartConfigCommand>("chartconfig")
                .WithDescription(DescribeActions(
                    "Chart configuration.",
                    ServiceRegistry.ChartConfig.ValidActions));

            // Connection commands
            config.AddCommand<ConnectionCommand>("connection")
                .WithDescription(DescribeActions(
                    "Connection operations.",
                    ServiceRegistry.Connection.ValidActions));

            // Calculation mode commands
            config.AddCommand<CalculationModeCommand>("calculationmode")
                .WithDescription(DescribeActions(
                    "Calculation mode operations.",
                    ServiceRegistry.Calculation.ValidActions));

            // NamedRange commands
            config.AddCommand<NamedRangeCommand>("namedrange")
                .WithDescription(DescribeActions(
                    "Named range operations.",
                    ServiceRegistry.NamedRange.ValidActions));

            // ConditionalFormat commands
            config.AddCommand<ConditionalFormatCommand>("conditionalformat")
                .WithDescription(DescribeActions(
                    "Conditional formatting.",
                    ServiceRegistry.ConditionalFormat.ValidActions));

            // VBA commands
            config.AddCommand<VbaCommand>("vba")
                .WithDescription(DescribeActions(
                    "VBA operations.",
                    ServiceRegistry.Vba.ValidActions));

            // DataModel commands
            config.AddCommand<DataModelCommand>("datamodel")
                .WithDescription(DescribeActions(
                    "Data Model operations.",
                    ServiceRegistry.DataModel.ValidActions));

            // DataModel relationship commands
            config.AddCommand<DataModelRelCommand>("datamodelrel")
                .WithDescription(DescribeActions(
                    "Data Model relationship operations.",
                    ServiceRegistry.DataModelRel.ValidActions));

            // Slicer commands
            config.AddCommand<SlicerCommand>("slicer")
                .WithDescription(DescribeActions(
                    "Slicer operations.",
                    ServiceRegistry.Slicer.ValidActions));
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

    private static async Task<int> RunServiceAsync()
    {
        using var service = new ExcelMcpService();

        // Handle Ctrl+C and process termination gracefully
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true; // Prevent immediate termination
            service.RequestShutdown();
        };

        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            service.RequestShutdown();
        };

        try
        {
            await service.RunAsync();
            return 0;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("already running"))
        {
            Console.Error.WriteLine("Service is already running.");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Service error: {ex.Message}");
            return 1;
        }
    }

    private static void RenderHeader()
    {
        AnsiConsole.Write(new FigletText("Excel CLI").Color(Spectre.Console.Color.Blue));
        AnsiConsole.MarkupLine("[dim]Excel automation powered by ExcelMcp Core[/]");
        AnsiConsole.MarkupLine("[yellow]Workflow:[/] [green]session open <file>[/] → run commands with [green]--session <id>[/] → [green]session close --save[/].");
        AnsiConsole.MarkupLine("[dim]A background service manages sessions for performance.[/]");
        AnsiConsole.WriteLine();
    }

    private static string DescribeActions(string baseDescription, IEnumerable<string> actions)
    {
        var actionList = string.Join(", ", actions);
        return $"{baseDescription} Actions: {actionList}.";
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
            AnsiConsole.MarkupLine($"[cyan]Run:[/] [white]dotnet tool update --global Sbroenne.ExcelMcp.McpServer[/]");
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
}


