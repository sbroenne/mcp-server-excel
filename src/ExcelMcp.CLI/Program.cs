using Microsoft.Extensions.DependencyInjection;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Commands.Chart;
using Sbroenne.ExcelMcp.CLI.Commands.ConditionalFormatting;
using Sbroenne.ExcelMcp.CLI.Commands.Connection;
using Sbroenne.ExcelMcp.CLI.Commands.DataModel;
using Sbroenne.ExcelMcp.CLI.Commands.File;
using Sbroenne.ExcelMcp.CLI.Commands.NamedRange;
using Sbroenne.ExcelMcp.CLI.Commands.PivotTable;
using Sbroenne.ExcelMcp.CLI.Commands.PowerQuery;
using Sbroenne.ExcelMcp.CLI.Commands.Range;
using Sbroenne.ExcelMcp.CLI.Commands.Session;
using Sbroenne.ExcelMcp.CLI.Commands.Sheet;
using Sbroenne.ExcelMcp.CLI.Commands.Table;
using Sbroenne.ExcelMcp.CLI.Commands.Vba;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI;

internal sealed class Program
{
    private static readonly string[] VersionFlags = ["--version", "-v"];

    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        if (args.Length == 0)
        {
            RenderHeader();
            AnsiConsole.MarkupLine("[dim]No command supplied. Use [green]--help[/] for usage examples.[/]");
            return 0;
        }

        if (args.Any(arg => VersionFlags.Contains(arg, StringComparer.OrdinalIgnoreCase)))
        {
            VersionReporter.WriteVersion();
            return 0;
        }

        RenderHeader();

        var services = new ServiceCollection();
        ConfigureServices(services);

        var registrar = new TypeRegistrar(services);
        var app = new CommandApp(registrar);

        app.Configure(config =>
        {
            config.SetApplicationName("excelcli");
            config.SetExceptionHandler((ex, _) =>
            {
                AnsiConsole.MarkupLine($"[red]Unhandled error:[/] {ex.Message.EscapeMarkup()}");
            });
            config.ValidateExamples();
            config.AddCommand<VersionCommand>("version")
                .WithDescription("Display excelcli version information.");

            config.AddBranch("session", branch =>
            {
                branch.SetDescription("Open, save, close, and list Excel sessions to reuse a single Excel process.");
                branch.AddCommand<SessionOpenCommand>("open");
                branch.AddCommand<SessionSaveCommand>("save");
                branch.AddCommand<SessionCloseCommand>("close");
                branch.AddCommand<SessionListCommand>("list");
            });

            config.AddCommand<CreateEmptyFileCommand>("create-empty")
                .WithDescription("Create a new empty workbook on disk (use --overwrite to replace existing files).");
            config.AddCommand<PowerQueryCommand>("powerquery")
                .WithDescription("Manage Power Query M code: list, import/export, update, and refresh queries.");
            config.AddCommand<RangeCommand>("range")
                .WithDescription("Work with worksheet ranges for values, formulas, formatting, validation, and hyperlinks.");
            config.AddCommand<SheetCommand>("sheet")
                .WithDescription("Manage worksheet lifecycle, tab colors, and visibility within a session.");
            config.AddCommand<NamedRangeCommand>("namedrange")
                .WithDescription("Create, update, delete, and list named ranges/parameters.");
            config.AddCommand<ConditionalFormattingCommand>("conditionalformat")
                .WithDescription("Add or clear conditional formatting rules on ranges.");
            config.AddCommand<TableCommand>("table")
                .WithDescription("Automate Excel Tables: create, resize, filter, sort, and manage totals.");
            config.AddCommand<PivotTableCommand>("pivottable")
                .WithDescription("Create and configure PivotTables, fields, and refresh behavior.");
            config.AddCommand<ChartCommand>("chart")
                .WithDescription("Create and manage Excel charts (Regular and PivotCharts).");
            config.AddCommand<ConnectionCommand>("connection")
                .WithDescription("Inspect, refresh, and update workbook data connections (OLEDB/ODBC/Text/Web).");
            config.AddCommand<DataModelCommand>("datamodel")
                .WithDescription("Create DAX measures/relationships and inspect the Power Pivot Data Model.");
            config.AddCommand<VbaCommand>("vba")
                .WithDescription("List, export/import, update, and run VBA modules or macros.");
        });

        try
        {
            return app.Run(args);
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

    private static void ConfigureServices(IServiceCollection services)
    {
        services.AddSingleton<SessionService>();
        services.AddSingleton<ISessionService>(sp => sp.GetRequiredService<SessionService>());
        services.AddSingleton<ICliConsole, SpectreCliConsole>();
        services.AddSingleton<IFileCommands, FileCommands>();
        services.AddSingleton<IDataModelCommands, DataModelCommands>();
        services.AddSingleton<IPowerQueryCommands>(sp => new PowerQueryCommands(sp.GetRequiredService<IDataModelCommands>()));
        services.AddSingleton<IRangeCommands, RangeCommands>();
        services.AddSingleton<ISheetCommands, SheetCommands>();
        services.AddSingleton<INamedRangeCommands, NamedRangeCommands>();
        services.AddSingleton<ITableCommands, TableCommands>();
        services.AddSingleton<IPivotTableCommands, PivotTableCommands>();
        services.AddSingleton<IChartCommands, ChartCommands>();
        services.AddSingleton<IConditionalFormattingCommands, ConditionalFormattingCommands>();
        services.AddSingleton<IConnectionCommands, ConnectionCommands>();
        services.AddSingleton<IVbaCommands, VbaCommands>();
    }

    private static void RenderHeader()
    {
        AnsiConsole.Write(new FigletText("Excel CLI").Color(Color.Blue));
        AnsiConsole.MarkupLine("[dim]Excel automation powered by ExcelMcp Core[/]");
        AnsiConsole.MarkupLine("[yellow]Workflow:[/] [green]session open <file>[/] → run commands with [green]--session <id>[/] → [green]session save[/] (optional) → [green]session close[/].");
        AnsiConsole.MarkupLine("[dim]Most commands expect an active session so they can reuse the same Excel instance.[/]");
        AnsiConsole.WriteLine();
    }
}
