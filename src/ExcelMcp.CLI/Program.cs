using Microsoft.Extensions.DependencyInjection;
using Spectre.Console;
using Spectre.Console.Cli;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Commands.File;
using Sbroenne.ExcelMcp.CLI.Commands.PowerQuery;
using Sbroenne.ExcelMcp.CLI.Commands.NamedRange;
using Sbroenne.ExcelMcp.CLI.Commands.Sheet;
using Sbroenne.ExcelMcp.CLI.Commands.Range;
using Sbroenne.ExcelMcp.CLI.Commands.Table;
using Sbroenne.ExcelMcp.CLI.Commands.Session;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;

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
            config.AddCommand<VersionCommand>("version");

            config.AddBranch("session", branch =>
            {
                branch.SetDescription("Manage Excel session lifecycle");
                branch.AddCommand<SessionOpenCommand>("open");
                branch.AddCommand<SessionSaveCommand>("save");
                branch.AddCommand<SessionCloseCommand>("close");
                branch.AddCommand<SessionListCommand>("list");
            });

            config.AddCommand<CreateEmptyFileCommand>("create-empty");
            config.AddCommand<PowerQueryCommand>("powerquery");
            config.AddCommand<RangeCommand>("range");
            config.AddCommand<SheetCommand>("sheet");
            config.AddCommand<NamedRangeCommand>("namedrange");
            config.AddCommand<TableCommand>("table");
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
    }

    private static void RenderHeader()
    {
        AnsiConsole.Write(new FigletText("Excel CLI").Color(Color.Blue));
        AnsiConsole.MarkupLine("[dim]Excel automation powered by ExcelMcp Core[/]");
        AnsiConsole.WriteLine();
    }
}
