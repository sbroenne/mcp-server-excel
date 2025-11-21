using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console.Cli;
using IODirectory = System.IO.Directory;
using IOFile = System.IO.File;
using IOPath = System.IO.Path;

namespace Sbroenne.ExcelMcp.CLI.Commands.Vba;

internal sealed class VbaCommand : Command<VbaCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IVbaCommands _vbaCommands;
    private readonly ICliConsole _console;

    public VbaCommand(ISessionService sessionService, IVbaCommands vbaCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _vbaCommands = vbaCommands ?? throw new ArgumentNullException(nameof(vbaCommands));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            _console.WriteError("Session ID is required. Use 'session open' first.");
            return -1;
        }

        var action = settings.Action?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(action))
        {
            _console.WriteError("Action is required.");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        return action switch
        {
            "list" => WriteResult(_vbaCommands.List(batch)),
            "view" => ExecuteView(batch, settings),
            "export" => ExecuteExport(batch, settings),
            "import" => ExecuteImport(batch, settings),
            "update" => ExecuteUpdate(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "run" => ExecuteRun(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteView(IExcelBatch batch, Settings settings)
    {
        if (!TryGetModuleName(settings, out var moduleName))
        {
            return -1;
        }

        return WriteResult(_vbaCommands.View(batch, moduleName));
    }

    private int ExecuteExport(IExcelBatch batch, Settings settings)
    {
        if (!TryGetModuleName(settings, out var moduleName))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.OutputPath))
        {
            _console.WriteError("--output is required for export.");
            return -1;
        }

        var viewResult = _vbaCommands.View(batch, moduleName);
        if (!viewResult.Success)
        {
            return WriteResult(viewResult);
        }

        try
        {
            var outputPath = settings.OutputPath!;
            var directory = IOPath.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory))
            {
                IODirectory.CreateDirectory(directory);
            }

            IOFile.WriteAllText(outputPath, viewResult.Code ?? string.Empty);

            var exportResult = new OperationResult
            {
                Success = true,
                Action = "vba-export",
                FilePath = viewResult.FilePath
            };

            _console.WriteJson(exportResult);
            return 0;
        }
        catch (Exception ex)
        {
            var errorResult = new OperationResult
            {
                Success = false,
                Action = "vba-export",
                FilePath = viewResult.FilePath,
                ErrorMessage = $"Failed to export module '{moduleName}': {ex.Message}"
            };

            _console.WriteJson(errorResult);
            return -1;
        }
    }

    private int ExecuteImport(IExcelBatch batch, Settings settings)
    {
        if (!TryGetModuleName(settings, out var moduleName))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.CodeFile))
        {
            _console.WriteError("--code-file is required for import.");
            return -1;
        }

        if (!IOFile.Exists(settings.CodeFile))
        {
            _console.WriteError($"File not found: {settings.CodeFile}");
            return -1;
        }

        string vbaCode = IOFile.ReadAllText(settings.CodeFile);
        return WriteResult(_vbaCommands.Import(batch, moduleName, vbaCode));
    }

    private int ExecuteUpdate(IExcelBatch batch, Settings settings)
    {
        if (!TryGetModuleName(settings, out var moduleName))
        {
            return -1;
        }

        if (string.IsNullOrWhiteSpace(settings.CodeFile))
        {
            _console.WriteError("--code-file is required for update.");
            return -1;
        }

        if (!IOFile.Exists(settings.CodeFile))
        {
            _console.WriteError($"File not found: {settings.CodeFile}");
            return -1;
        }

        string vbaCode = IOFile.ReadAllText(settings.CodeFile);
        return WriteResult(_vbaCommands.Update(batch, moduleName, vbaCode));
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (!TryGetModuleName(settings, out var moduleName))
        {
            return -1;
        }

        return WriteResult(_vbaCommands.Delete(batch, moduleName));
    }

    private int ExecuteRun(IExcelBatch batch, Settings settings)
    {
        var procedureName = settings.ProcedureName?.Trim();
        if (string.IsNullOrWhiteSpace(procedureName))
        {
            _console.WriteError("--procedure is required for run.");
            return -1;
        }

        TimeSpan? timeout = settings.TimeoutSeconds.HasValue
            ? TimeSpan.FromSeconds(settings.TimeoutSeconds.Value)
            : null;

        var parameters = settings.Parameters ?? Array.Empty<string>();
        return WriteResult(_vbaCommands.Run(batch, procedureName, timeout, parameters));
    }

    private bool TryGetModuleName(Settings settings, out string moduleName)
    {
        moduleName = settings.ModuleName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(moduleName))
        {
            _console.WriteError("--module is required for this action.");
            return false;
        }

        return true;
    }

    private int WriteResult(ResultBase result)
    {
        _console.WriteJson(result);
        return result.Success ? 0 : -1;
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown vba action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--module <NAME>")]
        public string? ModuleName { get; init; }

        [CommandOption("--procedure <NAME>")]
        public string? ProcedureName { get; init; }

        [CommandOption("--code-file <PATH>")]
        public string? CodeFile { get; init; }

        [CommandOption("--output <PATH>")]
        public string? OutputPath { get; init; }

        [CommandOption("--timeout-seconds <SECONDS>")]
        public int? TimeoutSeconds { get; init; }

        [CommandOption("--parameter <VALUE>")]
        public string[] Parameters { get; init; } = Array.Empty<string>();
    }
}
