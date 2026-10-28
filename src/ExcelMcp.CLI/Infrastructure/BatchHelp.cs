namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal static class BatchHelp
{
    internal const string InputHint = "See 'excelcli batch --help' for input examples.";
    internal const string RangeWriteEntry =
        """{"command":"range.set-values","args":{"sheetName":"Sales","rangeAddress":"A1:B2","values":[["Item","Amount"],["Widget",42]]}}""";

    internal static void Write()
    {
        CliCommandRuntime.Current.Output.WriteLine(
            $$$"""

            When to use batch:
              Run a known sequence in one CLI call to avoid repeated launches.
              Commands run one at a time. Use individual commands when you must
              inspect a result before choosing the next step.

            Input:
              Supply a JSON array or one command per line (NDJSON), in a file or stdin.
              --input commands.json reads a file; --input - or omitted --input reads stdin.
              Each entry has:
                command    Required group.action, for example "range.set-values".
                args       Object required when the action needs arguments.
                sessionId  Optional session override for this entry, outside args.
              Argument names are case-sensitive: use sheetName and rangeAddress,
              not the CLI aliases sheet and range. Supply arrays and objects as
              JSON values, not strings containing JSON. For exact keys, run:
                excelcli range set-values --help
              Substitute the group and action you need.

            Sessions:
              --session supplies the default; an entry's sessionId overrides it.
              Without a default, successful session.open/create/attach captures
              sessionId for later entries. Closing that default session successfully
              clears it. Returned values other than sessionId are not substituted.

            Results and failures:
              stdout has one JSON result per executed or rejected entry.
              Invalid JSON or missing command fields produce a JSON error on stdout.
              Match each result to its input using the zero-based index.
              Check every success flag, the expected result count, and $LASTEXITCODE:
              0 means all succeeded; 1 means an error occurred.
              Progress is written separately to stderr, including in quiet mode.
              Use --stop-on-error for dependent steps (default: continue after errors).
              A batch is not all-or-nothing: earlier changes remain after a failure.
              Batch does not undo or save automatically. Excel AutoSave, if enabled,
              can still persist edits. Inspect partial results before
              deciding what to do next; do not blindly repeat writes.
              Save/close separately only when requested, after checking results
              and the workbook. Use session save --help and session close --help.

            PowerShell example (no workbook required):
            @'
            [
              {"command":"diag.ping"},
              {"command":"diag.echo","args":{"message":"hello"}}
            ]
            '@ | excelcli.exe batch --stop-on-error

            Excel example:
              Requires $sessionId from an open/attached workbook, an existing Sales
              sheet, and authorization to write the inspected empty A1:B2 cells.
            @'
            [
              {{{RangeWriteEntry}}},
              {"command":"range.get-values","args":{"sheetName":"Sales","rangeAddress":"A1:B2"}}
            ]
            '@ | excelcli.exe batch --session $sessionId --stop-on-error
            """);
    }
}
