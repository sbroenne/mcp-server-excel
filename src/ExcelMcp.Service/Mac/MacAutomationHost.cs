using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;

namespace Sbroenne.ExcelMcp.Service.Mac;

public static class MacAutomationHost
{
    public const string Marker = "--excelmcp-mac-automation";
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    public static bool TryRun(string[] args, out int exitCode)
    {
        if (args.Length != 4 || !string.Equals(args[0], Marker, StringComparison.Ordinal))
        {
            exitCode = 0;
            return false;
        }

        try
        {
            RequireSelfParent(args[2]);
            var timeout = ReadTimeout(args[3]);
            var permission = MacAutomationAccess.Check();
            if (permission != 0)
            {
                Console.Out.Write(JsonSerializer.Serialize(new
                {
                    success = false,
                    errorCategory = "ComInterop",
                    errorMessage =
                        $"Mac Excel automation is not ready: {MacAutomationAccess.DescribeStatus(permission)} " +
                        $"(OSStatus {permission}). Open licensed Excel and grant Automation permission interactively. " +
                        "No permission prompt was requested."
                }, ServiceProtocol.JsonOptions));
                exitCode = 0;
                return true;
            }

            var arguments = Console.In.ReadToEnd();
            var result = args[1] switch
            {
                "session.prepare-open" or "session.open" or "session.close" or "session.is-open" =>
                    RunNativeSessionCommand(args[1], arguments, timeout),
                "sheet.list" => ListNativeWorksheets(arguments, timeout),
                "sheet.check-name-scope" => CheckNativeNameScope(arguments, timeout),
                "helper.check" => JsonSerializer.Serialize(
                    new { success = true, helper = MacNativeHelper.Check(timeout) }, ServiceProtocol.JsonOptions),
                "helper.build" => BuildNativeHelper(arguments, timeout),
                "sheet.create" => CreateNativeWorksheet(arguments, timeout),
                "sheet.rename" => RenameNativeWorksheet(arguments, timeout),
                "calculation.calculate" => CalculateNativeScope(arguments, timeout),
                "range.describe" or "range.read-data" or "range.set-formulas" => RunNativeRange(args[1], arguments, timeout),
                "sheet.delete" => DeleteWorksheet(arguments),
                "namedrange.create" or "namedrange.delete" => MutateNamedRange(args[1], arguments),
                _ => ExecuteBridge(args[1], arguments)
            };
            Console.Out.Write(result);
            exitCode = 0;
        }
        catch (Exception ex)
        {
            Console.Out.Write(SerializeFailure(ex));
            exitCode = 0;
        }

        return true;
    }

    internal static string SerializeFailure(Exception error) =>
        JsonSerializer.Serialize(new
        {
            success = false,
            errorCategory = error is MacExcelOperationException helperError ? helperError.ErrorCategory
                : error is TimeoutException ? "Timeout"
                : error is PlatformNotSupportedException ? "PlatformNotSupported" : "ComInterop",
            errorMessage = error.Message,
            exceptionType = (error as MacExcelOperationException)?.RemoteExceptionType,
            innerError = (error as MacExcelOperationException)?.RemoteInnerError
        }, ServiceProtocol.JsonOptions);

    internal static TimeSpan ReadTimeout(string ticks)
    {
        if (!long.TryParse(ticks, System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out var value)
            || value <= 0)
        {
            throw new InvalidOperationException("Mac automation timeout must be a positive tick count.");
        }

        var timeout = TimeSpan.FromTicks(value);
        MacAppleEvents.TimeoutTicks(timeout);
        return timeout;
    }

    private static string CalculateNativeScope(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var root = document.RootElement;
        var result = MacNativeCalculation.Calculate(root.GetProperty("filePath").GetString()!,
            root.GetProperty("scope").Deserialize<CalculationScope>(ServiceProtocol.JsonOptions),
            root.GetProperty("sheetName").GetString()!,
            root.TryGetProperty("rangeAddress", out var address) ? address.GetString() : null, timeout);
        return JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions);
    }

    private static string RunNativeRange(string command, string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var root = document.RootElement;
        var path = root.GetProperty("filePath").GetString()!;
        var sheet = root.GetProperty("sheetName").GetString()!;
        var address = root.GetProperty("rangeAddress").GetString()!;
        var style = root.TryGetProperty("referenceStyle", out var referenceStyle)
            ? referenceStyle.Deserialize<FormulaReferenceStyle>(ServiceProtocol.JsonOptions) : FormulaReferenceStyle.A1;
        object result = command switch
        {
            "range.describe" => MacNativeRange.Describe(path, sheet, address, root.GetProperty("forWrite").GetBoolean(), timeout),
            "range.read-data" => MacNativeRange.ReadData(path, sheet, address, style, timeout),
            "range.set-formulas" => MacNativeRange.SetFormulas(path, sheet, address,
                root.GetProperty("formulas").Deserialize<List<List<string>>>(ServiceProtocol.JsonOptions)!, style, timeout),
            _ => throw new InvalidOperationException($"Unknown native range command: {command}")
        };
        return JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions);
    }

    private static string BuildNativeHelper(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var path = document.RootElement.GetProperty("workbookPath").GetString();
        var output = document.RootElement.GetProperty("outputPath").GetString();
        var version = document.RootElement.GetProperty("helperVersion").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentException.ThrowIfNullOrWhiteSpace(output);
        ArgumentException.ThrowIfNullOrWhiteSpace(version);
        return JsonSerializer.Serialize(new { success = true, helper = MacNativeHelper.Build(path, output, version, timeout) },
            ServiceProtocol.JsonOptions);
    }

    private static string CheckNativeNameScope(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        MacNativeWorksheet.CheckNameScope(document.RootElement.GetProperty("filePath").GetString()!, timeout);
        return JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions);
    }

    private static string RunNativeSessionCommand(string command, string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        switch (command)
        {
            case "session.prepare-open":
                MacNativeWorkbook.PrepareOpen(filePath, timeout);
                break;
            case "session.open":
                MacNativeWorkbook.Attach(filePath, document.RootElement.GetProperty("show").GetBoolean(), timeout);
                break;
            case "session.close":
                MacNativeWorkbook.Close(filePath, document.RootElement.GetProperty("save").GetBoolean(), timeout);
                break;
            case "session.is-open":
                return JsonSerializer.Serialize(new
                {
                    success = true,
                    errorMessage = "",
                    open = MacNativeWorkbook.IsOpen(filePath, timeout)
                }, ServiceProtocol.JsonOptions);
            default:
                throw new InvalidOperationException($"Unknown native Mac session command: {command}");
        }
        return JsonSerializer.Serialize(new { success = true, errorMessage = "" }, ServiceProtocol.JsonOptions);
    }

    private static string ExecuteBridge(string command, string arguments)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException($"Embedded macOS bridge '{ResourceName}' was not found.");
        using var reader = new StreamReader(stream);
        return MacOsaScriptRuntime.Execute(reader.ReadToEnd(), command, arguments);
    }

    private static string ListNativeWorksheets(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        return JsonSerializer.Serialize(MacNativeWorksheet.List(filePath, timeout), ServiceProtocol.JsonOptions);
    }

    private static string CreateNativeWorksheet(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        var sheetName = document.RootElement.GetProperty("sheetName").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        return JsonSerializer.Serialize(MacNativeWorksheet.Create(filePath, sheetName, timeout), ServiceProtocol.JsonOptions);
    }

    private static string RenameNativeWorksheet(string arguments, TimeSpan timeout)
    {
        using var document = JsonDocument.Parse(arguments);
        var root = document.RootElement;
        var result = MacNativeWorksheet.Rename(root.GetProperty("filePath").GetString()!,
            root.GetProperty("oldName").GetString()!, root.GetProperty("newName").GetString()!, timeout);
        return JsonSerializer.Serialize(result, ServiceProtocol.JsonOptions);
    }

    private static void RequireSelfParent(string parentProcessId)
    {
        if (!int.TryParse(
                parentProcessId,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var expectedParentId)
            || expectedParentId <= 0)
        {
            throw new InvalidOperationException("Mac automation parent identity is invalid.");
        }

        var actualParentId = GetParentProcessId();
        if (actualParentId != expectedParentId)
        {
            throw new InvalidOperationException("Mac automation parent identity does not match.");
        }

        using var parent = Process.GetProcessById(actualParentId);
        var parentPath = parent.MainModule?.FileName;
        var currentPath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(parentPath)
            || string.IsNullOrWhiteSpace(currentPath)
            || !string.Equals(
                Path.GetFullPath(parentPath),
                Path.GetFullPath(currentPath),
                StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "Mac automation can only be invoked by its owning ExcelMcp process.");
        }
    }

    private static int GetParentProcessId() => getppid();

    [System.Runtime.InteropServices.DllImport("/usr/lib/libSystem.B.dylib")]
    private static extern int getppid();

    private static string DeleteWorksheet(string arguments)
    {
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var sheetName = document.RootElement.GetProperty("sheetName").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var mutation = $"delete worksheet \"{EscapeAppleScript(sheetName)}\" of workbook targetWorkbookIndex";
        return ExecuteWorkbookScript(filePath, mutation);
    }

    private static string MutateNamedRange(string command, string arguments)
    {
        var create = command == "namedrange.create";
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        var name = document.RootElement.GetProperty("name").GetString();
        var reference = create ? document.RootElement.GetProperty("reference").GetString() : null;
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (create) ArgumentException.ThrowIfNullOrWhiteSpace(reference);
        var rejected = JsonSerializer.Serialize(new
        {
            success = false,
            errorCategory = "InvalidOperation",
            errorMessage = create
                ? $"Named range '{name}' already exists."
                : $"Named range '{name}' not found."
        }, ServiceProtocol.JsonOptions);
        var scopeUnavailable = JsonSerializer.Serialize(new
        {
            success = false,
            errorCategory = "PlatformNotSupported",
            errorMessage = "Native named-range creation cannot safely create worksheet-scoped names or shadow an existing worksheet-local name. Existing names were not changed."
        }, ServiceProtocol.JsonOptions);
        if (create && name.Contains('!', StringComparison.Ordinal))
        {
            return scopeUnavailable;
        }
        var operation = create
            ? $"make new named item at end with properties {{name:\"{EscapeAppleScript(name)}\", references:\"{EscapeAppleScript(reference!)}\"}}"
            : "delete named item targetNameIndex";
        var mutation = $$"""
            tell workbook targetWorkbookIndex
                set targetNameIndex to 0
                set localNameCollision to false
                repeat with nameIndex from 1 to count of named items
                    if (name of named item nameIndex as text) is "{{EscapeAppleScript(name)}}" then
                        set targetNameIndex to nameIndex
                    end if
                    if (name of named item nameIndex as text) ends with "!{{EscapeAppleScript(name)}}" then
                        set localNameCollision to true
                    end if
                end repeat
                if targetNameIndex is {{(create ? "not " : "")}}0 then return "{{EscapeAppleScript(rejected)}}"
                {{(create ? $"if localNameCollision then return \"{EscapeAppleScript(scopeUnavailable)}\"" : "")}}
                {{operation}}
                set nameExists to false
                repeat with nameIndex from 1 to count of named items
                    if (name of named item nameIndex as text) is "{{EscapeAppleScript(name)}}" then
                        set nameExists to true
                    end if
                end repeat
                if nameExists is {{(create ? "false" : "true")}} then error "Excel did not retain the requested named range mutation."
            end tell
            """;
        var result = ExecuteWorkbookScript(filePath, mutation);
        using var response = JsonDocument.Parse(result);
        if (!response.RootElement.GetProperty("success").GetBoolean())
        {
            return result;
        }
        return JsonSerializer.Serialize(new
        {
            success = true,
            errorMessage = "",
            filePath
        }, ServiceProtocol.JsonOptions);
    }

    private static string ExecuteWorkbookScript(string filePath, string mutation)
    {
        var success = JsonSerializer.Serialize(new
        {
            success = true,
            errorMessage = "",
            filePath
        }, ServiceProtocol.JsonOptions);
        var script = $$"""
            set workbookPath to "{{EscapeAppleScript(filePath)}}"
            tell application "Microsoft Excel"
                set targetWorkbookIndex to 0
                repeat with workbookIndex from 1 to count of workbooks
                    try
                        if (full name of workbook workbookIndex as text) is workbookPath then
                            set targetWorkbookIndex to workbookIndex
                        end if
                    end try
                end repeat
                if targetWorkbookIndex is 0 then error "Workbook is not open in this ExcelMcp session."
                {{mutation}}
                return "{{EscapeAppleScript(success)}}"
            end tell
            """;
        return MacOsaScriptRuntime.ExecuteAppleScript(script);
    }

    private static string EscapeAppleScript(string value) =>
        value.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal);
}
