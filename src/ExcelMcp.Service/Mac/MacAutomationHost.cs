using System.Reflection;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

public static class MacAutomationHost
{
    public const string Marker = "--excelmcp-mac-automation";
    private const string ResourceName = "Sbroenne.ExcelMcp.Service.Mac.MacExcelBridge.js";

    public static bool TryRun(string[] args, out int exitCode)
    {
        if (args.Length != 2 || !string.Equals(args[0], Marker, StringComparison.Ordinal))
        {
            exitCode = 0;
            return false;
        }

        try
        {
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

            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName)
                ?? throw new InvalidOperationException($"Embedded macOS bridge '{ResourceName}' was not found.");
            using var reader = new StreamReader(stream);
            var arguments = Console.In.ReadToEnd();
            var result = args[1] switch
            {
                "sheet.create" or "sheet.delete" => MutateWorksheet(args[1], arguments),
                "namedrange.create" or "namedrange.delete" => MutateNamedRange(args[1], arguments),
                "helper.dispatch" => DispatchHelper(arguments),
                _ => MacOsaScriptRuntime.Execute(reader.ReadToEnd(), args[1], arguments)
            };
            Console.Out.Write(result);
            exitCode = 0;
        }
        catch (Exception ex)
        {
            Console.Out.Write(JsonSerializer.Serialize(new
            {
                success = false,
                errorCategory = "ComInterop",
                errorMessage = ex.Message
            }, ServiceProtocol.JsonOptions));
            exitCode = 0;
        }

        return true;
    }

    private static string MutateWorksheet(string command, string arguments)
    {
        using var document = JsonDocument.Parse(arguments);
        var filePath = document.RootElement.GetProperty("filePath").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        string mutation;
        if (command == "sheet.create")
        {
            var sheetName = document.RootElement.GetProperty("sheetName").GetString();
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
            mutation = $"""
                tell workbook targetWorkbookIndex
                    set createdWorksheet to make new worksheet at end
                    set name of createdWorksheet to "{EscapeAppleScript(sheetName)}"
                end tell
                """;
        }
        else
        {
            var sheetName = document.RootElement.GetProperty("sheetName").GetString();
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
            mutation = $"delete worksheet \"{EscapeAppleScript(sheetName)}\" of workbook targetWorkbookIndex";
        }
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
                return "{\"success\":true,\"errorMessage\":\"\"}"
            end tell
            """;
        return MacOsaScriptRuntime.ExecuteAppleScript(script);
    }

    private static string DispatchHelper(string arguments)
    {
        using var document = JsonDocument.Parse(arguments);
        var requestJson = document.RootElement.GetProperty("requestJson").GetString();
        var helperPath = document.RootElement.GetProperty("helperPath").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(requestJson);
        ArgumentException.ThrowIfNullOrWhiteSpace(helperPath);
        _ = MacVbaHelperProtocol.ParseRequest(requestJson);
        var responseJson = MacOsaScriptRuntime.ExecuteAppleScript(
            CreateHelperDispatchScript(helperPath, requestJson));
        return JsonSerializer.Serialize(new
        {
            success = true,
            errorMessage = "",
            responseJson
        }, ServiceProtocol.JsonOptions);
    }

    internal static string CreateHelperDispatchScriptForTests(
        string helperPath,
        string requestJson) =>
        CreateHelperDispatchScript(helperPath, requestJson);

    private static string CreateHelperDispatchScript(string helperPath, string requestJson)
    {
        var canonicalHelperPath = Path.GetFullPath(helperPath);
        if (!string.Equals(
                Path.GetFileName(canonicalHelperPath),
                "ExcelMcpHelper.xlam",
                StringComparison.Ordinal))
        {
            throw new ArgumentException(
                "The configured helper path must end with 'ExcelMcpHelper.xlam'.",
                nameof(helperPath));
        }

        return $$"""
            set helperPath to "{{EscapeAppleScript(canonicalHelperPath)}}"
            set requestJson to "{{EscapeAppleScript(requestJson)}}"
            tell application "Microsoft Excel"
                set helperWorkbookIndex to 0
                set helperWorkbookCount to 0
                repeat with workbookIndex from 1 to count of workbooks
                    if (name of workbook workbookIndex as text) is "ExcelMcpHelper.xlam" then
                        set helperWorkbookCount to helperWorkbookCount + 1
                        if (full name of workbook workbookIndex as text) is not helperPath then
                            error "A different ExcelMcpHelper.xlam is open."
                        end if
                    end if
                    if (full name of workbook workbookIndex as text) is helperPath then
                        set helperWorkbookIndex to workbookIndex
                    end if
                end repeat
                if helperWorkbookCount is not 1 then error "Exactly one ExcelMcpHelper.xlam must be open."
                if helperWorkbookIndex is 0 then error "The configured ExcelMcp helper add-in is not open."
                return run VB macro "ExcelMcpHelper.xlam!ExcelMcpDispatch" arg1 requestJson
            end tell
            """;
    }

    private static string EscapeAppleScript(string value) =>
        value.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\"", "\\\"", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal);
}
