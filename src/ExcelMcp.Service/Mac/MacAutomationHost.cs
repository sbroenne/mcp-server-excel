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
            var result = args[1] is "sheet.create" or "sheet.delete"
                ? MutateWorksheet(args[1], arguments)
                : MacOsaScriptRuntime.Execute(reader.ReadToEnd(), args[1], arguments);
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
        var sheetName = document.RootElement.GetProperty("sheetName").GetString();
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var mutation = command == "sheet.create"
            ? $"""
                tell workbook targetWorkbookIndex
                    set createdWorksheet to make new worksheet at end
                    set name of createdWorksheet to "{EscapeAppleScript(sheetName)}"
                end tell
                """
            : $"delete worksheet \"{EscapeAppleScript(sheetName)}\" of workbook targetWorkbookIndex";
        var script = $$"""
            set workbookPath to "{{EscapeAppleScript(filePath)}}"
            tell application "Microsoft Excel"
                set targetWorkbookIndex to 0
                repeat with workbookIndex from 1 to count of workbooks
                    if (full name of workbook workbookIndex as text) is workbookPath then
                        set targetWorkbookIndex to workbookIndex
                    end if
                end repeat
                if targetWorkbookIndex is 0 then error "Workbook is not open in this ExcelMcp session."
                {{mutation}}
                return "{\"success\":true,\"errorMessage\":\"\"}"
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
