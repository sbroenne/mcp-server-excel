using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNativeHelper
{
    internal static MacHelperInfo Check(TimeSpan timeout) =>
        MacHelperProtocol.ValidateResponse(
            Call($"'{MacHelperProtocol.FileName}'!{MacHelperProtocol.InfoFunction}", timeout),
            MacHelperProtocol.InfoPrimitives);

    internal static JsonNode? Call(string function, TimeSpan timeout, string? argument = null)
    {
        using var appleEvent = MacAppleEvents.Event(MacExcelDictionary.RunMacroClass, MacExcelDictionary.RunMacroId);
        using var name = MacAppleEvents.Text(function);
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), name);
        if (argument != null)
        {
            using var value = MacAppleEvents.Text(argument);
            MacAppleEvents.Put(appleEvent, MacExcelDictionary.MacroArgument1, value);
        }
        return MacAppleEvents.Send(appleEvent, timeout);
    }

    internal static MacHelperInfo Build(string workbookPath, string outputPath, string helperVersion, TimeSpan timeout)
    {
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(workbookPath, timeout);
        var qualifier = "'" + Path.GetFileName(workbookPath).Replace("'", "''", StringComparison.Ordinal) + "'!";
        var info = MacHelperProtocol.ValidateResponse(
            Call(qualifier + MacHelperProtocol.InfoFunction, MacAppleEvents.Remaining(timeout, started)),
            MacHelperProtocol.InfoPrimitives);
        if (!string.Equals(info.Version, helperVersion, StringComparison.Ordinal))
        {
            throw new MacExcelOperationException("HelperIncompatible",
                $"The bootstrap contains helper {info.Version}, but the requested source version is {helperVersion}. " +
                "Import the reviewed module for that source version before building; no build mutation was dispatched.");
        }
        var result = Call(qualifier + "ExcelMcpHelper_Build", MacAppleEvents.Remaining(timeout, started), outputPath);
        ValidateBuildResponse(result);
        if (!File.Exists(outputPath))
        {
            throw new MacExcelOperationException("HelperBuildFailed", "Excel did not persist the helper add-in at the requested output path.");
        }
        return info;
    }

    internal static void ValidateBuildResponse(JsonNode? result)
    {
        if (result is not JsonValue value || !value.TryGetValue<string>(out var json) || string.IsNullOrWhiteSpace(json))
        {
            throw new MacExcelOperationException("HelperBuildFailed", "Excel did not confirm a successful helper add-in build.");
        }
        try
        {
            using var document = JsonDocument.Parse(json);
            var response = document.RootElement;
            if (response.ValueKind != JsonValueKind.Object
                || response.EnumerateObject().GroupBy(property => property.Name, StringComparer.Ordinal).Any(group => group.Count() > 1)
                || !response.TryGetProperty("success", out var success)
                || success.ValueKind is not (JsonValueKind.True or JsonValueKind.False)
                || !response.TryGetProperty("errorMessage", out var errorMessage) || errorMessage.ValueKind != JsonValueKind.String)
            {
                throw new MacExcelOperationException("HelperBuildFailed", "Excel returned a malformed helper build confirmation.");
            }
            if (!success.GetBoolean() || !string.IsNullOrEmpty(errorMessage.GetString()))
            {
                var errorNumber = response.TryGetProperty("errorNumber", out var number) ? number.GetRawText() : "unreported";
                throw new MacExcelOperationException("HelperBuildFailed",
                    $"Excel helper build failed (VBA {errorNumber}): {errorMessage.GetString()}");
            }
        }
        catch (JsonException ex)
        {
            throw new MacExcelOperationException("HelperBuildFailed", "Excel returned invalid helper build confirmation JSON.", ex);
        }
    }
}
