using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacHelperInfo(string Version, IReadOnlyList<string> Primitives);

internal static partial class MacHelperProtocol
{
    internal const int SupportedMajor = 1;
    internal const string FileName = "ExcelMcpHelper.xlam";
    internal const string InfoFunction = "ExcelMcpHelper_Info";
    internal static readonly string[] InfoPrimitives = ["helper.info"];

    internal static MacHelperInfo ValidateResponse(JsonNode? response, IReadOnlyCollection<string> requiredPrimitives)
    {
        if (response is null)
        {
            return Validate(null, requiredPrimitives);
        }
        if (response is not JsonValue value || !value.TryGetValue<string>(out var json))
        {
            throw Malformed();
        }
        return Validate(json, requiredPrimitives);
    }

    internal static MacHelperInfo Validate(string? json, IReadOnlyCollection<string> requiredPrimitives)
    {
        if (string.IsNullOrWhiteSpace(json) || json == "null")
        {
            throw new MacExcelOperationException("HelperMissing",
                "Excel did not return the ExcelMcp helper handshake. Install and enable ExcelMcpHelper.xlam in Excel, " +
                "and approve its macros interactively. Native actions do not require the optional helper.");
        }
        try
        {
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object
                || root.EnumerateObject().GroupBy(property => property.Name, StringComparer.Ordinal).Any(group => group.Count() > 1)
                || !root.TryGetProperty("version", out var versionValue) || versionValue.ValueKind != JsonValueKind.String
                || !root.TryGetProperty("primitives", out var primitivesValue) || primitivesValue.ValueKind != JsonValueKind.Array)
            {
                throw Malformed();
            }
            var version = versionValue.GetString()!;
            var match = SemanticVersion().Match(version);
            if (!match.Success || !int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var major)
                || match.Groups[4].Value.Split('.', StringSplitOptions.RemoveEmptyEntries)
                    .Any(identifier => identifier.Length > 1 && identifier[0] == '0' && identifier.All(char.IsAsciiDigit)))
            {
                throw Malformed();
            }
            var primitives = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in primitivesValue.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.String || item.GetString() is not { Length: > 0 } primitive
                    || primitive.Any(char.IsWhiteSpace) || !primitives.Add(primitive))
                {
                    throw Malformed();
                }
            }
            if (major != SupportedMajor)
            {
                throw new MacExcelOperationException("HelperIncompatible",
                    $"ExcelMcp helper {version} is incompatible: this server requires helper major {SupportedMajor}. " +
                    "Server and helper release versions do not need to match.");
            }
            var missing = requiredPrimitives.Where(primitive => !primitives.Contains(primitive)).Order(StringComparer.Ordinal).ToArray();
            if (missing.Length != 0)
            {
                throw new MacExcelOperationException("HelperIncompatible",
                    $"ExcelMcp helper {version} is missing required primitives: {string.Join(", ", missing)}. " +
                    "Install a compatible helper that supplies those primitives; no workbook mutation was dispatched.");
            }
            return new MacHelperInfo(version, primitives.Order(StringComparer.Ordinal).ToArray());
        }
        catch (JsonException ex)
        {
            throw new MacExcelOperationException("HelperMalformed", "Excel returned invalid helper handshake JSON.", ex);
        }
    }

    private static MacExcelOperationException Malformed() =>
        new("HelperMalformed", "Excel returned a malformed helper version or primitive inventory. No helper mutation was dispatched.");

    [GeneratedRegex(@"\A(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?\z", RegexOptions.NonBacktracking)]
    private static partial Regex SemanticVersion();
}
