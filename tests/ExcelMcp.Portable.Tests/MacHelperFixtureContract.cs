using System.Text;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Portable.Tests;

internal sealed record MacHelperFixtureStep(
    string Action,
    IReadOnlyDictionary<string, object?> Arguments);

internal static class MacHelperFixtureContract
{
    public const int ProtocolVersion = 1;
    public const int MaximumUtf8Bytes = 262_144;
    public const string HelperVersion = "1.0.1";
    public const string QueryName = "ExcelMcpFixtureLiteral";
    public const string RenamedQueryName = "ExcelMcpFixtureLiteralRenamed";
    public const string QueryFormula =
        """let Source = #table({"Value"}, {{"fixture"}}) in Source""";
    public const string UpdatedQueryFormula =
        """let Source = #table({"Value"}, {{"updated"}}) in Source""";
    public const string ModuleName = "ExcelMcpFixtureModule";
    public const string ModuleSource =
        "Option Explicit\n\nPublic Function ExcelMcpFixtureValue() As String\n" +
        "    ExcelMcpFixtureValue = \"fixture\"\nEnd Function";
    public const string UpdatedModuleSource =
        "Option Explicit\n\nPublic Function ExcelMcpFixtureValue() As String\n" +
        "    ExcelMcpFixtureValue = \"updated\"\nEnd Function";

    public static IReadOnlyList<string> SupportedActions { get; } =
    [
        "helper.capabilities",
        "powerquery.list",
        "powerquery.view",
        "powerquery.create",
        "powerquery.update",
        "powerquery.rename",
        "powerquery.delete",
        "analysis.create-scenario",
        "analysis.show-scenario",
        "vba.list",
        "vba.view",
        "vba.import",
        "vba.update",
        "vba.delete"
    ];

    public static IReadOnlyList<MacHelperFixtureStep> PowerQueryLifecycle { get; } =
    [
        new("powerquery.create", new Dictionary<string, object?>
        {
            ["name"] = QueryName,
            ["formula"] = QueryFormula
        }),
        new("powerquery.list", new Dictionary<string, object?>()),
        new("powerquery.view", new Dictionary<string, object?> { ["name"] = QueryName }),
        new("powerquery.update", new Dictionary<string, object?>
        {
            ["name"] = QueryName,
            ["formula"] = UpdatedQueryFormula
        }),
        new("powerquery.rename", new Dictionary<string, object?>
        {
            ["name"] = QueryName,
            ["newName"] = RenamedQueryName
        }),
        new("powerquery.view", new Dictionary<string, object?> { ["name"] = RenamedQueryName }),
        new("powerquery.delete", new Dictionary<string, object?>
        {
            ["name"] = RenamedQueryName,
            ["deleteConnection"] = true
        }),
        new("powerquery.list", new Dictionary<string, object?>())
    ];

    public static IReadOnlyList<MacHelperFixtureStep> VbaLifecycle { get; } =
    [
        new("vba.import", new Dictionary<string, object?>
        {
            ["moduleName"] = ModuleName,
            ["source"] = ModuleSource
        }),
        new("vba.list", new Dictionary<string, object?>()),
        new("vba.view", new Dictionary<string, object?> { ["moduleName"] = ModuleName }),
        new("vba.update", new Dictionary<string, object?>
        {
            ["moduleName"] = ModuleName,
            ["source"] = UpdatedModuleSource
        }),
        new("vba.view", new Dictionary<string, object?> { ["moduleName"] = ModuleName }),
        new("vba.delete", new Dictionary<string, object?> { ["moduleName"] = ModuleName }),
        new("vba.list", new Dictionary<string, object?>())
    ];

    public static string CreateRequest(
        string workbookPath,
        string action,
        IReadOnlyDictionary<string, object?> arguments,
        string requestId)
    {
        if (requestId.Length != 32
            || requestId.Any(character => character is not (>= '0' and <= '9')
                and not (>= 'a' and <= 'f')))
        {
            throw new ArgumentException(
                "Helper request IDs must contain exactly 32 lowercase hexadecimal characters.",
                nameof(requestId));
        }

        var fullPath = Path.GetFullPath(workbookPath);
        var request = JsonSerializer.Serialize(new
        {
            version = ProtocolVersion,
            requestId,
            workbookPath = fullPath,
            action,
            arguments
        });
        if (Encoding.UTF8.GetByteCount(request) > MaximumUtf8Bytes)
        {
            throw new ArgumentException(
                $"Helper requests cannot exceed {MaximumUtf8Bytes} UTF-8 bytes.",
                nameof(arguments));
        }

        return request;
    }

    public static JsonElement ParseResponse(string response, string requestId)
    {
        if (Encoding.UTF8.GetByteCount(response) > MaximumUtf8Bytes)
        {
            throw new InvalidDataException(
                $"Helper responses cannot exceed {MaximumUtf8Bytes} UTF-8 bytes.");
        }

        using var document = JsonDocument.Parse(response);
        var root = document.RootElement;
        if (root.GetProperty("version").GetInt32() != ProtocolVersion
            || root.GetProperty("requestId").GetString() != requestId)
        {
            throw new InvalidDataException("Helper response identity does not match the request.");
        }

        var success = root.GetProperty("success").GetBoolean();
        var result = root.GetProperty("result");
        var error = root.GetProperty("error");
        if (success)
        {
            if (result.ValueKind != JsonValueKind.Object || error.ValueKind != JsonValueKind.Null)
            {
                throw new InvalidDataException(
                    "Successful helper responses require an object result and null error.");
            }
        }
        else
        {
            if (result.ValueKind != JsonValueKind.Null || error.ValueKind != JsonValueKind.Object)
            {
                throw new InvalidDataException(
                    "Failed helper responses require a null result and structured error.");
            }

            foreach (var property in new[] { "category", "code", "message" })
            {
                if (string.IsNullOrWhiteSpace(error.GetProperty(property).GetString()))
                {
                    throw new InvalidDataException(
                        $"Helper error property '{property}' cannot be empty.");
                }
            }
        }

        return root.Clone();
    }
}
