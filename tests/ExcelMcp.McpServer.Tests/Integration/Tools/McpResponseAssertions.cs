using System.Text.Json;
using ModelContextProtocol.Protocol;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

internal static class McpResponseAssertions
{
    internal static string ReadText(CallToolResult result, bool? expectedError = null)
    {
        Assert.NotNull(result);
        var text = Assert.Single(result.Content.OfType<TextContentBlock>()).Text;
        using var json = JsonDocument.Parse(text);
        var root = json.RootElement;
        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        if (expectedError.HasValue)
        {
            Assert.Equal(expectedError.Value, result.IsError == true);
        }

        var isFileValidation = root.TryGetProperty("canOpen", out _);
        if (isFileValidation)
        {
            foreach (var name in new[]
            {
                "success", "exists", "isValid", "canOpen", "isIrmProtected",
                "willOpenReadOnly", "requiresVisibleSession"
            })
            {
                Assert.True(root.TryGetProperty(name, out var property),
                    $"File validation omitted {name}: {text}");
                Assert.True(property.ValueKind is JsonValueKind.True or JsonValueKind.False,
                    $"Invalid file validation field {name}: {text}");
            }

            Assert.False(result.IsError == true, "File validation is an assessment, not a tool execution failure.");
            Assert.False(root.TryGetProperty("isError", out _),
                "File validation should not include a tool failure envelope.");
        }

        if (root.TryGetProperty("success", out var success) ||
            root.TryGetProperty("Success", out success))
        {
            Assert.True(success.ValueKind is JsonValueKind.True or JsonValueKind.False,
                $"Invalid success field: {text}");
            if (!isFileValidation)
            {
                Assert.Equal(success.ValueKind == JsonValueKind.False, result.IsError == true);
            }
        }

        if (result.StructuredContent.HasValue)
        {
            Assert.True(JsonElement.DeepEquals(root, result.StructuredContent.Value),
                "MCP text and structured results differ.");
        }

        return text;
    }

    internal static void AssertSuccess(string response, string operation)
    {
        JsonDocument json;
        try
        {
            json = JsonDocument.Parse(response);
        }
        catch (JsonException exception)
        {
            Assert.Fail($"{operation} returned invalid JSON: {exception.Message}\n{response}");
            throw;
        }

        using (json)
        {
            var root = json.RootElement;
            Assert.True(root.ValueKind == JsonValueKind.Object,
                $"{operation} returned a non-object result: {response}");
            Assert.True(root.TryGetProperty("success", out _) || root.TryGetProperty("Success", out _),
                $"{operation} omitted its success field: {response}");
            foreach (var name in new[] { "success", "Success" })
            {
                if (root.TryGetProperty(name, out var success))
                {
                    Assert.True(success.ValueKind == JsonValueKind.True,
                        $"{operation} did not succeed: {response}");
                }
            }

            foreach (var name in new[] { "error", "errorMessage", "ErrorMessage" })
            {
                if (root.TryGetProperty(name, out var error))
                {
                    Assert.True(error.ValueKind == JsonValueKind.Null ||
                        (error.ValueKind == JsonValueKind.String && error.GetString() == ""),
                        $"{operation} returned an error with success: {response}");
                }
            }

            if (root.TryGetProperty("isError", out var isError))
            {
                Assert.True(isError.ValueKind is JsonValueKind.False or JsonValueKind.Null,
                    $"{operation} returned isError with success: {response}");
            }
        }
    }
}
