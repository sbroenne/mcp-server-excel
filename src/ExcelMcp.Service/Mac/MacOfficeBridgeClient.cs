using System.Net;
using System.Net.Http.Headers;
using System.Net.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed record MacOfficeBridgeConfiguration(
    Uri Origin,
    string Token,
    string? CertificatePath,
    IReadOnlySet<string> EnabledActions)
{
    public static MacOfficeBridgeConfiguration LoadDefault()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var path = Path.Combine(
            home,
            "Library",
            "Application Support",
            "ExcelMcp",
            "officejs",
            "bridge.json");
        if (!File.Exists(path))
        {
            throw new MacOfficeBridgeException(
                "OfficeAddInUnavailable",
                "The optional Office.js bridge is not installed. Follow docs/MACOS-OFFICEJS.md, " +
                "start the bridge, and activate its task pane for this exact workbook.");
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var root = document.RootElement;
        var enabledActions = root.TryGetProperty("enabledActions", out var enabled)
            ? enabled.EnumerateArray()
                .Select(item => item.GetString())
                .OfType<string>()
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToHashSet(StringComparer.Ordinal)
            : new HashSet<string>(StringComparer.Ordinal) { "bridge.health" };
        return new MacOfficeBridgeConfiguration(
            new Uri(root.GetProperty("origin").GetString()!, UriKind.Absolute),
            root.GetProperty("token").GetString()!,
            root.GetProperty("certificatePath").GetString(),
            enabledActions);
    }

    public static bool IsActionEnabled(string action)
    {
        try
        {
            return LoadDefault().EnabledActions.Contains(action);
        }
        catch (Exception ex) when (ex is MacOfficeBridgeException
                                   or IOException
                                   or JsonException
                                   or InvalidOperationException
                                   or UriFormatException)
        {
            return false;
        }
    }
}

internal sealed class MacOfficeBridgeClient : IDisposable
{
    private readonly HttpClient _client;
    private readonly MacOfficeBridgeConfiguration _configuration;
    private readonly TimeSpan _pollInterval;
    private readonly bool _ownsClient;

    public MacOfficeBridgeClient(
        HttpClient client,
        MacOfficeBridgeConfiguration configuration,
        TimeSpan pollInterval)
    {
        _client = client;
        _configuration = configuration;
        _pollInterval = pollInterval;
    }

    private MacOfficeBridgeClient(MacOfficeBridgeConfiguration configuration)
    {
        _configuration = configuration;
        _pollInterval = TimeSpan.FromMilliseconds(50);
        _client = new HttpClient(CreatePinnedHandler(configuration))
        {
            BaseAddress = configuration.Origin
        };
        _ownsClient = true;
    }

    public static MacOfficeBridgeClient CreateDefault() =>
        new(MacOfficeBridgeConfiguration.LoadDefault());

    public async Task<JsonElement> InvokeAsync(
        string sessionId,
        string filePath,
        string action,
        JsonObject payload,
        TimeSpan timeout,
        bool mutation)
    {
        var workbookUrl = new Uri(Path.GetFullPath(filePath)).AbsoluteUri;
        using var deadline = new CancellationTokenSource(timeout);
        string? requestId = null;
        var wasDispatched = false;
        try
        {
            await PostAsync(
                "/v1/sessions",
                new JsonObject
                {
                    ["sessionId"] = sessionId,
                    ["workbookUrl"] = workbookUrl
                },
                deadline.Token);
            var created = await PostAsync(
                "/v1/requests",
                new JsonObject
                {
                    ["sessionId"] = sessionId,
                    ["workbookUrl"] = workbookUrl,
                    ["action"] = action,
                    ["payload"] = payload.DeepClone(),
                    ["timeoutMs"] = Math.Max(1, Math.Min(30_000, (int)Math.Ceiling(timeout.TotalMilliseconds)))
                },
                deadline.Token);
            requestId = created.GetProperty("requestId").GetString();

            while (true)
            {
                var status = await PostAsync(
                    "/v1/requests/status",
                    RequestIdentity(sessionId, workbookUrl, requestId!),
                    deadline.Token);
                var state = status.GetProperty("status").GetString();
                wasDispatched |= string.Equals(state, "active", StringComparison.Ordinal)
                    || status.TryGetProperty("dispatched", out var dispatched)
                        && dispatched.GetBoolean();
                switch (state)
                {
                    case "completed":
                        return status.GetProperty("result").GetProperty("value").Clone();
                    case "failed":
                        throw new MacOfficeBridgeException(
                            "OfficeJs",
                            status.GetProperty("result").GetProperty("errorMessage").GetString()
                                ?? "Office.js action failed.");
                    case "expired":
                    case "cancelled":
                        throw CreateTimeout(action, mutation && wasDispatched);
                }

                await Task.Delay(_pollInterval, deadline.Token);
            }
        }
        catch (OperationCanceledException) when (deadline.IsCancellationRequested)
        {
            if (requestId is not null)
            {
                wasDispatched |= await TryCancelAsync(sessionId, workbookUrl, requestId);
            }
            throw CreateTimeout(action, mutation && wasDispatched);
        }
    }

    public async Task TryUnregisterAsync(string sessionId, string filePath)
    {
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await PostAsync(
                "/v1/sessions/close",
                new JsonObject
                {
                    ["sessionId"] = sessionId,
                    ["workbookUrl"] = new Uri(Path.GetFullPath(filePath)).AbsoluteUri
                },
                timeout.Token);
        }
        catch (Exception ex) when (ex is HttpRequestException
                                   or OperationCanceledException
                                   or JsonException
                                   or MacOfficeBridgeException)
        {
            // Native session close remains authoritative when the optional broker is absent.
        }
    }

    private async Task<JsonElement> PostAsync(
        string path,
        JsonObject body,
        CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, new Uri(_configuration.Origin, path));
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _configuration.Token);
        request.Headers.TryAddWithoutValidation("Origin", _configuration.Origin.GetLeftPart(UriPartial.Authority));
        request.Content = new StringContent(
            body.ToJsonString(ServiceProtocol.JsonOptions),
            Encoding.UTF8,
            "application/json");
        using var response = await _client.SendAsync(request, cancellationToken);
        var text = await response.Content.ReadAsStringAsync(cancellationToken);
        using var document = JsonDocument.Parse(string.IsNullOrWhiteSpace(text) ? "{}" : text);
        if (!response.IsSuccessStatusCode)
        {
            var message = document.RootElement.TryGetProperty("errorMessage", out var error)
                ? error.GetString() ?? "Office.js bridge request failed."
                : $"Office.js bridge request failed ({(int)response.StatusCode}).";
            throw new MacOfficeBridgeException(ClassifyError(response.StatusCode, message), message);
        }
        return document.RootElement.Clone();
    }

    private async Task<bool> TryCancelAsync(string sessionId, string workbookUrl, string requestId)
    {
        try
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(1));
            var result = await PostAsync(
                "/v1/requests/cancel",
                RequestIdentity(sessionId, workbookUrl, requestId),
                timeout.Token);
            return result.TryGetProperty("dispatched", out var dispatched)
                && dispatched.GetBoolean();
        }
        catch (Exception ex) when (ex is HttpRequestException
                                   or OperationCanceledException
                                   or JsonException
                                   or MacOfficeBridgeException)
        {
            // Timeout semantics already treat an active mutation as uncertain.
            return false;
        }
    }

    private static JsonObject RequestIdentity(
        string sessionId,
        string workbookUrl,
        string requestId) =>
        new()
        {
            ["sessionId"] = sessionId,
            ["workbookUrl"] = workbookUrl,
            ["requestId"] = requestId
        };

    private static MacOfficeBridgeTimeoutException CreateTimeout(string action, bool uncertain) =>
        uncertain
            ? new MacOfficeMutationUncertainException(
                $"Office.js mutation '{action}' timed out after dispatch and may still complete in Excel. " +
                "The workbook session is no longer safe for more operations.")
            : new MacOfficeBridgeTimeoutException(
                $"Office.js action '{action}' did not complete before its deadline.");

    private static string ClassifyError(HttpStatusCode statusCode, string message)
    {
        if (message.Contains("bound", StringComparison.OrdinalIgnoreCase)
            || message.Contains("exact workbook", StringComparison.OrdinalIgnoreCase))
        {
            return "WorkbookBinding";
        }
        if (message.Contains("not active", StringComparison.OrdinalIgnoreCase)
            || message.Contains("not enabled", StringComparison.OrdinalIgnoreCase)
            || statusCode == HttpStatusCode.NotFound)
        {
            return "OfficeAddInUnavailable";
        }
        return statusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden
            ? "Authentication"
            : "InvalidOperation";
    }

    private static HttpClientHandler CreatePinnedHandler(MacOfficeBridgeConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(configuration.CertificatePath))
        {
            throw new MacOfficeBridgeException(
                "OfficeAddInUnavailable",
                "The Office.js bridge certificate path is missing.");
        }

        var pinned = X509Certificate2.CreateFromPemFile(configuration.CertificatePath);
        var pinnedHash = SHA256.HashData(pinned.RawData);
        return new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (_, certificate, _, errors) =>
            {
                if (certificate is null
                    || errors.HasFlag(SslPolicyErrors.RemoteCertificateNameMismatch))
                {
                    return false;
                }
                return CryptographicOperations.FixedTimeEquals(
                    pinnedHash,
                    SHA256.HashData(certificate.GetRawCertData()));
            }
        };
    }

    public void Dispose()
    {
        if (_ownsClient)
        {
            _client.Dispose();
        }
    }
}

internal class MacOfficeBridgeException(string errorCategory, string message)
    : InvalidOperationException(message)
{
    public string ErrorCategory { get; } = errorCategory;
}

internal class MacOfficeBridgeTimeoutException(string message)
    : TimeoutException(message);

internal sealed class MacOfficeMutationUncertainException(string message)
    : MacOfficeBridgeTimeoutException(message);
