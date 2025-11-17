using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Infrastructure;

namespace Sbroenne.ExcelMcp.CLI.Tests.Helpers;

internal sealed class TestCliConsole : ICliConsole
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false
    };

    private readonly List<string> _infoMessages = new();
    private readonly List<string> _warningMessages = new();
    private readonly List<string> _errorMessages = new();
    private readonly List<string> _jsonMessages = new();

    public IReadOnlyList<string> InfoMessages => _infoMessages;
    public IReadOnlyList<string> WarningMessages => _warningMessages;
    public IReadOnlyList<string> ErrorMessages => _errorMessages;
    public IReadOnlyList<string> JsonMessages => _jsonMessages;

    public void WriteInfo(string message)
    {
        _infoMessages.Add(message);
    }

    public void WriteWarning(string message)
    {
        _warningMessages.Add(message);
    }

    public void WriteError(string message)
    {
        _errorMessages.Add(message);
    }

    public void WriteJson(object payload)
    {
        var json = JsonSerializer.Serialize(payload, JsonOptions);
        _jsonMessages.Add(json);
    }

    public string GetLastJson()
    {
        if (_jsonMessages.Count == 0)
        {
            throw new InvalidOperationException("No JSON messages were written to the console.");
        }

        return _jsonMessages[^1];
    }
}
