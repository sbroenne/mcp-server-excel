using System.Text.Json;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure;

internal interface ICliConsole
{
    void WriteInfo(string message);
    void WriteWarning(string message);
    void WriteError(string message);
    void WriteJson(object payload);
}

internal sealed class SpectreCliConsole : ICliConsole
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public void WriteInfo(string message)
    {
        AnsiConsole.MarkupLine($"[green]{message.EscapeMarkup()}[/]");
    }

    public void WriteWarning(string message)
    {
        AnsiConsole.MarkupLine($"[yellow]{message.EscapeMarkup()}[/]");
    }

    public void WriteError(string message)
    {
        AnsiConsole.MarkupLine($"[red]{message.EscapeMarkup()}[/]");
    }

    public void WriteJson(object payload)
    {
        var json = JsonSerializer.Serialize(payload, JsonOptions);
        AnsiConsole.WriteLine(json);
    }
}
