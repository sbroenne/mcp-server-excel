using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Commands;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console.Cli;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "ActionValidation")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CanonicalParameterOptionTests
{
    [Theory]
    [InlineData("--sheet", "Sheet1", "--range-address", "A1")]
    [InlineData("--sheet-name", "Sheet1", "--range", "A1")]
    public void RangeOptions_RemovedAliases_AreRejected(
        string sheetOption, string sheetName, string rangeOption, string rangeAddress)
    {
        var app = CreateApp();

        var exception = Assert.Throws<CommandParseException>(() =>
            app.Run(["get-values", sheetOption, sheetName, rangeOption, rangeAddress]));

        var removedOption = sheetOption == "--sheet" ? sheetOption : rangeOption;
        Assert.Contains($"'{removedOption.TrimStart('-')}'", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RangeOptions_CanonicalNames_AreAccepted()
    {
        var app = CreateApp();

        var exitCode = app.Run([
            "get-values", "--sheet-name", "Sheet1", "--range-address", "A1",
            "--session", "session-id", "--output", "result.json"]);

        Assert.Equal(0, exitCode);
    }

    [Theory]
    [InlineData("-s")]
    [InlineData("-o")]
    public void RangeOptions_ShortFlags_AreRejected(string option)
    {
        var app = CreateApp();

        var exception = Assert.Throws<CommandParseException>(() =>
            app.Run(["get-values", "--sheet-name", "Sheet1", "--range-address", "A1", option, "value"]));

        Assert.Contains($"'{option.TrimStart('-')}'", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("batch", "-i")]
    [InlineData("batch", "-s")]
    [InlineData("close", "-s")]
    public void HandwrittenOptions_ShortFlags_AreRejected(string command, string option)
    {
        var app = CreateHandwrittenApp();

        var exception = Assert.Throws<CommandParseException>(() => app.Run([command, option, "value"]));

        Assert.Contains($"'{option.TrimStart('-')}'", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("batch", "--input")]
    [InlineData("batch", "--session")]
    [InlineData("close", "--session")]
    public void HandwrittenOptions_CanonicalNames_AreAccepted(string command, string option)
    {
        Assert.Equal(0, CreateHandwrittenApp().Run([command, option, "value"]));
    }

    [Fact]
    public async Task GlobalOptions_RemovedQuietShortFlag_IsRejected()
    {
        var result = await CliProcessHelper.RunAsync("-q service status");

        Assert.NotEqual(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal("CommandParseException", output.RootElement.GetProperty("exceptionType").GetString());
        Assert.Contains("'q'", output.RootElement.GetProperty("error").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task GlobalOptions_CanonicalQuiet_IsAccepted()
    {
        var result = await CliProcessHelper.RunAsync(
            "--quiet service status",
            environmentVariables: new Dictionary<string, string>
            {
                ["EXCELMCP_CLI_PIPE"] = $"excelmcp-quiet-option-{Guid.NewGuid():N}"
            });

        Assert.Equal(0, result.ExitCode);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.True(output.RootElement.GetProperty("success").GetBoolean());
    }

    [Fact]
    public async Task GlobalOptions_CanonicalVersion_IsAccepted()
    {
        var result = await CliProcessHelper.RunAsync("--version");

        Assert.Equal(0, result.ExitCode);
        Assert.False(string.IsNullOrWhiteSpace(result.Stdout));
    }

    [Fact]
    public async Task GlobalOptions_FrameworkVersionShortFlag_IsAccepted()
    {
        var result = await CliProcessHelper.RunAsync("-v");

        Assert.Equal(0, result.ExitCode);
        Assert.False(string.IsNullOrWhiteSpace(result.Stdout));
    }

    private static CommandApp CreateHandwrittenApp()
    {
        var app = new CommandApp();
        app.Configure(config =>
        {
            config.PropagateExceptions();
            config.Settings.StrictParsing = true;
            config.AddCommand<ParseBatchCommand>("batch");
            config.AddCommand<ParseCloseCommand>("close");
        });
        return app;
    }

    private static CommandApp<ParseRangeCommand> CreateApp()
    {
        var app = new CommandApp<ParseRangeCommand>();
        app.Configure(config =>
        {
            config.PropagateExceptions();
            config.Settings.StrictParsing = true;
        });
        return app;
    }

    public sealed class ParseRangeCommand : Command<ServiceRegistry.Range.CliSettings>
    {
        protected override int Execute(
            CommandContext context,
            ServiceRegistry.Range.CliSettings settings,
            CancellationToken cancellationToken)
        {
            Assert.Equal("Sheet1", settings.SheetName);
            Assert.Equal("A1", settings.RangeAddress);
            return 0;
        }
    }

    internal sealed class ParseBatchCommand : Command<BatchCommand.Settings>
    {
        public ParseBatchCommand() { }

        protected override int Execute(CommandContext context, BatchCommand.Settings settings, CancellationToken cancellationToken)
        {
            Assert.Equal("value", settings.InputFile ?? settings.SessionId);
            return 0;
        }
    }

    internal sealed class ParseCloseCommand : Command<SessionCloseCommand.Settings>
    {
        public ParseCloseCommand() { }

        protected override int Execute(CommandContext context, SessionCloseCommand.Settings settings, CancellationToken cancellationToken)
        {
            Assert.Equal("value", settings.SessionId);
            return 0;
        }
    }
}
