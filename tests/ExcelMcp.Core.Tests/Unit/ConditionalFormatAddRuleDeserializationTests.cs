using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Generated;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "ConditionalFormat")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class ConditionalFormatAddRuleDeserializationTests
{
    [Fact]
    public void DispatchToCore_AddRule_WithBoolAndIntArgs_DeserializesAndPassesTypedValues()
    {
        var commands = new CapturingConditionalFormattingCommands();
        var argsJson = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1:A10",
            ruleType = "top10",
            rank = 7,
            top10Percent = true,
            fontBold = true,
            fontItalic = false,
            dataBarShowValue = true,
            iconSetReverse = true,
            iconSetShowIconOnly = false
        });

        var ex = Record.Exception(() => ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.AddRule,
            null!,
            argsJson));

        Assert.Null(ex);
        Assert.True(commands.AddRuleCalled);
        Assert.Equal(7, commands.Rank);
        Assert.True(commands.Top10Percent);
        Assert.True(commands.FontBold);
        Assert.False(commands.FontItalic);
        Assert.True(commands.DataBarShowValue);
        Assert.True(commands.IconSetReverse);
        Assert.False(commands.IconSetShowIconOnly);
    }

    [Fact]
    public void DispatchToCore_AddRule_WithEmptySheetName_UsesActiveSheet()
    {
        var commands = new CapturingConditionalFormattingCommands();
        var argsJson = JsonSerializer.Serialize(new
        {
            sheetName = "",
            rangeAddress = "A1:A10",
            ruleType = "expression",
            formula1 = "=A1>0"
        });

        ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.AddRule,
            null!,
            argsJson);

        Assert.True(commands.AddRuleCalled);
        Assert.Equal("", commands.SheetName);
    }

    [Fact]
    public void DispatchToCore_ClearRules_WithEmptySheetName_UsesActiveSheet()
    {
        var commands = new CapturingConditionalFormattingCommands();
        var argsJson = JsonSerializer.Serialize(new
        {
            sheetName = "",
            rangeAddress = "A1:A10"
        });

        ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.ClearRules,
            null!,
            argsJson);

        Assert.True(commands.ClearRulesCalled);
        Assert.Equal("", commands.SheetName);
    }

    [Fact]
    public void DispatchToCore_ListRules_WithEmptySheetName_UsesActiveSheet()
    {
        var commands = new CapturingConditionalFormattingCommands();
        var argsJson = JsonSerializer.Serialize(new
        {
            sheetName = "",
            rangeAddress = "A1:A10"
        });

        ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.ListRules,
            null!,
            argsJson);

        Assert.True(commands.ListRulesCalled);
        Assert.Equal("", commands.SheetName);
    }

    [Fact]
    public void DispatchToCore_ListWorksheetRules_WithEmptySheetName_UsesActiveSheet()
    {
        var commands = new CapturingConditionalFormattingCommands();
        var argsJson = JsonSerializer.Serialize(new { sheetName = "" });

        ServiceRegistry.ConditionalFormat.DispatchToCore(
            commands,
            ConditionalFormatAction.ListWorksheetRules,
            null!,
            argsJson);

        Assert.True(commands.ListWorksheetRulesCalled);
        Assert.Equal("", commands.SheetName);
    }

    [Fact]
    public void DispatchToCore_ListWorksheetRules_WithOmittedSheetName_IsRejected()
    {
        var commands = new CapturingConditionalFormattingCommands();

        var exception = Assert.Throws<ArgumentException>(() =>
            ServiceRegistry.ConditionalFormat.DispatchToCore(
                commands,
                ConditionalFormatAction.ListWorksheetRules,
                null!,
                "{}"));

        Assert.Contains("sheetName", exception.Message, StringComparison.Ordinal);
        Assert.False(commands.ListWorksheetRulesCalled);
    }

    private sealed class CapturingConditionalFormattingCommands : IConditionalFormattingCommands
    {
        public bool AddRuleCalled { get; private set; }
        public bool ClearRulesCalled { get; private set; }
        public bool ListRulesCalled { get; private set; }
        public bool ListWorksheetRulesCalled { get; private set; }
        public string? SheetName { get; private set; }
        public bool? FontBold { get; private set; }
        public bool? FontItalic { get; private set; }
        public bool? DataBarShowValue { get; private set; }
        public bool? IconSetReverse { get; private set; }
        public bool? IconSetShowIconOnly { get; private set; }
        public int? Rank { get; private set; }
        public bool? Top10Percent { get; private set; }

        public OperationResult AddRule(
            IExcelBatch batch,
            string sheetName,
            string rangeAddress,
            string ruleType,
            string? operatorType,
            string? formula1,
            string? formula2,
            string? interiorColor = null,
            string? interiorPattern = null,
            string? fontColor = null,
            bool? fontBold = null,
            bool? fontItalic = null,
            string? borderStyle = null,
            string? borderColor = null,
            string? colorScaleMinType = null,
            string? colorScaleMinValue = null,
            string? colorScaleMinColor = null,
            string? colorScaleMidType = null,
            string? colorScaleMidValue = null,
            string? colorScaleMidColor = null,
            string? colorScaleMaxType = null,
            string? colorScaleMaxValue = null,
            string? colorScaleMaxColor = null,
            string? dataBarColor = null,
            string? dataBarNegativeColor = null,
            string? dataBarDirection = null,
            bool? dataBarShowValue = null,
            string? dataBarMinType = null,
            string? dataBarMinValue = null,
            string? dataBarMaxType = null,
            string? dataBarMaxValue = null,
            string? iconSetId = null,
            bool? iconSetReverse = null,
            bool? iconSetShowIconOnly = null,
            string? iconThreshold1Type = null,
            string? iconThreshold1Value = null,
            string? iconThreshold2Type = null,
            string? iconThreshold2Value = null,
            string? iconThreshold3Type = null,
            string? iconThreshold3Value = null,
            string? iconThreshold4Type = null,
            string? iconThreshold4Value = null,
            int? rank = null,
            bool? top10Percent = null,
            string? topBottom = null,
            string? aboveBelow = null,
            string? datePeriod = null)
        {
            AddRuleCalled = true;
            SheetName = sheetName;
            FontBold = fontBold;
            FontItalic = fontItalic;
            DataBarShowValue = dataBarShowValue;
            IconSetReverse = iconSetReverse;
            IconSetShowIconOnly = iconSetShowIconOnly;
            Rank = rank;
            Top10Percent = top10Percent;
            return new OperationResult { Success = true };
        }

        public OperationResult ClearRules(IExcelBatch batch, string sheetName, string rangeAddress)
        {
            ClearRulesCalled = true;
            SheetName = sheetName;
            return new OperationResult { Success = true };
        }

        public ConditionalFormatListResult ListRules(IExcelBatch batch, string sheetName, string rangeAddress)
        {
            ListRulesCalled = true;
            SheetName = sheetName;
            return new ConditionalFormatListResult
            {
                Success = true,
                SheetName = sheetName,
                RangeAddress = rangeAddress
            };
        }

        public ConditionalFormatListResult ListWorksheetRules(IExcelBatch batch, string sheetName)
        {
            ListWorksheetRulesCalled = true;
            SheetName = sheetName;
            return new ConditionalFormatListResult { Success = true, SheetName = sheetName };
        }
    }
}
