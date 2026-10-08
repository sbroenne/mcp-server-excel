using System.Diagnostics;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNativeWorksheet
{
    internal static void CheckNameScope(string filePath, TimeSpan timeout)
    {
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        if (MacNativeRange.Count(workbook, MacExcelDictionary.ChartSheetClass, MacAppleEvents.Remaining(timeout, started)) != 0)
            throw new PlatformNotSupportedException(
                "macOS worksheet create/rename in workbooks containing chart sheets is not yet accepted; no naming mutation was attempted.");
    }

    internal static OperationResult Create(string filePath, string sheetName, TimeSpan timeout)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        using var activeSheet = MacAppleEvents.Property(workbook, MacExcelDictionary.ActiveSheet);
        using var location = MacAppleEvents.Record(MacAppleEvents.Code("insl"));
        using var before = MacAppleEvents.Create(MacAppleEvents.Code("enum"), BitConverter.GetBytes(MacAppleEvents.Code("befo")));
        MacAppleEvents.PutKey(location, MacAppleEvents.Code("kobj"), activeSheet);
        MacAppleEvents.PutKey(location, MacAppleEvents.Code("kpos"), before);
        using var worksheetClass = MacAppleEvents.Create(MacAppleEvents.Code("type"), BitConverter.GetBytes(MacExcelDictionary.WorksheetClass));
        using var create = MacAppleEvents.Event(MacExcelDictionary.MakeClass, MacExcelDictionary.MakeId);
        MacAppleEvents.Put(create, MacExcelDictionary.MakeClassParameter, worksheetClass);
        MacAppleEvents.Put(create, MacExcelDictionary.MakeLocationParameter, location);
        using var created = MacAppleEvents.SendSpecifier(create, MacAppleEvents.Remaining(timeout, started));
        var existingName = MacNativeRange.Read(created, MacExcelDictionary.Name, MacAppleEvents.Remaining(timeout, started))
            ?.GetValue<string>() ?? throw new InvalidDataException("Excel did not return the created worksheet name.");
        SetName(created, sheetName, existingName, isNewSheet: true, MacAppleEvents.Remaining(timeout, started));
        return new OperationResult { Success = true, FilePath = filePath };
    }

    internal static OperationResult Rename(string filePath, string oldName, string newName, TimeSpan timeout)
    {
        WorksheetCommandValidation.ValidateNewSheetName(newName);
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        using var selector = MacAppleEvents.Text(oldName);
        using var worksheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), selector);
        SetName(worksheet, newName, oldName, isNewSheet: false, MacAppleEvents.Remaining(timeout, started));
        return new OperationResult { Success = true, FilePath = filePath };
    }

    private static void SetName(MacAppleEvents.Descriptor worksheet, string name, string existingName, bool isNewSheet, TimeSpan timeout)
    {
        using var value = MacAppleEvents.Text(name);
        try
        {
            MacNativeRange.Write(worksheet, MacExcelDictionary.Name, value, timeout);
        }
        catch (MacAppleEventException error) when (error.IsExcelReply)
        {
            throw new MacExcelOperationException("ComInterop",
                WorksheetCommandValidation.NamingRejectionMessage(name, existingName, isNewSheet), error,
                remoteExceptionType: nameof(InvalidOperationException), remoteInnerError: error.Message);
        }
    }

    internal static WorksheetListResult List(string filePath, TimeSpan timeout)
    {
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        using var all = MacAppleEvents.Create(MacAppleEvents.Code("enum"), BitConverter.GetBytes(MacAppleEvents.Code("all ")));
        using var worksheets = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("indx"), all);
        var names = Read(worksheets, MacExcelDictionary.Name, MacAppleEvents.Remaining(timeout, started));
        var visibility = Read(worksheets, MacExcelDictionary.Visible, MacAppleEvents.Remaining(timeout, started));
        var verifiedNames = Read(worksheets, MacExcelDictionary.Name, MacAppleEvents.Remaining(timeout, started));
        if (!JsonNode.DeepEquals(names, verifiedNames))
        {
            throw new InvalidOperationException("The worksheet collection changed during native Excel discovery.");
        }
        return CreateList(filePath, names, visibility);
    }

    private static JsonNode? Read(MacAppleEvents.Descriptor worksheets, uint property, TimeSpan timeout)
    {
        using var target = MacAppleEvents.Property(worksheets, property);
        using var appleEvent = MacAppleEvents.Event(MacAppleEvents.Code("core"), MacAppleEvents.Code("getd"));
        MacAppleEvents.Put(appleEvent, MacAppleEvents.Code("----"), target);
        return MacAppleEvents.Send(appleEvent, timeout);
    }

    internal static WorksheetListResult CreateList(string filePath, JsonNode? names, JsonNode? visibility)
    {
        if (names is not JsonArray nameList || visibility is not JsonArray visibilityList
            || nameList.Count != visibilityList.Count)
        {
            throw new InvalidOperationException("Excel did not return consistent worksheet names and visibility.");
        }
        var result = new WorksheetListResult { FilePath = filePath };
        for (var index = 0; index < nameList.Count; index++)
        {
            var name = nameList[index]?.GetValue<string>();
            if (string.IsNullOrEmpty(name))
            {
                throw new InvalidOperationException("Excel returned an empty worksheet name.");
            }
            var code = visibilityList[index]?.GetValue<uint>()
                ?? throw new InvalidOperationException("Excel returned an empty worksheet visibility.");
            if (code != MacExcelDictionary.SheetVisible && code != MacExcelDictionary.SheetHidden
                && code != MacExcelDictionary.SheetVeryHidden)
            {
                throw new InvalidOperationException($"Excel returned an unknown worksheet visibility code 0x{code:X8}.");
            }
            result.Worksheets.Add(new WorksheetInfo
            {
                Name = name,
                Index = index + 1,
                Visible = code == MacExcelDictionary.SheetVisible
            });
        }
        result.Success = true;
        return result;
    }
}
