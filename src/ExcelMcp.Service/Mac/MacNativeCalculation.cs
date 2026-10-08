using System.Diagnostics;
using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal static class MacNativeCalculation
{
    internal static OperationResult Calculate(string filePath, CalculationScope scope, string sheetName,
        string? rangeAddress, TimeSpan timeout)
    {
        CalculationCommandValidation.Validate(scope, sheetName, rangeAddress, CalculationKind.Normal);
        if (scope == CalculationScope.Application)
            throw new PlatformNotSupportedException("macOS application-scope calculation would affect unrelated workbooks in shared Excel; no calculation was attempted.");
        var started = Stopwatch.GetTimestamp();
        using var workbook = MacNativeWorkbook.Resolve(filePath, timeout);
        using var name = MacAppleEvents.Text(sheetName);
        using var sheet = MacAppleEvents.Object(MacExcelDictionary.WorksheetClass, workbook, MacAppleEvents.Code("name"), name);
        using var range = scope == CalculationScope.Range ? MacNativeRange.Resolve(sheet, rangeAddress!) : null;
        using var command = MacAppleEvents.Event(
            scope == CalculationScope.Range ? MacExcelDictionary.CalculateRangeClass : MacExcelDictionary.CalculateSheetClass,
            scope == CalculationScope.Range ? MacExcelDictionary.CalculateRangeId : MacExcelDictionary.CalculateSheetId);
        MacAppleEvents.Put(command, MacAppleEvents.Code("----"), range ?? sheet);
        MacAppleEvents.SendCommand(command, MacAppleEvents.Remaining(timeout, started));
        return CalculationCommandValidation.CreateResult(filePath, scope, CalculationKind.Normal);
    }
}
