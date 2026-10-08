using Sbroenne.ExcelMcp.Core.Commands.Calculation;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class CalculationCommandValidation
{
    internal static void Validate(CalculationScope scope, string? sheetName, string? rangeAddress, CalculationKind kind)
    {
        if (!Enum.IsDefined(scope))
            throw new ArgumentOutOfRangeException(nameof(scope), scope, $"Unknown calculation scope: {scope}");
        if (!Enum.IsDefined(kind))
            throw new ArgumentOutOfRangeException(nameof(kind));
        if (scope != CalculationScope.Application && kind != CalculationKind.Normal)
            throw new ArgumentException("Full/rebuild calculation requires application scope.", nameof(kind));
        if (scope != CalculationScope.Application && string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("sheetName is required for Sheet/Range scope calculation.", nameof(sheetName));
        if (scope == CalculationScope.Range && string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("rangeAddress is required for Range scope calculation.", nameof(rangeAddress));
        if (scope == CalculationScope.Application && !string.IsNullOrEmpty(sheetName) ||
            scope != CalculationScope.Range && !string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("Sheet/range inputs must match the requested calculation scope.");
    }

    internal static OperationResult CreateResult(string filePath, CalculationScope scope, CalculationKind kind) => new()
    {
        Success = true,
        FilePath = filePath,
        Action = "calculate",
        Message = $"{kind} calculation completed for {scope}; asynchronous refresh/Python completion is not established."
    };
}
