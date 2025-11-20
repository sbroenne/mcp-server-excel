namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Represents the interval for grouping date/time fields in PivotTables.
/// </summary>
/// <remarks>
/// These intervals correspond to Excel's date grouping options for PivotTable fields.
/// Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.xlcalculation
/// </remarks>
public enum DateGroupingInterval
{
    /// <summary>
    /// Group by days.
    /// </summary>
    Days = 1,

    /// <summary>
    /// Group by months.
    /// </summary>
    Months = 2,

    /// <summary>
    /// Group by quarters (Q1, Q2, Q3, Q4).
    /// </summary>
    Quarters = 3,

    /// <summary>
    /// Group by years.
    /// </summary>
    Years = 4
}
